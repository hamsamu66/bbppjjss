// 1. SEMUA REQUIRE WAJIB BERADA DI PALING ATAS
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const { spawn } = require('child_process'); // <-- INI WAJIB ADA DI SINI

// 2. BARU KEMUDIAN DEKLARASI CLASS OCR
class OcrSolver {
    constructor() {
        // Arahkan ke file .exe
        const ocrPath = path.join(process.cwd(), 'bin', 'ocr_server.exe');

        // Mengeksekusi file .exe
        this.pythonProcess = spawn(ocrPath);
        this.pendingRequests = [];

        // Menangkap output hasil tebakan OCR
        this.pythonProcess.stdout.on('data', (data) => {
            const result = data.toString().trim();
            if (this.pendingRequests.length > 0) {
                const resolve = this.pendingRequests.shift();
                resolve(result);
            }
        });

        this.pythonProcess.stderr.on('data', (data) => {
            console.error(`[Error dari AI OCR]: ${data.toString()}`);
        });
    }

    // ... (lanjutkan dengan fungsi solve() dan isi bot lainnya seperti biasa) ...

    // Fungsi untuk memanggil Python dari Node.js
    async solve(base64Image) {
        return new Promise((resolve) => {
            this.pendingRequests.push(resolve);
            // Kirim gambar ditambah newline (\n) agar terbaca oleh perulangan Python
            this.pythonProcess.stdin.write(base64Image + '\n');
        });
    }

    close() {
        this.pythonProcess.stdin.end();
    }
}

// Inisialisasi satu solver agar model ddddocr hanya di-load sekali
const ocr = new OcrSolver();

// ==========================================
// FUNGSI CEK DAN POTONG KUOTA FIREBASE
// ==========================================
async function cekDanPotongKuota(idLisensi) {
    const firebaseUrl = `https://bbppjjss-default-rtdb.asia-southeast1.firebasedatabase.app/lisensi_pengguna/${idLisensi}/kuota.json`;

    try {
        // 1. Ambil data kuota saat ini dari server
        const responCek = await fetch(firebaseUrl);
        let kuotaSaatIni = await responCek.json();

        // Jika lisensi tidak ada atau kuota habis
        if (kuotaSaatIni === null || kuotaSaatIni <= 0) {
            return { izin: false, sisa: 0 };
        }

        // 2. Kurangi kuota sebanyak 1
        const sisaBaru = kuotaSaatIni - 1;
        const responUpdate = await fetch(firebaseUrl, {
            method: 'PUT',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(sisaBaru)
        });

        // 3. Kembalikan status ke bot
        if (responUpdate.ok) {
            return { izin: true, sisa: sisaBaru };
        } else {
            return { izin: false, sisa: 0 };
        }
    } catch (error) {
        console.error(`[!] Gagal memotong kuota: ${error.message}`);
        return { izin: false, sisa: 0 };
    }
}
// ==========================================
// 1. AMBIL VARIABEL DARI ELECTRON
// ==========================================
const ID_LISENSI = process.env.ID_LISENSI || "";
const TARGET_EXCEL = process.env.TARGET_EXCEL || "";
const TEMP_EXCEL_OUT = process.env.TEMP_EXCEL_OUT || "";
const IS_HEADED = process.env.IS_HEADED === 'true';

// ==========================================
// EKSEKUSI UTAMA BOT
// ==========================================
(async () => {
    // 2. CEK DATA
    if (!ID_LISENSI) return console.log("❌ ID LISENSI KOSONG!");
    if (!TARGET_EXCEL || !fs.existsSync(TARGET_EXCEL)) return console.log("❌ FILE EXCEL TIDAK DITEMUKAN!");

    console.log(">>> MENGHUBUNGKAN KE GOOGLE CHROME BAWAAN PC... <<<");

    // 3. LAUNCH GOOGLE CHROME BAWAAN PC (Bukan Chromium Playwright)
    const browser = await chromium.launch({
        headless: !IS_HEADED,
        channel: 'chrome', // <-- MEMAKSA PAKAI CHROME ASLI
        args: ['--start-maximized']
    });

    const context = await browser.newContext({ viewport: null });
    const page = await context.newPage();

    // 4. BACA EXCEL
    const outputFile = TEMP_EXCEL_OUT;
    const inputWb = XLSX.readFile(TARGET_EXCEL);
    const inputData = XLSX.utils.sheet_to_json(inputWb.Sheets[inputWb.SheetNames[0]]);

    let outputData = [];
    if (fs.existsSync(outputFile)) {
        try {
            const outWb = XLSX.readFile(outputFile);
            outputData = XLSX.utils.sheet_to_json(outWb.Sheets[outWb.SheetNames[0]]);
        } catch (e) { outputData = []; }
    }

    const simpanKeExcel = (data) => {
        const ws = XLSX.utils.json_to_sheet(data);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'HASIL');
        XLSX.writeFile(wb, outputFile);
    };

    console.log(`[SYS] EXCEL DIBACA: ${inputData.length} baris data`);

    // ==========================================
    // LOOP UTAMA NIK
    // ==========================================
    for (const row of inputData) {
        const nikTarget = String(row.NIK).trim();
        const sudahSelesai = outputData.find(d => String(d.NIK).trim() === nikTarget && d.STATUS === 'SUKSES');

        if (sudahSelesai) {
            console.log(`[SKIP] NIK ${nikTarget} (Sudah Sukses)`);
            continue;
        }

        // 1. CEK & POTONG KUOTA FIREBASE
        const statusKuota = await cekDanPotongKuota(ID_LISENSI);
        if (!statusKuota.izin) {
            if (statusKuota.sisa === 0) console.log("\n[!] SYSTEM HALT: KUOTA LISENSI ANDA TELAH HABIS!");
            else console.log("\n[!] SYSTEM HALT: KONEKSI SERVER FIREBASE ERROR!");
            break;
        }

        let isDone = false;
        let attempt = 0;

        while (!isDone && attempt < 3) {
            attempt++;
            console.log(`\n--- MEMPROSES NIK: ${nikTarget} (Sisa Kuota: ${statusKuota.sisa}) ---`);

            try {
                await page.goto('https://webskrining.bpjs-kesehatan.go.id/skrining', { waitUntil: 'networkidle', timeout: 20000 });
                await page.waitForSelector('#nik_txt', { state: 'visible', timeout: 10000 });

                // FILL DATA AWAL
                await page.fill('#nik_txt', nikTarget);
                await page.click('#TglLahir_src');
                await page.keyboard.press('Control+A');
                await page.keyboard.press('Backspace');
                await page.locator('#TglLahir_src').pressSequentially(String(row.TGL_LAHIR), { delay: 50 });
                await page.keyboard.press('Enter');

                // 2. LOOP CAPTCHA OTO-AI (Maksimal 10x)
                let captchaLolos = false;
                let capTry = 0;

                while (!captchaLolos && capTry < 10) {
                    capTry++;
                    const capEl = page.locator('#AppCaptcha_CaptchaImage');
                    await capEl.waitFor({ state: 'visible', timeout: 10000 });
                    await page.waitForTimeout(1000);

                    // ====== MENJADI SEPERTI INI ======
                    const captchaBuffer = await capEl.screenshot();
                    const base64Image = captchaBuffer.toString('base64');

                    // Panggil AI OCR (OcrSolver) yang sedang stand-by
                    const kodeOcr = await ocr.solve(base64Image);

                    if (kodeOcr.length < 3) {
                        await page.click('#AppCaptcha_ReloadLink', { force: true });
                        await page.waitForTimeout(1500);
                        continue;
                    }

                    console.log(`[AI] Membaca Captcha: ${kodeOcr} (Try ${capTry}/10)`);
                    const inputCap = page.locator('#captchaCode_txt');
                    await inputCap.click({ force: true });
                    await inputCap.fill('');
                    await inputCap.pressSequentially(kodeOcr, { delay: 100 });
                    await page.click('#btnCariPetugas', { force: true });

                    // BALAPAN RESPONSE (LOGIKA TANGGUH NODE.JS)
                    const raceResult = await Promise.race([
                        page.waitForSelector('.bootbox-body', { state: 'visible', timeout: 15000 }).then(() => 'POPUP'),
                        page.waitForSelector('#beratBadan_txt', { state: 'visible', timeout: 15000 }).then(() => 'FORM'),
                        page.waitForSelector('#hasilSkrJns_Top', { state: 'visible', timeout: 15000 }).then(() => 'HASIL'),
                        page.waitForTimeout(15000).then(() => 'TIMEOUT')
                    ]);

                    if (raceResult === 'HASIL') {
                        await page.waitForTimeout(1000);
                        const tglSkrining = await page.locator('#hasilSkrining_tglSkr').innerText();
                        const namaFktp = await page.locator('#hasilSkrining_nmppk').innerText();

                        console.log(`[OK] SUDAH SKRINING SEBELUMNYA (${tglSkrining.trim()})`);
                        outputData.push({ ...row, STATUS: 'SUKSES', KETERANGAN: 'SUDAH SKRINING SEBELUMNYA', TGL_SKRINING: tglSkrining.trim(), FKTP: namaFktp.trim() });

                        captchaLolos = true;
                        isDone = true;
                    }
                    else if (raceResult === 'FORM') {
                        captchaLolos = true;
                    }
                    else if (raceResult === 'POPUP') {
                        const msg = await page.innerText('.bootbox-body');
                        const btnOk = page.locator('button[data-bb-handler="ok"], .bootbox-accept');

                        // Ubah bagian ini menggunakan dblclick atau clickCount
                        if (await btnOk.isVisible()) {
                            // Menggunakan method double click bawaan:
                            await btnOk.dblclick({ force: true });

                            // Catatan: Jika dblclick() dirasa kurang cocok dengan website-nya, 
                            // Anda bisa memakai alternatif ini:
                            // await btnOk.click({ clickCount: 2, force: true });
                        }

                        if (msg.toLowerCase().includes('captcha')) {
                            console.log("[!] Captcha Salah! Me-reload otomatis...");
                            await page.click('#AppCaptcha_ReloadLink', { force: true });
                            await page.waitForTimeout(1500);
                        }
                        else if (msg.toLowerCase().includes('bukan peserta jkn') || msg.toLowerCase().includes('tanggal lahir')) {
                            console.log(`[!] Gagal: ${msg.trim()}`);
                            outputData.push({ ...row, STATUS: 'GAGAL', KETERANGAN: msg.trim() });
                            captchaLolos = true;
                            isDone = true;
                        }
                        else if (msg.toLowerCase().includes('sadar')) {
                            console.log("[+] Menyetujui Persetujuan Skrining (Sadar)");
                            await page.click('button[data-bb-handler="confirm"]', { force: true });
                            captchaLolos = true;
                        }
                    } else {
                        console.log("[?] Response lambat, refresh captcha...");
                        await page.click('#AppCaptcha_ReloadLink', { force: true });
                    }
                }

                if (!captchaLolos && !isDone) {
                    console.log(`[X] Captcha Gagal 10x untuk NIK ${nikTarget}`);
                    outputData.push({ ...row, STATUS: 'GAGAL', KETERANGAN: 'CAPTCHA GAGAL 10X' });
                    simpanKeExcel(outputData);
                    isDone = true;
                    break;
                }

                if (isDone) {
                    simpanKeExcel(outputData);
                    break;
                }

                // 3. PENGISIAN KUESIONER BARU
                if (await page.locator('#beratBadan_txt').isVisible()) {
                    const defaultBbTb = hitungRataRataBbTb(row.TGL_LAHIR);
                    const targetBB = (row.BB !== undefined && String(row.BB).trim() !== '') ? row.BB : defaultBbTb.bb;
                    const targetTB = (row.TB !== undefined && String(row.TB).trim() !== '') ? row.TB : defaultBbTb.tb;

                    console.log(`[+] Input BB: ${targetBB} | TB: ${targetTB}`);
                    const ketikAman = async (selector, value) => {
                        const locator = page.locator(selector);
                        await locator.click({ force: true });
                        await page.keyboard.press('Control+A');
                        await page.keyboard.press('Backspace');
                        await locator.pressSequentially(String(value), { delay: 10 });
                    };
                    await ketikAman('#beratBadan_txt', targetBB);
                    await ketikAman('#tinggiBadan_txt', targetTB);

                    let formulirSelesai = false;
                    let safety = 0;
                    while (!formulirSelesai && safety < 20) {
                        safety++;
                        await page.waitForTimeout(500);

                        await page.evaluate(() => {
                            const inputs = document.querySelectorAll('input[type="radio"][value="B"]');
                            inputs.forEach(input => {
                                const parent = input.closest('.question');
                                if (parent && window.getComputedStyle(parent).display !== 'none') {
                                    input.click();
                                    if (typeof storedAnswer === 'function') storedAnswer(input.value, input.name);
                                }
                            });
                        });

                        const btnNext = page.locator('#nextGenBtn');
                        if (await btnNext.isVisible()) {
                            const teks = await btnNext.innerText();
                            await btnNext.click({ force: true });
                            if (teks.includes('Simpan')) {
                                await page.waitForTimeout(1000);
                                const konfirmFinal = page.locator('button[data-bb-handler="confirm"], .btn-primary:has-text("OK")').filter({ visible: true });
                                if (await konfirmFinal.count() > 0) {
                                    await konfirmFinal.first().click({ force: true });
                                    formulirSelesai = true;
                                }
                            }
                        } else if (await page.locator('#hasilSkrJns_Top').isVisible()) {
                            formulirSelesai = true;
                        }
                    }

                    await page.waitForSelector('#hasilSkrJns_Top', { state: 'visible', timeout: 15000 });
                    await page.waitForTimeout(1000);
                    const tglSkriningAkhir = await page.locator('#hasilSkrining_tglSkr').innerText();
                    const namaFktpAkhir = await page.locator('#hasilSkrining_nmppk').innerText();

                    console.log(`[OK] DATA BARU DISIMPAN (${tglSkriningAkhir.trim()} - ${namaFktpAkhir.trim()})`);
                    outputData.push({ ...row, BB: targetBB, TB: targetTB, STATUS: 'SUKSES', KETERANGAN: 'BARU SELESAI', TGL_SKRINING: tglSkriningAkhir.trim(), FKTP: namaFktpAkhir.trim() });

                    simpanKeExcel(outputData);
                    isDone = true;
                }

            } catch (err) {
                console.log(`[?] Error Navigasi (${err.message.split('\n')[0]}). Reload (Try ${attempt}/3)...`);
                await page.waitForTimeout(2000);
            }
        }

        if (!isDone) {
            console.log(`[X] GAGAL TOTAL NIK ${nikTarget} (Server Error)`);
            outputData.push({ ...row, STATUS: 'GAGAL', KETERANGAN: 'GAGAL SETELAH 3X PERCOBAAN / SERVER MACET' });
            simpanKeExcel(outputData);
        }
    }

    // ====== MENJADI SEPERTI INI ======
    console.log("[=] BATCH PROCESSING SELESAI [=]");
    ocr.close(); // Wajib ditambahkan agar proses Python mati setelah selesai
    await browser.close();

})().catch(err => {
    console.error(`[FATAL ERROR]: ${err.message}`);
});