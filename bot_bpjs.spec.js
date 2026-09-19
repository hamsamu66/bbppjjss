const { test } = require('@playwright/test');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

test.use({ channel: 'chrome' });
// ==========================================
// KONFIGURASI FIREBASE & VARIABEL ELECTRON
// ==========================================
const FIREBASE_URL = "https://bbppjjss-default-rtdb.asia-southeast1.firebasedatabase.app/lisensi_pengguna"; // <-- GANTI URL INI
const ID_LISENSI = process.env.LISENSI_BPJS || "";
const TARGET_EXCEL = process.env.TARGET_EXCEL || "";

// ==========================================
// FUNGSI CEK KUOTA (FIREBASE)
// ==========================================
async function cekDanPotongKuota(idUser) {
    try {
        const response = await fetch(`${FIREBASE_URL}/${idUser}/kuota.json`);
        const kuotaSekarang = await response.json();

        if (kuotaSekarang > 0) {
            const kuotaBaru = kuotaSekarang - 1;
            await fetch(`${FIREBASE_URL}/${idUser}.json`, {
                method: 'PATCH',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ kuota: kuotaBaru })
            });
            return { izin: true, sisa: kuotaBaru };
        }
        return { izin: false, sisa: 0 };
    } catch (error) {
        console.log(`[ERR] Gagal koneksi Firebase: ${error.message}`);
        return { izin: false, sisa: -1 };
    }
}

// ==========================================
// FUNGSI JEMBATAN AI (ddddocr Python)
// ==========================================
function bacaCaptchaDdddOcr(imgBuffer) {
    try {
        const base64Image = imgBuffer.toString('base64');
        // Panggil script ocr_server.py[cite: 5]
        const result = execSync('python ocr_server.py', {
            input: base64Image,
            encoding: 'utf-8'
        }).trim();
        return result.toUpperCase();
    } catch (error) {
        console.log(`[ERR] Python ddddocr gagal: ${error.message}`);
        return "";
    }
}

// ==========================================
// FUNGSI KALKULATOR BB & TB
// ==========================================
function hitungRataRataBbTb(tglLahirStr) {
    try {
        let parts = String(tglLahirStr).trim().split(/[-/]/);
        let tahun = parts.length === 3 ? (parts[0].length === 4 ? parseInt(parts[0]) : parseInt(parts[2])) : 2000;
        let umur = 2026 - tahun;
        if (umur < 18) return { bb: 50, tb: 155 };
        else if (umur >= 18 && umur <= 59) return { bb: 60, tb: 165 };
        else return { bb: 55, tb: 160 };
    } catch (e) {
        return { bb: 60, tb: 165 };
    }
}

// ==========================================
// SKRIP OTOMATISASI PLAYWRIGHT
// ==========================================
test('Skrining BPJS - Electron Payload', async ({ page }) => {
    test.setTimeout(0); // Matikan limit waktu[cite: 3]

    if (!ID_LISENSI) return console.error("❌ ID LISENSI KOSONG!");
    if (!TARGET_EXCEL || !fs.existsSync(TARGET_EXCEL)) return console.error("❌ FILE EXCEL TIDAK DITEMUKAN!");

    // Buat nama file output dinamis di folder yang sama dengan file aslinya
    // Ambil folder Temp rahasia dari Electron
    const TEMP_EXCEL_OUT = process.env.TEMP_EXCEL_OUT;
    const TARGET_EXCEL = process.env.TARGET_EXCEL || "";

    // Output File TIDAK LAGI diletakkan di sebelah file asli, melainkan di Temp Folder
    const outputFile = TEMP_EXCEL_OUT;

    // Baca Excel
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

                // FILL DATA AWAL[cite: 3]
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
                    await page.waitForTimeout(1000); // Tunggu render

                    const imgPath = path.join(__dirname, `temp_${nikTarget}.png`);
                    await capEl.screenshot({ path: imgPath });

                    // Panggil AI OCR
                    const kodeOcr = bacaCaptchaDdddOcr(fs.readFileSync(imgPath));
                    if (fs.existsSync(imgPath)) fs.unlinkSync(imgPath);

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

                    // BALAPAN RESPONSE (LOGIKA TANGGUH NODE.JS)[cite: 1]
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
                        if (await btnOk.isVisible()) await btnOk.click({ force: true });

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

                // Jika Captcha buntu 10x
                if (!captchaLolos && !isDone) {
                    console.log(`[X] Captcha Gagal 10x untuk NIK ${nikTarget}`);
                    outputData.push({ ...row, STATUS: 'GAGAL', KETERANGAN: 'CAPTCHA GAGAL 10X' });
                    simpanKeExcel(outputData);
                    isDone = true;
                    break;
                }

                if (isDone) {
                    simpanKeExcel(outputData);
                    break; // Pindah NIK
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

                        // KLIK OTOMATIS OPSI 'TIDAK' (Value B)
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

                    // TUNGGU HALAMAN HASIL AKHIR
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
});