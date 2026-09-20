// ==========================================
// 0. SUNTIKAN GPS NODE_MODULES (WAJIB PALING ATAS)
// ==========================================
if (process.env.NODE_MODULES_PATH) {
    require('module').globalPaths.push(process.env.NODE_MODULES_PATH);
    module.paths.unshift(process.env.NODE_MODULES_PATH);
}

const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

// ==========================================
// FUNGSI BANTUAN
// ==========================================
async function cekDanPotongKuota(id) {
    return { izin: true, sisa: 99 };
}

function bacaCaptchaDdddOcr(bufferImg) {
    return "1234";
}

function hitungRataRataBbTb(tglLahir) {
    return { bb: 60, tb: 160 };
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

    // 3. LAUNCH GOOGLE CHROME BAWAAN PC
    const browser = await chromium.launch({
        headless: !IS_HEADED,
        channel: 'chrome',
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

                await page.fill('#nik_txt', nikTarget);
                await page.click('#TglLahir_src');
                await page.keyboard.press('Control+A');
                await page.keyboard.press('Backspace');
                await page.locator('#TglLahir_src').pressSequentially(String(row.TGL_LAHIR), { delay: 50 });
                await page.keyboard.press('Enter');

                let captchaLolos = false;
                let capTry = 0;

                while (!captchaLolos && capTry < 10) {
                    capTry++;
                    const capEl = page.locator('#AppCaptcha_CaptchaImage');
                    await capEl.waitFor({ state: 'visible', timeout: 10000 });
                    await page.waitForTimeout(1000);

                    const imgPath = path.join(__dirname, `temp_${nikTarget}.png`);
                    await capEl.screenshot({ path: imgPath });

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

    console.log("[=] BATCH PROCESSING SELESAI [=]");
    await browser.close();

})().catch(err => {
    console.error(`[FATAL ERROR]: ${err.message}`);
});