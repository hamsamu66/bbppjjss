# Isi dari bot_logic.py (Simpan ini di GitHub)
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import os
import time
import requests
import pandas as pd
import threading
import ddddocr
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

FIREBASE_URL = "https://bbppjjss-default-rtdb.asia-southeast1.firebasedatabase.app/lisensi_pengguna" # GANTI DENGAN URL FIREBASE KAMU

# ==========================================
# KONFIGURASI TEMA HACKER / RETRO
# ==========================================
BG_COLOR = "#000000"       # Hitam Pekat
FG_COLOR = "#00FF00"       # Hijau Neon / Lime
FONT_STYLE = ("Courier New", 10, "bold")
BTN_STYLE = {
    "bg": BG_COLOR, 
    "fg": FG_COLOR, 
    "font": FONT_STYLE, 
    "relief": "solid", 
    "bd": 1,
    "activebackground": FG_COLOR,
    "activeforeground": BG_COLOR
}

# ==========================================
# MODUL 1: LOGIKA FIREBASE (LOGIN & KUOTA)
# ==========================================
class FirebaseAuth:
    @staticmethod
    def verifikasi_login(id_user):
        try:
            res = requests.get(f"{FIREBASE_URL}/{id_user}.json", timeout=10)
            data = res.json()
            if not data:
                return False, "ID LISENSI TIDAK TERDAFTAR!"
            if data.get('kuota', 0) <= 0:
                return False, "KUOTA HABIS! SILAKAN ISI ULANG."
            return True, data['kuota']
        except Exception as e:
            return False, f"GAGAL KONEKSI SERVER: {e}"

    @staticmethod
    def cek_dan_potong_kuota(id_user):
        try:
            res = requests.get(f"{FIREBASE_URL}/{id_user}/kuota.json", timeout=5)
            kuota_sekarang = int(res.json())
            if kuota_sekarang > 0:
                kuota_baru = kuota_sekarang - 1
                requests.patch(f"{FIREBASE_URL}/{id_user}.json", json={"kuota": kuota_baru})
                return True, kuota_baru
            else:
                return False, 0
        except Exception as e:
            return False, -1

# ==========================================
# MODUL 2: LOGIKA BOT PLAYWRIGHT
# ==========================================
def hitung_rata_rata_bbtb(tgl_lahir_str):
    try:
        parts = str(tgl_lahir_str).strip().replace('/', '-').split('-')
        tahun = int(parts[0]) if len(parts[0]) == 4 else int(parts[2])
        umur = 2026 - tahun
        if umur < 18: return 50, 155
        elif 18 <= umur <= 59: return 60, 165
        else: return 55, 160
    except:
        return 60, 165

def jalankan_bot(file_excel, id_user, log_callback, stats_callback, on_finish_callback):
    ocr = ddddocr.DdddOcr(show_ad=False)
    
    try:
        df_input = pd.read_excel(file_excel)
        input_data = df_input.to_dict('records')
    except Exception as e:
        log_callback(f"[!] GAGAL MEMBACA EXCEL: {e}")
        on_finish_callback()
        return

    output_file = file_excel.replace(".xlsx", "_HASIL.xlsx")
    output_data = []
    
    if os.path.exists(output_file):
        try:
            df_out = pd.read_excel(output_file)
            output_data = df_out.to_dict('records')
        except:
            pass

    def simpan_excel():
        pd.DataFrame(output_data).to_excel(output_file, index=False)

    with sync_playwright() as p:
        browser = p.chromium.launch(channel="chrome", headless=False)
        page = browser.new_page()
        page.set_default_timeout(15000)

        for row in input_data:
            nik_target = str(row['NIK']).strip()
            
            sudah_diproses = any(str(d.get('NIK', '')).strip() == nik_target for d in output_data)
            if sudah_diproses:
                log_callback(f"[>] SKIP NIK {nik_target} (SUDAH ADA DI HASIL)")
                continue

            # POTONG KUOTA DI AWAL SEBELUM PROSES
            bisa_proses, sisa_kuota_live = FirebaseAuth.cek_dan_potong_kuota(id_user)
            
            if not bisa_proses:
                if sisa_kuota_live == 0:
                    log_callback("[!] SYSTEM HALT: KUOTA LISENSI HABIS!")
                else:
                    log_callback("[!] KONEKSI SERVER ERROR. BOT DIHENTIKAN SEMENTARA.")
                break 

            log_callback(f"\n[*] MEMPROSES NIK: {nik_target} (SISA KUOTA BERSAMA: {sisa_kuota_live})")
            is_nik_done = False
            attempt = 0

            while not is_nik_done and attempt < 3:
                attempt += 1
                try:
                    page.goto('https://webskrining.bpjs-kesehatan.go.id/skrining', wait_until="load", timeout=15000)
                    page.wait_for_selector('#nik_txt', state='visible', timeout=10000)

                    page.fill('#nik_txt', nik_target)
                    page.click('#TglLahir_src')
                    page.keyboard.press('Control+A')
                    page.keyboard.press('Backspace')
                    page.locator('#TglLahir_src').press_sequentially(str(row['TGL_LAHIR']), delay=50)
                    page.keyboard.press('Enter')

                    captcha_lolos = False
                    cap_try = 0

                    while not captcha_lolos and cap_try < 10:
                        cap_try += 1
                        cap_el = page.locator('#AppCaptcha_CaptchaImage')
                        cap_el.wait_for(state='visible', timeout=10000)
                        time.sleep(1)

                        img_bytes = cap_el.screenshot()
                        kode_ocr = ocr.classification(img_bytes).upper()

                        if len(kode_ocr) < 3:
                            page.click('#AppCaptcha_ReloadLink')
                            time.sleep(1.5)
                            continue

                        log_callback(f"[+] OCR DECODE: {kode_ocr}")
                        
                        input_cap = page.locator('#captchaCode_txt')
                        input_cap.click()
                        input_cap.fill('')
                        input_cap.press_sequentially(kode_ocr, delay=100)
                        page.click('#btnCariPetugas')

                        try:
                            page.wait_for_selector(".bootbox-body, #beratBadan_txt, #hasilSkrJns_Top", state="visible", timeout=15000)
                            if page.locator("#hasilSkrJns_Top").is_visible():
                                race_result = 'HASIL'
                            elif page.locator("#beratBadan_txt").is_visible():
                                race_result = 'FORM'
                            else:
                                race_result = 'POPUP'
                        except PlaywrightTimeoutError:
                            race_result = 'TIMEOUT'

                        if race_result == 'HASIL':
                            time.sleep(1)
                            tgl = page.locator('#hasilSkrining_tglSkr').inner_text().strip()
                            fktp = page.locator('#hasilSkrining_nmppk').inner_text().strip()
                            log_callback(f"[OK] SUDAH SKRINING ({tgl} - {fktp})")
                            
                            row.update({'STATUS': 'SUKSES', 'KETERANGAN': 'SUDAH SKRINING SEBELUMNYA', 'TGL_SKRINING': tgl, 'FKTP': fktp})
                            output_data.append(row)
                            stats_callback('SUKSES') # Update Tracker GUI
                            
                            captcha_lolos = True
                            is_nik_done = True
                            
                        elif race_result == 'FORM':
                            captcha_lolos = True
                            
                        elif race_result == 'POPUP':
                            msg = page.inner_text('.bootbox-body').lower()
                            if page.locator('button[data-bb-handler="ok"]').is_visible():
                                page.click('button[data-bb-handler="ok"]')
                                
                            if 'captcha' in msg:
                                page.click('#AppCaptcha_ReloadLink')
                                time.sleep(1.5)
                            elif 'bukan peserta' in msg or 'tanggal lahir' in msg:
                                row.update({'STATUS': 'GAGAL', 'KETERANGAN': msg.upper()})
                                output_data.append(row)
                                stats_callback('GAGAL') # Update Tracker GUI
                                
                                captcha_lolos = True
                                is_nik_done = True
                            elif 'sadar' in msg:
                                page.click('button[data-bb-handler="confirm"]')
                                captcha_lolos = True
                        else:
                            page.click('#AppCaptcha_ReloadLink')

                    if not captcha_lolos and not is_nik_done:
                        row.update({'STATUS': 'GAGAL', 'KETERANGAN': 'CAPTCHA GAGAL 10X'})
                        output_data.append(row)
                        stats_callback('GAGAL') # Update Tracker GUI
                        is_nik_done = True
                        break

                    if is_nik_done:
                        simpan_excel()
                        break

                    # ISI FORM (DATA BARU)
                    if page.locator('#beratBadan_txt').is_visible():
                        def_bb, def_tb = hitung_rata_rata_bbtb(row['TGL_LAHIR'])
                        target_bb = row.get('BB') if pd.notna(row.get('BB')) else def_bb
                        target_tb = row.get('TB') if pd.notna(row.get('TB')) else def_tb

                        page.locator('#beratBadan_txt').fill('')
                        page.locator('#beratBadan_txt').press_sequentially(str(target_bb), delay=10)
                        page.locator('#tinggiBadan_txt').fill('')
                        page.locator('#tinggiBadan_txt').press_sequentially(str(target_tb), delay=10)
                        
                        log_callback(f"[>] INJECT BB:{target_bb} TB:{target_tb}")

                        page.evaluate("""
                            const inputs = document.querySelectorAll('input[type="radio"][value="B"]');
                            inputs.forEach(i => { if (i.closest('.question').style.display !== 'none') i.click(); });
                        """)

                        if page.locator('#nextGenBtn').is_visible():
                            page.click('#nextGenBtn', force=True)
                            time.sleep(1)
                            if page.locator('button[data-bb-handler="confirm"]').is_visible():
                                page.click('button[data-bb-handler="confirm"]')

                        page.wait_for_selector('#hasilSkrJns_Top', state='visible', timeout=15000)
                        time.sleep(1)
                        tgl = page.locator('#hasilSkrining_tglSkr').inner_text().strip()
                        fktp = page.locator('#hasilSkrining_nmppk').inner_text().strip()

                        log_callback(f"[OK] DATA DISIMPAN ({tgl} - {fktp})")
                        row.update({'BB': target_bb, 'TB': target_tb, 'STATUS': 'SUKSES', 'KETERANGAN': 'BARU SELESAI', 'TGL_SKRINING': tgl, 'FKTP': fktp})
                        output_data.append(row)
                        stats_callback('SUKSES') # Update Tracker GUI
                        
                        simpan_excel()
                        is_nik_done = True

                except Exception as e:
                    log_callback(f"[!] SYSTEM TIMEOUT (ATTEMPT {attempt}/3). RETRYING...")
                    time.sleep(2)

            if not is_nik_done:
                log_callback(f"[!] NIK {nik_target} SKIPPED (SERVER UNREACHABLE 3X)")
                row.update({'STATUS': 'GAGAL', 'KETERANGAN': 'GAGAL 3X (SERVER MACET)'})
                output_data.append(row)
                stats_callback('GAGAL') # Update Tracker GUI
                simpan_excel()
                
        browser.close()
    
    log_callback("\n[=] BATCH PROCESSING COMPLETED [=]")
    on_finish_callback()

# ==========================================
# MODUL 3: GUI TKINTER (RETRO THEME)
# ==========================================
class AppGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Aplikasi Penghitung Pengeluaran selama 10 tahun")
        self.root.geometry("600x550")
        self.root.configure(bg=BG_COLOR)
        
        # Variabel State
        self.id_aktif = None
        self.file_target = None
        self.count_sukses = 0
        self.count_gagal = 0
        
        # Frame Login
        self.frame_login = tk.Frame(root, bg=BG_COLOR)
        self.frame_login.pack(pady=50)
        
        tk.Label(self.frame_login, text="[ AUTHENTICATION REQUIRED ]", bg=BG_COLOR, fg=FG_COLOR, font=("Courier New", 14, "bold")).pack(pady=10)
        tk.Label(self.frame_login, text="ENTER LICENSE ID:", bg=BG_COLOR, fg=FG_COLOR, font=FONT_STYLE).pack(pady=5)
        
        # Entry khusus tema hacker
        self.entry_id = tk.Entry(self.frame_login, width=30, justify="center", bg=BG_COLOR, fg=FG_COLOR, insertbackground=FG_COLOR, font=("Courier New", 12, "bold"), relief="solid", bd=1)
        self.entry_id.pack(pady=10)
        
        self.btn_login = tk.Button(self.frame_login, text="CONNECT TO SERVER", command=self.proses_login, **BTN_STYLE)
        self.btn_login.pack(pady=10)

        # Frame Kontrol (Sembunyi sebelum login)
        self.frame_kontrol = tk.Frame(root, bg=BG_COLOR)
        
        self.lbl_info = tk.Label(self.frame_kontrol, text="", bg=BG_COLOR, fg=FG_COLOR, font=FONT_STYLE)
        self.lbl_info.pack(pady=5)

        frame_btn = tk.Frame(self.frame_kontrol, bg=BG_COLOR)
        frame_btn.pack(pady=10)
        
        tk.Button(frame_btn, text="[1] DOWNLOAD TEMPLATE", command=self.unduh_template, **BTN_STYLE).grid(row=0, column=0, padx=5)
        tk.Button(frame_btn, text="[2] MOUNT EXCEL FILE", command=self.upload_file, **BTN_STYLE).grid(row=0, column=1, padx=5)
        
        self.lbl_file = tk.Label(self.frame_kontrol, text="NO FILE MOUNTED", bg=BG_COLOR, fg="red", font=FONT_STYLE)
        self.lbl_file.pack()
        
        self.btn_mulai = tk.Button(self.frame_kontrol, text=">> EXECUTE PAYLOAD <<", state="disabled", command=self.mulai_bot, bg=BG_COLOR, fg=FG_COLOR, font=("Courier New", 12, "bold"), relief="solid", bd=2, activebackground=FG_COLOR, activeforeground=BG_COLOR)
        self.btn_mulai.pack(pady=10)

        # LIVE TRACKER STATISTIK
        self.lbl_stats = tk.Label(self.frame_kontrol, text="[ BERHASIL: 0  |  GAGAL: 0 ]", bg=BG_COLOR, fg="yellow", font=("Courier New", 12, "bold"))
        self.lbl_stats.pack(pady=5)

        # Terminal Text Box
        self.log_box = scrolledtext.ScrolledText(self.frame_kontrol, width=65, height=13, state='disabled', bg=BG_COLOR, fg=FG_COLOR, font=("Courier New", 9), relief="solid", bd=1)
        self.log_box.pack(pady=5)

    def print_log(self, text):
        # Aman dipanggil dari thread lain
        self.root.after(0, self._print_log_safe, text)

    def _print_log_safe(self, text):
        self.log_box.config(state='normal')
        self.log_box.insert(tk.END, text + "\n")
        self.log_box.see(tk.END)
        self.log_box.config(state='disabled')

    def update_stats(self, status):
        self.root.after(0, self._update_stats_safe, status)
        
    def _update_stats_safe(self, status):
        if status == 'SUKSES':
            self.count_sukses += 1
        elif status == 'GAGAL':
            self.count_gagal += 1
        self.lbl_stats.config(text=f"[ BERHASIL: {self.count_sukses}  |  GAGAL: {self.count_gagal} ]")

    def proses_login(self):
        id_input = self.entry_id.get().strip()
        if not id_input: return
        
        self.btn_login.config(text="AUTHENTICATING...", state="disabled")
        self.root.update()
        
        sukses, pesan = FirebaseAuth.verifikasi_login(id_input)
        
        if sukses:
            self.id_aktif = id_input
            self.frame_login.pack_forget()
            self.frame_kontrol.pack(fill="both", expand=True)
            # ID dihilangkan, diganti dengan teks status yang lebih keren ala hacker
            self.lbl_info.config(text=f"STATUS: AUTHORIZED | ACTIVE QUOTA: {pesan}") 
            self.print_log(">>> SECURE CONNECTION ESTABLISHED <<<")
        else:
            messagebox.showerror("ACCESS DENIED", pesan)
            self.btn_login.config(text="CONNECT TO SERVER", state="normal")

    def unduh_template(self):
        lokasi = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="Template_Skrining.xlsx")
        if lokasi:
            df = pd.DataFrame(columns=["NIK", "TGL_LAHIR", "BB", "TB"])
            df.to_excel(lokasi, index=False)
            messagebox.showinfo("OK", "TEMPLATE EXCEL EXPORTED.")

    def upload_file(self):
        file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file:
            self.file_target = file
            self.lbl_file.config(text=f"MOUNTED: {os.path.basename(file)}", fg=FG_COLOR)
            self.btn_mulai.config(state="normal")

    def mulai_bot(self):
        self.btn_mulai.config(state="disabled", text="PROCESSING PAYLOAD...")
        # Reset counter jika start ulang file baru
        self.count_sukses = 0
        self.count_gagal = 0
        self.lbl_stats.config(text="[ BERHASIL: 0  |  GAGAL: 0 ]")
        
        # Jalankan di Thread terpisah agar GUI tetap responsif
        threading.Thread(target=jalankan_bot, args=(self.file_target, self.id_aktif, self.print_log, self.update_stats, self.selesai_bot), daemon=True).start()

    def selesai_bot(self):
        self.root.after(0, self._selesai_bot_safe)

    def _selesai_bot_safe(self):
        self.btn_mulai.config(state="normal", text=">> EXECUTE PAYLOAD <<")
        messagebox.showinfo("SYSTEM MESSAGE", "BATCH PROCESSING COMPLETED.\nDATA SAVED AS _HASIL.xlsx")

# ==========================================
# ENTRY POINT LAUNCHER
# ==========================================
def jalankan_aplikasi():
    root = tk.Tk()
    app = AppGUI(root)
    root.mainloop()

# Jika script dijalankan langsung (bukan lewat launcher), tetap bisa jalan untuk testing
if __name__ == "__main__":
    jalankan_aplikasi()