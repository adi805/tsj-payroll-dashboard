#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Kit Upah TSJ — Dashboard (100% Python)
GUI wrapper (tkinter) untuk pipeline_upah.py + auto-update dari GitHub repo.
"""
import os, sys, json, queue, runpy, shutil, threading, time, traceback, datetime, zipfile

APP_VERSION = '1.0.5'
APP_TITLE = 'Kit Upah TSJ — Dashboard'

GH_REPO = 'adi805/tsj-payroll-dashboard'
GH_API = 'https://api.github.com/repos/' + GH_REPO + '/releases/latest'
GH_TOKEN = 'UPDATE_ME'   # token read-only khusus repo; ganti via menu kalau di-rotate

# Import openpyxl + submodul yang dipakai pipeline_upah.py.
# PENTING: pipeline_upah.py di-bundle sebagai DATA (bukan dianalisis PyInstaller),
# jadi semua dependensinya harus di-import di sini supaya ikut ter-freeze ke .exe.
import openpyxl  # noqa: F401
from openpyxl import Workbook  # noqa: F401
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side  # noqa: F401
from openpyxl.utils import get_column_letter  # noqa: F401
from openpyxl.worksheet.datavalidation import DataValidation  # noqa: F401

import tkinter as tk
from tkinter import ttk, filedialog, messagebox


def app_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))

APP_DIR = app_dir()
LOG_DIR = os.path.join(APP_DIR, 'logs')
UPDATE_DIR = os.path.join(APP_DIR, '_update')


def pipeline_path():
    ext = os.path.join(APP_DIR, 'pipeline_upah.py')
    if os.path.isfile(ext):
        return ext
    if getattr(sys, 'frozen', False):
        bundled = os.path.join(getattr(sys, '_MEIPASS', APP_DIR), 'pipeline_upah.py')
        if os.path.isfile(bundled):
            try:
                with open(bundled, 'rb') as f: data = f.read()
                with open(ext, 'wb') as f: f.write(data)
                return ext
            except Exception as e:
                sys.stderr.write('Gagal extract bundled pipeline_upah.py: %s\n' % e)
    return None


def find_default_base():
    cand = os.path.join(APP_DIR, 'closingan-juli-2026')
    if os.path.isdir(cand):
        return cand
    try:
        for name in sorted(os.listdir(APP_DIR)):
            full = os.path.join(APP_DIR, name)
            if name.lower().startswith('closingan') and os.path.isdir(full):
                return full
    except Exception:
        pass
    return ''


def default_output():
    return os.path.join(APP_DIR, 'Kit-Upah-TSJ-AUTO.xlsx')


def http_get(url, headers=None, timeout=30):
    import urllib.request
    req = urllib.request.Request(url, headers=headers or {})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.status, r.read()


def http_download(url, dest, headers=None, timeout=120, progress=None):
    import urllib.request
    req = urllib.request.Request(url, headers=headers or {})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        total = int(r.headers.get('Content-Length') or 0)
        done = 0
        with open(dest, 'wb') as f:
            while True:
                chunk = r.read(65536)
                if not chunk: break
                f.write(chunk); done += len(chunk)
                if progress and total: progress(done, total)


UPDATER_BAT = r'''@echo off
rem Kit Upah TSJ - apply update (dipanggil otomatis oleh aplikasi)
cd /d "%~dp0"
echo Menunggu aplikasi ditutup...
taskkill /IM KitUpahTSJ.exe /F >nul 2>&1
timeout /t 2 /nobreak >nul
if exist _update (
  echo Menerapkan file baru...
  copy /Y "_update\*" "." >nul
  rmdir /S /Q _update
  echo Update selesai.
) else (
  echo Folder _update tidak ditemukan - update dibatalkan.
)
echo.
pause
'''


def gh_headers():
    h = {'Accept': 'application/vnd.github+json',
         'User-Agent': 'KitUpahTSJ-Dashboard/' + APP_VERSION}
    if GH_TOKEN and GH_TOKEN != 'UPDATE_ME':
        h['Authorization'] = 'Bearer ' + GH_TOKEN
    return h


def check_latest_release(log=None):
    def _log(m):
        if log: log(m)
    _log('Cek release terbaru di github.com/%s ...' % GH_REPO)
    status, body = http_get(GH_API, headers=gh_headers(), timeout=20)
    if status != 200:
        raise RuntimeError('GitHub API balas HTTP %s' % status)
    rel = json.loads(body.decode('utf-8'))
    tag = rel.get('tag_name') or ''
    assets = rel.get('assets') or []
    zips = [a for a in assets if str(a.get('name','')).lower().endswith('.zip')]
    if not zips:
        raise RuntimeError('Release %s tidak punya asset zip' % tag)
    return tag, zips[0]


def newer_than_remote(tag):
    def parse(v):
        v = str(v).lstrip('vV')
        try: return tuple(int(x) for x in v.split('.'))
        except Exception: return (0,)
    return parse(tag) > parse(APP_VERSION)


def apply_update_stage():
    bat_path = os.path.join(APP_DIR, 'updater.bat')
    with open(bat_path, 'w', encoding='ascii', errors='replace', newline='\r\n') as f:
        f.write(UPDATER_BAT)
    if sys.platform == 'win32':
        os.startfile(bat_path)
    else:
        for name in os.listdir(UPDATE_DIR):
            src = os.path.join(UPDATE_DIR, name); dst = os.path.join(APP_DIR, name)
            try:
                if os.path.isfile(dst): os.remove(dst)
                shutil.move(src, dst)
            except Exception as e:
                raise RuntimeError('Gagal swap %s: %s' % (name, e))
        shutil.rmtree(UPDATE_DIR, ignore_errors=True)


def run_update_flow(log=None):
    def _log(m):
        if log: log(m)
    tag, asset = check_latest_release(_log)
    _log('Release terbaru: %s (versi berjalan: v%s)' % (tag, APP_VERSION))
    if not newer_than_remote(tag):
        return 'Sudah versi terbaru (v%s = %s). Tidak perlu update.' % (APP_VERSION, tag)
    os.makedirs(UPDATE_DIR, exist_ok=True)
    zip_path = os.path.join(UPDATE_DIR, 'release.zip')
    url = asset.get('url')
    headers = dict(gh_headers()); headers['Accept'] = 'application/octet-stream'
    size_mb = (asset.get('size') or 0) / 1048576.0
    _log('Download %s (%.1f MB)...' % (asset.get('name'), size_mb))
    http_download(url, zip_path, headers=headers, timeout=300)
    _log('Download selesai: %d bytes' % os.path.getsize(zip_path))
    with zipfile.ZipFile(zip_path) as z:
        names = z.namelist(); z.extractall(UPDATE_DIR)
    os.remove(zip_path)
    _log('Isi update: %s' % ', '.join(os.path.basename(n) for n in names if not n.endswith('/')))
    if not any(n.lower().endswith('.exe') for n in names):
        raise RuntimeError('Zip release tidak berisi .exe — update dibatalkan.')
    apply_update_stage()
    return ('Update %s sudah di-download. Setelah aplikasi ini ditutup, '
            'updater.bat akan mengganti file lama otomatis. '
            'Buka lagi aplikasinya setelah selesai.' % tag)


class QWriter:
    def __init__(self, q): self.q = q; self.buf = ''
    def write(self, s):
        self.buf += s
        while '\n' in self.buf:
            line, self.buf = self.buf.split('\n', 1)
            self.q.put(('log', line))
    def flush(self):
        if self.buf: self.q.put(('log', self.buf)); self.buf = ''


def run_pipeline_worker(base, out, q):
    pp = pipeline_path(); t0 = time.time()
    if not pp:
        q.put(('log', 'ERROR: pipeline_upah.py tidak ditemukan.'))
        q.put(('log', 'Taruh pipeline_upah.py di folder yang sama dengan aplikasi ini.'))
        q.put(('done', False, 0.0)); return
    if not os.path.isdir(base):
        q.put(('log', 'ERROR: folder sumber tidak ditemukan: %s' % base))
        q.put(('done', False, 0.0)); return
    old_argv = sys.argv[:]; old_out, old_err = sys.stdout, sys.stderr
    w = QWriter(q); sys.stdout = w; sys.stderr = w; ok = False
    try:
        q.put(('log', '='*60))
        q.put(('log', 'MULAI PROSES — %s' % datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
        q.put(('log', 'Pipeline : %s' % pp)); q.put(('log', 'Folder   : %s' % base))
        q.put(('log', 'Output   : %s' % out)); q.put(('log', '='*60))
        sys.argv = [pp, base, out]
        runpy.run_path(pp, run_name='__main__'); ok = True
    except SystemExit as e:
        ok = e.code in (0, None)
        if not ok: q.put(('log', 'Pipeline berhenti dengan kode: %s' % e.code))
    except Exception:
        traceback.print_exc(); ok = False
    finally:
        w.flush(); sys.argv = old_argv; sys.stdout, sys.stderr = old_out, old_err
        q.put(('done', ok, time.time() - t0))


def update_worker(q):
    t0 = time.time()
    try:
        def log(m): q.put(('log', '[update] ' + m))
        msg = run_update_flow(log)
        q.put(('log', '[update] ' + msg)); q.put(('update-done', True, msg, time.time()-t0))
    except Exception as e:
        q.put(('log', '[update] ERROR: %s' % e)); q.put(('update-done', False, str(e), time.time()-t0))


def selftest():
    base = sys.argv[2] if len(sys.argv) > 2 else find_default_base()
    out = sys.argv[3] if len(sys.argv) > 3 else default_output()
    q = queue.Queue()
    t = threading.Thread(target=run_pipeline_worker, args=(base, out, q), daemon=True); t.start()
    ok_final = False
    while True:
        item = q.get(); kind = item[0]
        if kind == 'log': print(item[1])
        else:
            ok_final, elapsed = item[1], item[2]
            print('[selftest] done ok=%s elapsed=%.1fs' % (ok_final, elapsed)); break
    sys.exit(0 if ok_final else 1)


class App:
    def __init__(self, root):
        self.root = root; self.q = queue.Queue(); self.running = False
        self.t_start = 0.0; self.last_out = ''; self.log_file = None
        root.title('%s v%s' % (APP_TITLE, APP_VERSION))
        root.geometry('960x660'); root.minsize(780, 540)
        style = ttk.Style(root)
        try: style.theme_use('clam')
        except Exception: pass
        style.configure('Big.TButton', font=('Segoe UI', 12, 'bold'), padding=8, background='#1e7d32', foreground='white')
        style.configure('Status.TLabel', padding=(6, 3))
        self._build_header(); self._build_menu(); self._build_form(); self._build_log(); self._build_statusbar()
        self.var_base.set(find_default_base()); self.var_out.set(default_output())
        self.root.after(100, self._poll); self._tick()

    def _build_header(self):
        BG = '#17324d'
        hdr = tk.Frame(self.root, bg=BG); hdr.pack(fill='x')
        tk.Label(hdr, text='KIT UPAH TSJ', bg=BG, fg='white',
                 font=('Segoe UI', 15, 'bold')).grid(row=0, column=0, sticky='w', padx=14, pady=(10, 0))
        tk.Label(hdr, text='Dashboard Payroll  -  v%s' % APP_VERSION, bg=BG, fg='#9fb6cc',
                 font=('Segoe UI', 9)).grid(row=1, column=0, sticky='w', padx=14, pady=(0, 10))
        tk.Label(hdr, text='Support Windows 7 / 10 / 11  -  tanpa install apapun', bg=BG, fg='#9fb6cc',
                 font=('Segoe UI', 9)).grid(row=0, column=1, rowspan=2, sticky='e', padx=14)
        hdr.columnconfigure(0, weight=1)

    def _build_menu(self):
        menubar = tk.Menu(self.root)
        m_file = tk.Menu(menubar, tearoff=0)
        m_file.add_command(label='Pilih folder sumber…', command=self.browse_base, accelerator='Ctrl+O')
        m_file.add_command(label='Pilih file output…', command=self.browse_out)
        m_file.add_separator()
        m_file.add_command(label='Buka folder output', command=self.open_output_dir)
        m_file.add_command(label='Keluar', command=self.root.destroy)
        menubar.add_cascade(label='File', menu=m_file)
        m_tools = tk.Menu(menubar, tearoff=0)
        m_tools.add_command(label='▶ Proses Kit Upah', command=self.start, accelerator='F5')
        m_tools.add_separator()
        m_tools.add_command(label='Cek Update dari GitHub', command=self.check_update)
        m_tools.add_separator()
        m_tools.add_command(label='Bersihkan layar log', command=self.clear_log)
        m_tools.add_command(label='Simpan log sebagai…', command=self.save_log_as)
        menubar.add_cascade(label='Tools', menu=m_tools)
        m_help = tk.Menu(menubar, tearoff=0)
        m_help.add_command(label='Cara Pakai', command=self.show_help)
        m_help.add_command(label='Tentang', command=self.show_about)
        menubar.add_cascade(label='Bantuan', menu=m_help)
        self.root.config(menu=menubar)
        self.root.bind_all('<Control-o>', lambda e: self.browse_base())
        self.root.bind_all('<F5>', lambda e: self.start())

    def _build_form(self):
        frm = ttk.Frame(self.root, padding=(10, 8)); frm.pack(fill='x')
        ttk.Label(frm, text='Folder closingan:').grid(row=0, column=0, sticky='w', pady=2)
        self.var_base = tk.StringVar()
        ttk.Entry(frm, textvariable=self.var_base).grid(row=0, column=1, sticky='ew', padx=6, pady=2)
        ttk.Button(frm, text='Pilih…', command=self.browse_base).grid(row=0, column=2, padx=2)
        ttk.Label(frm, text='File output:').grid(row=1, column=0, sticky='w', pady=2)
        self.var_out = tk.StringVar()
        ttk.Entry(frm, textvariable=self.var_out).grid(row=1, column=1, sticky='ew', padx=6, pady=2)
        ttk.Button(frm, text='Pilih…', command=self.browse_out).grid(row=1, column=2, padx=2)
        frm.columnconfigure(1, weight=1)
        frm_btn = ttk.Frame(self.root, padding=(10, 2)); frm_btn.pack(fill='x')
        self.btn_run = ttk.Button(frm_btn, text='▶  PROSES KIT UPAH', style='Big.TButton', command=self.start)
        self.btn_run.pack(side='left')
        self.btn_open = ttk.Button(frm_btn, text='Buka file hasil', command=self.open_output_file, state='disabled')
        self.btn_open.pack(side='left', padx=8)
        self.btn_update = ttk.Button(frm_btn, text='Cek Update', command=self.check_update)
        self.btn_update.pack(side='right')

    def _build_log(self):
        wrap = ttk.LabelFrame(self.root, text='Log  (Ctrl+A pilih semua  -  Ctrl+C salin  -  klik kanan menu)', padding=4)
        wrap.pack(fill='both', expand=True, padx=10, pady=(6, 4))
        # Toolbar tombol log (Copy primary, lainnya sekunder) — prominent di atas text
        bar = ttk.Frame(wrap); bar.grid(row=0, column=0, columnspan=2, sticky='ew', pady=(0, 4))
        bar.columnconfigure(0, weight=1)
        style_big = ttk.Style(); style_big.configure('Copy.TButton', font=('Segoe UI', 10, 'bold'), padding=(10, 6))
        self.btn_copy = ttk.Button(bar, text='▶▶  SALIN SEMUA LOG', style='Copy.TButton', command=self._log_copy_all)
        self.btn_copy.grid(row=0, column=0, sticky='w')
        ttk.Button(bar, text='Pilih Semua', command=self._log_select_all).grid(row=0, column=1, padx=(8, 0))
        ttk.Button(bar, text='Simpan ke File...', command=self._log_save_as).grid(row=0, column=2, padx=(4, 0))
        ttk.Button(bar, text='Bersihkan', command=self._log_clear).grid(row=0, column=3, padx=(4, 0))
        self.txt = tk.Text(wrap, wrap='none', bg='#111418', fg='#d8dee9',
                           insertbackground='#d8dee9', font=('Consolas', 10), state='normal',
                           undo=False, maxundo=0)
        vsb = ttk.Scrollbar(wrap, orient='vertical', command=self.txt.yview)
        hsb = ttk.Scrollbar(wrap, orient='horizontal', command=self.txt.xview)
        self.txt.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.txt.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns'); hsb.grid(row=1, column=0, sticky='ew')
        wrap.rowconfigure(0, weight=1); wrap.columnconfigure(0, weight=1)
        self.txt.tag_configure('err', foreground='#ff6b6b')
        self.txt.tag_configure('ok', foreground='#7ec699', font=('Consolas', 10, 'bold'))
        self.txt.tag_configure('hdr', foreground='#88c0d0')
        # Shortcut: Ctrl+A = select all (override Text default)
        self.txt.bind('<Control-a>', lambda e: (self.txt.tag_add('sel', '1.0', 'end'), 'break')[1])
        self.txt.bind('<Control-A>', lambda e: (self.txt.tag_add('sel', '1.0', 'end'), 'break')[1])
        # Right-click context menu: Copy / Select All / Clear / Save As
        self._log_menu = tk.Menu(self.root, tearoff=0)
        self._log_menu.add_command(label='Salin (Copy)', command=self._log_copy)
        self._log_menu.add_command(label='Pilih Semua (Select All)', command=self._log_select_all)
        self._log_menu.add_separator()
        self._log_menu.add_command(label='Bersihkan Log', command=self._log_clear)
        self._log_menu.add_command(label='Simpan Log ke File...', command=self._log_save_as)
        self.txt.bind('<Button-3>', self._log_popup)
        self._append_log('%s v%s siap.' % (APP_TITLE, APP_VERSION), 'hdr')
        self._append_log('Pipeline: %s' % (pipeline_path() or 'TIDAK DITEMUKAN — taruh di samping aplikasi'), 'hdr')
        self._append_log('Tip: Ctrl+A pilih semua  -  Ctrl+C salin  -  klik kanan untuk menu lengkap.', 'hdr')

    def _log_popup(self, event):
        try: self._log_menu.tk_popup(event.x_root, event.y_root)
        finally: self._log_menu.grab_release()

    def _log_copy(self):
        try:
            self.txt.config(state='normal')
            sel = self.txt.get('sel.first', 'sel.last')
            self.txt.config(state='disabled')
            if sel:
                self.root.clipboard_clear()
                self.root.clipboard_append(sel)
                self.lbl_status.config(text='Log disalin ke clipboard (%d karakter).' % len(sel))
        except tk.TclError:
            self.lbl_status.config(text='Tidak ada teks yang dipilih.')

    def _log_copy_all(self):
        try:
            # Enable widget first — state='disabled' on some Windows tkinter causes get() to return empty
            self.txt.config(state='normal')
            content = self.txt.get('1.0', 'end-1c')
            self.txt.config(state='disabled')
            if not content.strip():
                self.lbl_status.config(text='Log kosong — tidak ada yang disalin.')
                return
            self.root.clipboard_clear()
            self.root.clipboard_append(content)
            # update_idletasks only processes pending idle callbacks — safer than update()
            # (update() can re-enter event handlers and cause race conditions)
            self.root.update_idletasks()
            n = len(content)
            self.lbl_status.config(text='SELESAI — %d karakter log disalin ke clipboard.' % n)
            self._append_log('[COPY] %d karakter log disalin ke clipboard oleh user.' % n, 'ok')
        except Exception as e:
            self.lbl_status.config(text='Gagal salin: %s' % e)
            messagebox.showerror(APP_TITLE, 'Gagal salin log:\n%s' % e)

    def _log_select_all(self):
        self.txt.tag_add('sel', '1.0', 'end')
        self.lbl_status.config(text='Semua log dipilih  -  tekan Ctrl+C untuk salin.')

    def _log_clear(self):
        self.txt.config(state='normal')
        self.txt.delete('1.0', 'end')
        self.txt.config(state='disabled')
        self.lbl_status.config(text='Log dibersihkan.')

    def _log_save_as(self):
        f = filedialog.asksaveasfilename(
            title='Simpan log ke file', defaultextension='.log',
            initialfile='kit-upah-tsj-%s.log' % datetime.datetime.now().strftime('%Y%m%d-%H%M%S'),
            filetypes=[('Log file', '*.log'), ('Text', '*.txt'), ('Semua', '*.*')])
        if f:
            try:
                self.txt.config(state='normal')
                content = self.txt.get('1.0', 'end')
                self.txt.config(state='disabled')
                with open(f, 'w', encoding='utf-8') as fp: fp.write(content)
                self.lbl_status.config(text='Log disimpan ke %s' % os.path.basename(f))
            except Exception as e:
                messagebox.showerror(APP_TITLE, 'Gagal simpan log: %s' % e)

    def _build_statusbar(self):
        bar = ttk.Frame(self.root, relief='sunken'); bar.pack(fill='x', side='bottom')
        self.lbl_status = ttk.Label(bar, text='Siap', style='Status.TLabel'); self.lbl_status.pack(side='left')
        self.lbl_time = ttk.Label(bar, text='', style='Status.TLabel'); self.lbl_time.pack(side='right')

    def browse_base(self):
        d = filedialog.askdirectory(title='Pilih folder closingan', initialdir=self.var_base.get() or APP_DIR)
        if d:
            self.var_base.set(d)
            parent = os.path.dirname(d.rstrip('/\\'))
            if parent and os.path.isdir(parent):
                self.var_out.set(os.path.join(parent, 'Kit-Upah-TSJ-AUTO.xlsx'))

    def browse_out(self):
        f = filedialog.asksaveasfilename(
            title='Simpan file output', defaultextension='.xlsx',
            initialfile=os.path.basename(self.var_out.get() or 'Kit-Upah-TSJ-AUTO.xlsx'),
            initialdir=os.path.dirname(self.var_out.get()) or APP_DIR,
            filetypes=[('Excel', '*.xlsx')])
        if f: self.var_out.set(f)

    def start(self):
        if self.running: return
        base = self.var_base.get().strip(); out = self.var_out.get().strip()
        if not base: messagebox.showwarning(APP_TITLE, 'Pilih folder closingan dulu.'); return
        if not out: messagebox.showwarning(APP_TITLE, 'Tentukan nama file output dulu.'); return
        self.running = True; self.t_start = time.time(); self.last_out = out
        self.btn_run.config(state='disabled'); self.btn_open.config(state='disabled')
        self.btn_update.config(state='disabled'); self.lbl_status.config(text='Memproses…')
        try:
            os.makedirs(LOG_DIR, exist_ok=True)
            stamp = datetime.datetime.now().strftime('%Y%m%d-%H%M%S')
            self.log_file = open(os.path.join(LOG_DIR, 'run-%s.log' % stamp), 'w', encoding='utf-8')
        except Exception as e:
            self.log_file = None
            self._append_log('PERHATIAN: tidak bisa buat file log: %s' % e, 'err')
        threading.Thread(target=run_pipeline_worker, args=(base, out, self.q), daemon=True).start()

    def check_update(self):
        if self.running: return
        self.running = True
        self.btn_run.config(state='disabled'); self.btn_update.config(state='disabled')
        self.lbl_status.config(text='Cek update…')
        threading.Thread(target=update_worker, args=(self.q,), daemon=True).start()

    def open_output_dir(self):
        target = os.path.dirname(self.last_out or self.var_out.get()) or APP_DIR
        self._open_path(target)

    def open_output_file(self):
        if self.last_out and os.path.isfile(self.last_out): self._open_path(self.last_out)
        else: messagebox.showinfo(APP_TITLE, 'File output belum ada.')

    @staticmethod
    def _open_path(path):
        try:
            if sys.platform == 'win32': os.startfile(path)
            elif sys.platform == 'darwin': os.system('open "%s"' % path)
            else: os.system('xdg-open "%s"' % path)
        except Exception: pass

    def clear_log(self):
        self.txt.config(state='normal'); self.txt.delete('1.0', 'end'); self.txt.config(state='disabled')

    def save_log_as(self):
        f = filedialog.asksaveasfilename(defaultextension='.log', initialfile='log-kit-upah.log',
                                         filetypes=[('Log', '*.log'), ('Text', '*.txt')])
        if f:
            self.txt.config(state='normal')
            content = self.txt.get('1.0', 'end')
            self.txt.config(state='disabled')
            with open(f, 'w', encoding='utf-8') as fh: fh.write(content)

    def show_help(self):
        messagebox.showinfo(APP_TITLE + ' — Cara Pakai',
            '1. Taruh folder closingan-<bulan> di samping aplikasi (atau pilih folder mana saja).\n'
            '2. Klik "PROSES KIT UPAH" (atau tekan F5).\n'
            '3. Tunggu sampai muncul "PIPELINE OK" di log.\n'
            '4. Buka file hasil → cek sheet VALIDASI (semua harus OK) dan TRACE INTEGRASI.\n\n'
            'Update aplikasi:\n'
            '• Menu Tools → "Cek Update dari GitHub" — download versi terbaru otomatis,\n'
            '   lalu updater.bat mengganti file lama setelah aplikasi ditutup.\n\n'
            'Catatan:\n'
            '• Nama 12 file sumber harus PERSIS (lihat pesan error kalau ada yang kurang).\n'
            '• Untuk bulan baru: edit list "needed" dan RATE di pipeline_upah.py\n'
            '   (file-nya ada di samping aplikasi ini).\n'
            '• Log tiap run tersimpan otomatis di folder logs/.')

    def show_about(self):
        messagebox.showinfo(APP_TITLE + ' — Tentang',
            '%s v%s\n\nGUI wrapper untuk pipeline_upah.py — 100%% Python (tkinter + openpyxl).\n'
            'Repo: github.com/%s\n\nSumber data read-only: aplikasi tidak pernah mengubah file di folder closingan.\n'
            'Logika rekon: 110/120 zero-diff vs ASLI c26 (audit Juli 2026).'
            % (APP_TITLE, APP_VERSION, GH_REPO))

    def _append_log(self, line, tag=None):
        self.txt.config(state='normal')
        if tag: self.txt.insert('end', line + '\n', tag)
        else: self.txt.insert('end', line + '\n')
        self.txt.see('end'); self.txt.config(state='disabled')

    def _classify(self, line):
        up = line.upper()
        if 'ERROR' in up or 'GAGAL' in up or 'TRACEBACK' in up: return 'err'
        if 'PIPELINE OK' in up or 'SELESAI' in up or 'OK —' in up: return 'ok'
        if line.startswith('====') or line.startswith('MULAI PROSES'): return 'hdr'
        return None

    def _poll(self):
        try:
            while True:
                item = self.q.get_nowait(); kind = item[0]
                if kind == 'log':
                    line = item[1]
                    self._append_log(line, self._classify(line))
                    if self.log_file:
                        try: self.log_file.write(line + '\n'); self.log_file.flush()
                        except Exception: pass
                elif kind == 'update-done':
                    ok, msg = item[1], item[2]
                    self.running = False
                    self.btn_run.config(state='normal'); self.btn_update.config(state='normal')
                    self.lbl_status.config(text='Update selesai' if ok else 'Update gagal')
                    if ok and 'updater.bat' in msg:
                        if messagebox.askyesno(APP_TITLE, msg + '\n\nTutup aplikasi sekarang?'):
                            self.root.destroy()
                    else:
                        messagebox.showinfo(APP_TITLE, msg)
                else:
                    ok, elapsed = item[1], item[2]
                    self.running = False
                    self.btn_run.config(state='normal'); self.btn_update.config(state='normal')
                    if self.log_file:
                        try: self.log_file.close()
                        except Exception: pass
                        self.log_file = None
                    if ok:
                        self._append_log('SELESAI dalam %.1f detik — %s' % (elapsed, self.last_out), 'ok')
                        self.lbl_status.config(text='Selesai — %s' % os.path.basename(self.last_out))
                        self.btn_open.config(state='normal')
                    else:
                        self._append_log('GAGAL setelah %.1f detik — lihat log di atas.' % elapsed, 'err')
                        self.lbl_status.config(text='Gagal — cek log')
        except queue.Empty:
            pass
        self.root.after(100, self._poll)

    def _tick(self):
        if self.running: self.lbl_time.config(text='%.0fs' % (time.time() - self.t_start))
        self.root.after(500, self._tick)


def main():
    if '--selftest' in sys.argv:
        selftest(); return
    if sys.platform == 'win32':
        try:
            import ctypes; ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception: pass
    root = tk.Tk(); App(root); root.mainloop()


if __name__ == '__main__':
    main()
