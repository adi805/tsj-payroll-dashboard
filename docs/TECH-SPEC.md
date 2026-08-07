# Kit Upah TSJ — Technical Specification

## 🏗️ Arsitektur

```
┌─────────────────────────────────────────────┐
│ App (main thread, tkinter UI)                │
│  - menu, form, log text, status bar          │
│  - var_base (folder), var_out (file output)  │
│  - btn_proses, btn_open, btn_update          │
└──────────────┬───────────────────────────────┘
               │ click "PROSES"
               ▼
┌─────────────────────────────────────────────┐
│ Worker thread (daemon)                       │
│  - runpy.run_path('pipeline_upah.py')        │
│  - stdout → QWriter → queue.Queue            │
│  - return code 0/1                           │
└──────────────┬───────────────────────────────┘
               │ queue.get() tiap 100ms
               ▼
┌─────────────────────────────────────────────┐
│ Log text widget (dark, 5 tag warna)          │
│  - err (red), ok (green), hdr (cyan)         │
│  - dim (grey), default (white)               │
│  - Ctrl+A/C + right-click Copy menu          │
│  - auto-save ke logs/run-<stamp>.log         │
└─────────────────────────────────────────────┘
```

**Kenapa runpy.run_path dan bukan import?**
- Pipeline bisa di-edit per-bulan (RATE, formula adjustment) tanpa rebuild exe
- Exe extract ke `_MEIPASS` saat jalan; pipeline_upah.py di-copy ke `dashboard/` saat build, satu folder dengan exe
- Untuk mode portable (zip extract), pipeline_upah.py ada di sebelah exe → `pipeline_path()` resolve ini dulu
- Untuk mode dev (`python app_gui.py`), resolve ke `..\pipeline_upah.py`

**Threading model:**
- Main thread = tkinter UI (single-threaded, gak boleh di-block)
- Worker thread = `runpy.run_path`, stdout di-capture via `contextlib.redirect_stdout` + custom QWriter
- queue.Queue = bridge stdout → UI (thread-safe)
- `_drain_queue()` dipanggil `root.after(100, ...)` setiap 100ms, baca queue, append ke log widget

## 📦 Dependencies

| Package | Version | Kenapa |
|---------|---------|--------|
| Python | 3.9 | Last version yang support Windows 7 |
| tkinter | stdlib | GUI built-in, no install |
| openpyxl | 3.1.2 | xlsx read/write (pipeline) |
| PyInstaller | 6.3.0 | onefile packaging |
| requests | stdlib (urllib) | GitHub API call (no extra dep) |

**Kenapa gak pake library lain?**
- `colorama` → gak perlu, tkinter Tag config udah handle warna
- `tqdm` → pipeline udah ada progress print sendiri
- `PyQt` → terlalu berat, gak perlu
- `customtkinter` → extra dep, theme udah cukup via `ttk.Style`

## 🔄 Updater Contract

**Trigger:** Tools → Cek Update (atau `Ctrl+U`)

**Flow:**
1. `GET https://api.github.com/repos/adi805/tsj-payroll-dashboard/releases/latest` (no auth, public repo)
2. Parse JSON, ambil `tag_name` (e.g. `v1.0.1`) + `assets[0].browser_download_url` (zip)
3. Compare `tag_name` vs `APP_VERSION` (semver tuple compare) → kalau remote > local, prompt user
4. User klik "Ya, update" → download zip ke `%TEMP%\tsj_update\`
5. Extract → `KitUpahTSJ.exe` di-copy ke `_update/` (sebelah current exe)
6. Tulis `UPDATER_BAT` script:
   ```bat
   @echo off
   timeout /t 2 /nobreak > nul
   taskkill /F /IM KitUpahTSJ.exe
   timeout /t 1 /nobreak > nul
   copy /Y "_update\KitUpahTSJ.exe" "KitUpahTSJ.exe"
   rmdir /S /Q "_update"
   start "" "KitUpahTSJ.exe"
   del "%~f0"
   ```
7. `os.startfile('UPDATER_BAT')` → batch run async
8. `root.destroy()` → app close

**Idempotent:** kalau gagal (no internet, 404, dll), log error, gak break app.

## 🪟 Windows 7 Compatibility Matrix

| Komponen | Win 7 SP1 x64 | Status |
|----------|---------------|--------|
| Python 3.9.x | ✅ Last supported | Bundled di PyInstaller |
| tkinter (Tcl/Tk 8.6) | ✅ Built-in | PyInstaller bundles Tcl |
| openpyxl | ✅ Pure Python | Works |
| PyInstaller 5.13.2 | ✅ Tested Win 7 | onefile boot OK (v6+ tidak dijamin Win7 — pin 5.13.2) |
| DPI awareness (shcore.dll) | ⚠️ Optional | Try/except wrapped |
| UCRT (api-ms-win-crt-runtime) | ✅ DIBUNDEL di exe | ucrtbase.dll + api-ms-win-crt-*.dll di dalam bundle (v1.0.1+) — zero install |
| Emoji (U+1F600+) | ❌ Font gap | Tidak ada di source code |
| Non-ASCII (—, …, •, →, ▶) | ✅ Segoe UI Symbol | Used sparingly |
| Segoe UI font | ✅ Built-in Win 7+ | Default font |

**UCRT bundling (v1.0.1+):** DLL Universal C Runtime (`ucrtbase.dll` + `api-ms-win-crt-*.dll`) ikut ter-bundle **DI DALAM exe** (PyInstaller `--add-binary`, di-stage dari `C:\Windows\System32` di build runner). **Win 7 jalan tanpa install KB2999226 / apa pun.** Check di `__init__` di bawah tetap ada sebagai belt-and-suspenders (kalau ada DLL yang gagal load → dialog informatif, bukan crash cryptic):
```python
try:
    ctypes.CDLL('api-ms-win-crt-runtime-l1-1-0.dll')
except OSError:
    # Show dialog dengan link KB2999226, exit gracefully
    pass
```

**Tcl/Tk theme:** default 'clam' lebih ringan dari 'vista'/'xpnative', gak butuh extra DLL.

## 💻 PC Kentang Analysis

**Spec target (PC kentang):**
- CPU: Intel Celeron / AMD A4 (1.6 GHz)
- RAM: 4 GB DDR3
- Disk: HDD 5400 RPM
- GPU: Intel HD integrated
- OS: Windows 7 SP1 / Windows 10

### Opsi yang sudah dievaluasi

| Opsi | Size | RAM idle | Cold start | Verdict |
|------|------|----------|------------|---------|
| **tkinter + PyInstaller onefile** | 12-15 MB | 40-60 MB | 3-8 s (HDD extract) | ✅ **Pilih** |
| tkinter + PyInstaller onedir | 12-15 MB (split) | 40-60 MB | <1 s (no extract) | ✅ Alternative |
| Electron + Node.js | 200+ MB | 300+ MB | 5-15 s | ❌ Too heavy |
| PyQt5 + PyInstaller | 30-40 MB | 80-100 MB | 4-10 s | ❌ Overkill |
| .NET WPF + C# | 8-12 MB | 50-70 MB | 2-5 s | ❌ Stack beda, maintainability |
| Pure .bat (no GUI) | <1 KB | 0 MB | <1 s | ❌ Gak ada dashboard |

**Rekomendasi final: tkinter + PyInstaller onefile.**

### Kenapa onefile vs onedir?

- **onefile**: 1 file `.exe`, extract ke `%TEMP%` setiap launch (+3-8s di HDD lemot). Lebih simple untuk user.
- **onedir**: folder `KitUpahTSJ/` dengan banyak file (.exe + .dll + Tcl). Cold start <1s karena gak ada extract. Tapi user harus maintain folder.

**Default release = onefile** (simpler). Kalau user complain cold start lambat di HDD → publish 2 asset di Release (onefile + onedir) dan kasih instruksi.

### Kenapa bukan UPX compression?

- UPX memperkecil size 30-50% tapi:
  - Antivirus false positive rate tinggi (Avast, AVG, Windows Defender kadang flag)
  - Decompress overhead di startup (+1-2s)
  - Bundle Tcl/Tk yang udah gede, savings kecil
- Trade-off: prefer size normal tapi zero AV issue + startup cepat

### Real bottleneck di PC kentang

Bukan GUI — **pipeline parsing xlsx 2.9MB** = 8-15 detik (openpyxl pure Python, gak ada C extension). Gak bisa dihindari, ini cost unavoidable. Log widget nongol progress "OK - roster: 120 karyawan" dll supaya user gak mengira hang.

## 🧪 Testing Strategy

| Test | Method | Pass criteria |
|------|--------|---------------|
| Pipeline correctness | `--selftest` headless | Exit 0, file generated, 110/120 recon |
| Win 7 boot | VM Win 7 SP1 fresh | App launch < 5s, no missing DLL |
| Cold start HDD | Stopwatch | < 8s di 5400 RPM |
| RAM peak | Task Manager | < 100 MB during run |
| Update flow | Mock GH API | Download → extract → bat run → restart |
| Log copy | Manual: Ctrl+A/C | Text ter-copy ke clipboard |
| Log save | Manual: "Save As" | File ter-create di lokasi yang dipilih |

## 🚀 Deployment

**Channel:** GitHub Releases (zero cost, no infra)

**Trigger:** Push tag `v*` ke main → Action auto-build Windows → publish Release

**Versioning:** Semver (`vMAJOR.MINOR.PATCH`)
- MAJOR: breaking change (UI redesign, data format change)
- MINOR: new feature (F-number baru)
- PATCH: bug fix

**Rollback:** User bisa download Release lama + replace exe manual. App gak auto-downgrade.
