# Kit Upah TSJ — Dashboard Payroll

Dashboard desktop ringan untuk auto-populate `Kit-Upah-TSJ-AUTO.xlsx` dari folder closing bulanan (12 file Excel).

**Versi:** 1.0.0
**Status:** Production-ready
**Target:** Windows 7 SP1 / 10 / 11 (64-bit)
**Author:** adi805

---

## 🎯 Masalah
Pipeline `pipeline_upah.py` (CLI, Python 3.9) sudah jalan dan menghasilkan Excel yang validated 110/120 zero-diff vs ASLI. Tapi untuk finance staff:
- Harus buka terminal / command prompt
- Harus tau argumen CLI
- Log print ke terminal, gak ke-save → kalau error susah trace
- Update = download manual zip, replace exe manual

## 👥 User
- **Primary:** Finance staff TSJ (operator), non-technical, pakai Win7 lama
- **Secondary:** Adi (owner / developer), maintenance + update

## ✅ Goals
1. **1 exe, double-click** — gak perlu install Python, gak perlu terminal
2. **GUI form** — pilih folder closing → klik tombol → log jalan di window
3. **Log per-transaksi tersimpan** ke `logs/run-<timestamp>.log` (auto)
4. **Update dari aplikasi** — Tools → Cek Update, no manual download
5. **Win 7 SP1+** support
6. **Ringan** — bisa jalan di PC kentang (RAM 4GB, HDD)

## ❌ Non-Goals
- Tidak ada web UI / browser
- Tidak ada database / multi-user
- Tidak ada network call (kecuali manual update check)
- Tidak handle multiple closing folder paralel

## 📦 Features (F1-F8)

| # | Feature | Acceptance |
|---|---------|-----------|
| F1 | Folder picker (tkinter filedialog) | User bisa browse & pilih folder closingan |
| F2 | File output path (default: parent/Kit-Upah-TSJ-AUTO.xlsx) | Editable text field |
| F3 | PROSES button → run pipeline via worker thread | UI gak freeze, log update realtime via queue |
| F4 | Log area (dark console, 5 tag warna: err/ok/hdr/dim) | Bisa di-copy (Ctrl+A/C + klik kanan + tombol Save As) |
| F5 | Auto-save log per run ke `logs/run-<YYYYMMDD-HHMMSS>.log` | Format = persis sama dengan log window |
| F6 | Buka file hasil button | `os.startfile(path)` di Windows |
| F7 | Cek Update via GitHub Releases API | Download zip → extract ke `_update/` → close app → replace exe → restart |
| F8 | `--selftest` CLI mode | Headless test pipeline tanpa GUI, exit code 0/1 |

## 📊 Success Metrics
- Runs di Win 7 SP1 x64 zero install (cukup UCRT KB2999226)
- 1-click process (isi folder → klik tombol → selesai)
- Log file selalu ter-generate per run
- Update tanpa download manual (Cek Update dari app)
- Cold start < 5 detik di HDD mekanik

## 🛠️ Constraints
- **Python only** — tkinter (built-in) + openpyxl + PyInstaller
- **Onefile** — `--onefile` PyInstaller, self-contained, zero target install
- **Ringan** — exe < 20MB, RAM idle < 100MB
- **No emoji** di source code (font fallback Win 7 bisa blank)
- **No DB** — semua ephemeral
- **Privacy** — no employee data, no salary data di repo (cuma logic + RATE)

## 📅 Timeline
| Phase | Target | Status |
|-------|--------|--------|
| v1.0.0 | Single-user dashboard Win7+ | ✅ Ready (turn ini) |
| v1.1.0 | Multi-bulan (dropdown pilih bulan) | Planned |
| v1.2.0 | Schedule auto-run (cron-like) | Planned |
| v2.0.0 | Multi-user via shared network folder | Future |

## 🧪 Testing
- `python app_gui.py --selftest <closingan> /tmp/out.xlsx` — pipeline headless
- Manual GUI test: klik-klik flow
- Win 7 SP1 VM test: zero install
- PC kentang test: cold start time, RAM peak

Lihat [User Flow](USER-FLOW.md) untuk wireframe + journey detail.
Lihat [Tech Spec](TECH-SPEC.md) untuk arsitektur + Win7 matrix + PC kentang analysis.
