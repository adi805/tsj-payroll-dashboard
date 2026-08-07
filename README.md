# Kit Upah TSJ — Dashboard Payroll

**Aplikasi desktop ringan untuk auto-populate Kit-Upah-TSJ-AUTO.xlsx dari folder closing bulanan (12 file).**

| Spec | Value |
|------|-------|
| Bahasa | Python 3.9 (kompatibel Windows 7 SP1+) |
| GUI | Tkinter (built-in, no Electron, no browser) |
| Packager | PyInstaller 5.13.2 (`--onefile` self-contained) |
| Ukuran exe | ±12-15 MB |
| RAM idle | ±40-60 MB |
| Disk | Tidak ada (portable, no install) |
| Network | Hanya saat klik "Cek Update" |
| UCRT (Win 7) | ✅ Dibundel di dalam exe (v1.0.1+) — zero install |

## 🎯 Untuk User (Finance)

1. **Download** `KitUpahTSJ-windows.zip` dari [Releases](../../releases)
2. **Extract** ke folder mana saja (mis. `Desktop\KitUpahTSJ\`)
3. **Double-click** `KitUpahTSJ.exe`
4. **Win 7**: UCRT (`ucrtbase.dll` + `api-ms-win-crt-*`) sudah dibundel di dalam exe sejak v1.0.1 — **langsung jalan, tanpa install apa pun**
5. Isi **Folder closingan** → klik **▶  PROSES KIT UPAH** → tunggu selesai → klik **Buka file hasil**

**Update otomatis**: Tools → Cek Update. Tidak perlu download manual lagi.

## 🛠️ Untuk Developer

### Repo structure
```
dashboard/
├── app_gui.py          # Main GUI (tkinter, 500 lines, no emoji)
├── pipeline_upah.py    # Payroll engine (runpy-loaded, editable per-bulan)
├── requirements.txt    # openpyxl==3.1.2
├── docs/
│   ├── PRD.md          # Product Requirements
│   ├── TECH-SPEC.md    # Arsitektur + Win7 matrix + PC-kentang
│   └── USER-FLOW.md    # Wireframe + journey
└── .github/workflows/
    └── build.yml       # Windows build → Release
```

### Local dev
```bash
cd dashboard
pip install -r requirements.txt
python app_gui.py
# self-test (no GUI):
python app_gui.py --selftest ../closingan-juli-2026 /tmp/out.xlsx
```

### Build lokal (butuh Windows)
```bash
pip install pyinstaller==5.13.2
# (opsional) staging UCRT buat target Win7:
#   New-Item -ItemType Directory -Force ucrt
#   Copy-Item C:\Windows\System32\ucrtbase.dll ucrt\
#   Copy-Item C:\Windows\System32\api-ms-win-crt-*.dll ucrt\
#   lalu tambahkan --add-binary "ucrt\<dll>;." per DLL ke perintah di bawah
pyinstaller --noconfirm --onefile --windowed \
  --name "KitUpahTSJ" \
  --add-data "pipeline_upah.py;." \
  app_gui.py
# Output: dist/KitUpahTSJ.exe
```

### Release flow
```bash
git tag v1.0.1
git push origin v1.0.1
# GitHub Action auto-build → publish Release dengan zip asset
```

## 📚 Docs
- [PRD](dashboard/docs/PRD.md) — Product Requirements
- [Tech Spec](dashboard/docs/TECH-SPEC.md) — Arsitektur, Win7 matrix, PC kentang
- [User Flow](dashboard/docs/USER-FLOW.md) — Wireframe + journey

## 🔒 Privacy
Repo ini **hanya berisi logic + konstanta (RATE)**. TIDAK ada data karyawan, tidak ada data gaji, tidak ada folder closing. Semua data payroll tetap lokal di mesin user.

## 📝 License
Private — Adi / TSJ internal use.
