# Changelog

All notable changes to Kit Upah TSJ Dashboard.

## [1.0.5] - 2026-08-07

### Fixed
- SALIN SEMUA LOG: widget state=normal sebelum get() agar clipboard tidak kosong (Windows tkinter)
- Right-click > Salin: same fix — enable widget sebelum get() selected text
- Toolbar/Simpan log: same fix — enable widget sebelum get() + save ke file
- _log_clear: state='disabled' setelah delete (widget harus tetap readonly setelah clear)

## [1.0.4] - 2026-08-07

### Fixed
- SALIN SEMUA LOG: fix button tidak berfungsi (clipboard copy) pada Windows dengan tkinter tertentu — widget state dinormalkan sebelum get() agar clipboard tidak kosong

## [1.0.3] - 2026-08-07

### Fixed
- Pipeline: sheet DATA INPUT tidak wajib ada di file sumber - jika tidak ada, pipeline auto-rebuild ledger dari roster (Upah Juli 2026) + Data Security (kolom AU) - tidak lagi crash KeyError di step 2/6
- Aturan rebuild diverifikasi 270/270 cocok dengan ledger asli (265 baris roster + 5 BBM audit)
- Handling baris 'T O T A L' subtotal dan nama duplikat (mapping NIK via urutan baris, bukan nama)

## [1.0.2] - 2026-08-07

### Fixed
- Updater: nama exe di updater.bat salah (KitUpahTSJ-Dashboard.exe -> KitUpahTSJ.exe) - update in-app gagal diam-diam sejak v1.0.0

## [1.0.1] - 2026-08-07

### Fixed
- UCRT: 16 DLL (ucrtbase.dll + api-ms-win-crt-*.dll) dibundle di dalam exe onefile — Windows 7 langsung jalan tanpa install KB2999226
- ASLI-backup.xlsx jadi opsional — pipeline tetap jalan dengan 11 file (fallback ke file utama)

## [1.0.0] - 2026-08-06

### Added
- Dashboard payroll GUI (tkinter) untuk pipeline_upah.py
- Auto-update dari GitHub releases
- Support Windows 7 SP1 / 10 / 11 (64-bit), tanpa install runtime
- Log setiap run tersimpan otomatis di folder logs/
- Toolbar: SALIN SEMUA LOG, Pilih Semua, Simpan ke File, Bersihkan
- Right-click menu: Salin, Pilih Semua, Bersihkan, Simpan Log
