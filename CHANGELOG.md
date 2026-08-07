# Changelog

All notable changes to Kit Upah TSJ Dashboard.

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

### Added
- UCRT di-bundle ke dalam exe (ucrtbase.dll + api-ms-win-crt-*.dll dari Windows SDK Redist) - Win 7 tanpa install KB2999226

### Fixed
- Pipeline: ASLI-backup jadi opsional (fallback cache ke file utama) - tidak lagi menghentikan pipeline jika file backup tidak ada

## [1.0.0] - 2026-08-07

### Added
- Initial release
- F1: Folder picker (Browse button)
- F2: File output path with default value
- F3: PROSES button with worker thread (no UI freeze)
- F4: Log area with 5 color tags (err/ok/hdr/dim/default)
- F5: Auto-save log per run to `logs/run-<timestamp>.log`
- F6: Buka file hasil button (os.startfile)
- F7: Cek Update via GitHub Releases API + auto-replace
- F8: `--selftest` CLI mode (headless pipeline test)
- Copy log: Ctrl+A/C, right-click menu, "Copy log" button
- Save As log: export log to user-chosen file
- Docs: PRD, Tech Spec, User Flow
