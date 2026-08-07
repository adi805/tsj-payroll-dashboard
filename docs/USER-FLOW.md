# Kit Upah TSJ — User Flow & Wireframe

## 📐 Wireframe (ASCII)

```
┌────────────────────────────────────────────────────────────────────┐
│  KIT UPAH TSJ                              Support Windows 7 / 10 / 11 │
│  Dashboard Payroll  -  v1.0.0                  tanpa install apapun │
├────────────────────────────────────────────────────────────────────┤
│ File   Tools   Bantuan                                              │
├────────────────────────────────────────────────────────────────────┤
│                                                                    │
│  Folder closingan bulan ini                                         │
│  ┌──────────────────────────────────────────────┐  [ Browse... ]   │
│  │ C:\Users\finance\closingan-juli-2026         │                  │
│  └──────────────────────────────────────────────┘                  │
│                                                                    │
│  File output (kosongkan untuk default)                              │
│  ┌──────────────────────────────────────────────┐  [ Browse... ]   │
│  │ C:\Users\finance\Kit-Upah-TSJ-AUTO.xlsx     │                  │
│  └──────────────────────────────────────────────┘                  │
│                                                                    │
│         ┌──────────────────────────────┐  ┌──────────────────┐    │
│         │  ▶  PROSES KIT UPAH          │  │   Buka file hasil │    │
│         └──────────────────────────────┘  └──────────────────┘    │
│                                                          [ Cek     │
│                                                            Update ] │
│                                                                    │
├────────────────────────────────────────────────────────────────────┤
│ Log:                                                               │
│ ┌���──────────────────────────────��──────────────────────────────┐   │
│ │ MULAI PROSES — base: closingan-juli-2026                     │   │
│ │ =====================================                        │   │
│ │ OK - Roster: 120 karyawan (baris 9-128)                      │   │
│ │ OK - Rekon c26 vs ASLI: 110/120 zero-diff                    │   │
│ │ OK - Master coverage: 116/116 NIK                            │   │
│ │ OK - VALIDASI: 11 kategori + TOTAL                           │   │
│ │ OK - TF BNI: 115 transfer (111 exact + 4 Rp1)                │   │
│ │ OK - c28: 88 baris dikarantina                               │   │
│ │ SELESAI dalam 14.2 detik — Kit-Upah-TSJ-AUTO.xlsx            │   │
│ │                                                              │   │
│ │                                                              │   │
│ │ (Ctrl+A / Ctrl+C untuk copy • Klik kanan untuk menu)        │   │
│ └──────────────────────────────────────────────────────────────┘   │
│ [ Copy log ]  [ Save As... ]  [ Clear ]                            │
├────────────────────────────────────────────────────────────────────┤
│ Selesai — Kit-Upah-TSJ-AUTO.xlsx                  v1.0.0   14s     │
└────────────────────────────────────────────────────────────────────┘
```

## 🗺️ User Journey

### Journey 1: Proses Bulanan (Primary)

```
START
  │
  ▼
[1] Buka app (double-click KitUpahTSJ.exe)
    - Splash < 3 detik di SSD, < 8 detik di HDD
  │
  ▼
[2] Lihat form kosong, default value folder + output
    - Field "Folder closingan" = kosong, wajib diisi
    - Field "File output" = parent folder + Kit-Upah-TSJ-AUTO.xlsx
  │
  ▼
[3] Klik "Browse..." → file dialog → pilih folder closingan
    - 12 file Excel harus ada di folder (validasi di pipeline)
  │
  ▼
[4] (Opsional) Edit field "File output" kalau mau custom path
  │
  ▼
[5] Klik "▶  PROSES KIT UPAH"
    - Button jadi disabled, label "Processing..."
    - Status bar: "Memproses..."
  │
  ▼
[6] Log jalan realtime, scroll otomatis
    - Worker thread update log widget via queue
    - User lihat progress: roster → rekon → validasi → selesai
  │
  ▼
[7A] SUKSES: SELESAI dalam X detik
    - Status bar: "Selesai — Kit-Upah-TSJ-AUTO.xlsx"
    - Time: "Xs"
    - Button "Buka file hasil" enabled
    - Log file ter-save di logs/run-<timestamp>.log
  │
  ▼
[7B] GAGAL: ERROR line merah
    - Status bar: "Gagal: <reason>"
    - Log file ter-save (ada error)
    - User bisa Ctrl+A/C log → paste ke chat support
  │
  ▼
[8] Klik "Buka file hasil" → Excel terbuka otomatis
  │
  ▼
END (close app atau proses bulan depan)
```

### Journey 2: Update Aplikasi

```
START (Tools → Cek Update)
  │
  ▼
[1] App hit GH API: GET /repos/adi805/tsj-payroll-dashboard/releases/latest
    - Kalau no internet → "Tidak bisa cek update (no network)"
    - Kalau rate limit → "API rate limited, coba lagi nanti"
  │
  ▼
[2] Parse JSON, compare version
    - Sama / lebih lama → "Anda sudah versi terbaru"
    - Lebih baru → "v1.0.1 tersedia (saat ini v1.0.0). Update sekarang?"
  │
  ▼
[3] User klik "Ya"
    - Download zip ke %TEMP%\tsj_update\
    - Progress: "Downloading... 45%"
  │
  ▼
[4] Extract ke _update/ (sebelah exe)
    - Progress: "Extracting..."
  │
  ▼
[5] Tulis UPDATER_BAT (taskkill + copy + restart)
  │
  ▼
[6] Dialog: "App akan close dan restart dengan versi baru. Lanjut?"
  │
  ▼
[7] User klik "Ya" → os.startfile('UPDATER_BAT') → root.destroy()
    - App close
    - 2 detik kemudian: UPDATER_BAT kill app, copy exe baru, restart
    - User lihat splash screen versi baru
  │
  ▼
END
```

### Journey 3: Log Copy untuk Support

```
START (ada error, mau kirim log ke support)
  │
  ▼
[1] Log area visible dengan error merah
  │
  ▼
[2A] Klik tombol "Copy log" di bawah log
    - Semua log text ke clipboard
  │
  ▼
[2B] Atau: Ctrl+A (select all) → Ctrl+C (copy)
  │
  ▼
[2C] Atau: Klik kanan → Copy (atau Select All → Copy)
  │
  ▼
[3] Paste ke chat (Telegram/WhatsApp)
  │
  ▼
END
```

## 🔍 Edge Cases

| Case | Handling |
|------|----------|
| Folder closingan kosong / tidak ada 12 file | Pipeline: log error "Folder tidak valid", exit 1 |
| File output read-only (Excel masih terbuka) | Pipeline: log error "File sedang dibuka", exit 1. User close Excel, retry |
| Disk penuh | Pipeline: log error "[Errno 28] No space left", exit 1 |
| No internet saat Cek Update | "Tidak bisa cek update (no network)" — gak break app |
| API rate limit | "GitHub API rate limited, coba lagi nanti" |
| App crash mid-process | Excel yang ter-generate mungkin corrupt → user hapus, retry. Log gak ke-save (ini kenapa pakai thread yg write langsung, bukan buffered) |
| Multiple PROSES click | Button disabled saat running, gak bisa double-click |
| App force close saat update | Exe baru sudah ter-download ke _update/, next launch: "Update terdeteksi, install sekarang?" |
| Win 7 (UCRT) | ✅ Dibundel di exe sejak v1.0.1 — tidak perlu install KB2999226 |
| Folder path ada spasi / Unicode | tkinter filedialog handle, pipeline string-safe |

## ♿ Accessibility

- **Keyboard:** Tab navigation antar field, Enter = trigger default button
- **Font:** Segoe UI 9pt (default Windows, gak butuh extra font)
- **Color blind:** Log tag pakai prefix (OK-/ERROR-/WARNING), bukan cuma warna
- **Screen reader:** Status bar text-to-speech friendly ("Selesai", "Gagal: ...")

## 🌐 i18n

v1.0.0: Indonesia only (sesuai primary user). Future: English translation (gak urgent).
