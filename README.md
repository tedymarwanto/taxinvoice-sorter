# Tax Invoice Sorter — Build EXE via GitHub

Panduan build aplikasi jadi file `.exe` Windows secara otomatis di cloud GitHub, tanpa perlu laptop Windows.

---

## Cara Build (Step by Step)

### 1. Buat Repository Baru di GitHub
- Login ke https://github.com
- Klik tombol **"+"** di pojok kanan atas → **"New repository"**
- Kasih nama, misal: `tax-invoice-sorter`
- Pilih **Public** atau **Private** (bebas)
- Klik **"Create repository"**

### 2. Upload File
- Di halaman repo yang baru dibuat, klik **"uploading an existing file"**
- Drag & drop SEMUA file & folder dari folder ini:
  - `tax_invoice_sorter.py`
  - `requirements.txt`
  - `README.md`
  - Folder `.github` (beserta isinya)
- **PENTING:** Pastikan folder `.github/workflows/build.yml` ikut terupload
- Klik **"Commit changes"**

> **Catatan:** Kalau drag-drop folder `.github` ga bisa, upload manual:
> Klik **"Create new file"** → ketik nama: `.github/workflows/build.yml` → paste isi file build.yml → commit

### 3. Tunggu GitHub Build Otomatis
- Setelah upload, klik tab **"Actions"** di repo lo
- Bakal ada proses **"Build Windows EXE"** yang sedang jalan (ikon kuning berputar)
- Tunggu 3-5 menit sampai jadi centang hijau ✅

### 4. Download EXE
- Klik proses build yang udah selesai (centang hijau)
- Scroll ke bawah ke bagian **"Artifacts"**
- Klik **"TaxInvoiceSorter-Windows"** untuk download
- File nya berupa `.zip` — extract untuk dapat `TaxInvoiceSorter.exe`

### 5. Jalankan di PC
- Double-click `TaxInvoiceSorter.exe`
- Aplikasi langsung jalan tanpa perlu install Python!

---

## Catatan Penting

- **Windows Defender / Antivirus:** Kadang `.exe` hasil PyInstaller ke-flag sebagai "unknown app". Klik **"More info"** → **"Run anyway"** kalau muncul warning. Ini normal untuk aplikasi yang belum ada digital signature.
- **Cache:** Aplikasi menyimpan cache di `C:\Users\NamaUser\.tax_invoice_sorter_cache\` — otomatis menyesuaikan dengan user PC.
- **Tidak butuh database:** Semua diproses dari file PDF & Excel langsung, plus cache JSON di disk.

---

## Cara Update Aplikasi

Kalau ada perubahan code:
1. Edit/replace `tax_invoice_sorter.py` di repo
2. Commit changes
3. GitHub otomatis build ulang → download EXE baru dari Actions

---

## Troubleshooting

**Build gagal (centang merah):**
- Klik proses build → lihat log error nya
- Biasanya karena file `tax_invoice_sorter.py` ga keupload dengan benar

**EXE ga mau jalan:**
- Pastikan Windows nya 64-bit (mayoritas PC sekarang)
- Coba klik kanan → "Run as administrator"
