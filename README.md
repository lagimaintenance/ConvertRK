# ConvertRK Multi-Bank — PDF Rekening Koran → GL

Aplikasi Streamlit ini menghasilkan **satu sheet GL** dengan **font, warna, border, ukuran kolom, judul, dan format nominal mengikuti `Bank_template.xlsx`** (template kosong dari contoh `Bank(1).xlsx`). Tidak melakukan klasifikasi/mapping akun.

## Bank yang telah diuji

- **BCA:** 6 transaksi, saldo akhir cocok dengan rekening koran Januari 2026.
- **BSI:** 111 transaksi, total masuk/keluar dan saldo akhir cocok.
- **CIMB Niaga:** 67 transaksi; baris PDF dari terbaru ke terlama dibalik ke urutan kronologis.
- **Mandiri:** 394 transaksi; semua total dan jumlah transaksi sesuai ringkasan PDF.
- **Muamalat:** 80 transaksi, semua saldo transaksi cocok. **Saldo awal dan akhir tidak dicantumkan sebagai ringkasan tersendiri**; saldo awal diturunkan dari saldo setelah transaksi pertama dan nominal transaksi pertama, sedangkan saldo akhir berasal dari saldo transaksi terakhir. Aplikasi memberi catatan ini.
- **Permata:** 806 transaksi (nomor 1–806 tanpa celah), saldo akhir cocok dengan saldo ledger PDF.

Untuk bank lain terdapat parser tabel generik yang memerlukan kolom tanggal, debit, kredit, dan saldo dengan teks yang dapat diekstrak. **Belum bisa menjamin semua format bank di dunia**: jika format tidak dikenali, PDF dinyatakan gagal. PDF scan/gambar tanpa teks belum mendukung OCR.

## Instalasi Mac

Unduh ZIP dan tempatkan ke `~/Documents/ConvertRK` (folder bernama ConvertRK, huruf besar/kecil harus sesuai nama folder Anda). Di Terminal:

```bash
cd ~/Documents/ConvertRK
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

**Opsional untuk PDF Permata panjang:** jika memiliki Homebrew, jalankan `brew install poppler`. Ini memasang `pdftotext` yang mempercepat pembacaan PDF Permata. Bila tidak terpasang, aplikasi memakai `pdfplumber` sebagai cadangan (lebih lambat). Tidak membutuhkan `fitz` / PyMuPDF.

Jika menggantikan versi lama, **ganti keempat file** `app.py`, `multi_bank_parser.py`, `Bank_template.xlsx`, dan `requirements.txt`. File `packages.txt` berguna untuk hosting Linux.

Akses lokal: http://localhost:8501

## Deploy Streamlit Community Cloud (gratis)

Unggah ke repository GitHub **hanya**:

- `app.py`
- `multi_bank_parser.py`
- `Bank_template.xlsx`
- `requirements.txt`
- `packages.txt`
- `README.md` (opsional)

Buat aplikasi di https://share.streamlit.io, pilih branch, lalu `app.py` sebagai Main file. Streamlit menggunakan `requirements.txt` untuk instalasi Python dan `packages.txt` untuk Poppler pada Linux.

**Keamanan:** jangan pernah unggah PDF rekening koran dan Excel berisi transaksi asli ke GitHub. Gunakan uploader di aplikasi. Untuk kerahasiaan rekening koran, jalankan secara lokal; Streamlit Community Cloud bukan tempat ideal untuk dokumen keuangan rahasia.

## Ketentuan GL

- Debet GL Bank = **uang masuk** pada rekening koran (Credit/CR/K dari bank).
- Kredit GL Bank = **uang keluar** pada rekening koran (Debit/DB/D dari bank).
- Satu baris **Saldo Bank Awal** untuk setiap PDF, bukan mutasi.
- Tidak membuat jurnal double; semua mutasi termasuk biaya kecil (misalnya admin dan pajak) tetap dipertahankan.
- Uraian dan referensi bank dimasukkan ke kolom Uraian untuk ditelusuri.
- Kode lawan akun dan nama akun dikosongkan untuk verifikasi manual.
- Jika ada selisih saldo/ringkasan, aplikasi akan memberi peringatan; jangan gunakan data bermasalah tanpa pengecekan.
- Nomor rekening dipertahankan pada uraian saldo awal per dokumen.

Untuk bank yang tidak tercantum dalam contoh, usahakan memakai PDF digital asli (bukan hasil foto/scan). Bila hasilnya gagal atau tidak cocok, contoh tambahan bank tersebut diperlukan untuk menyesuaikan parser.
