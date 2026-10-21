# ConvertRK — PDF Rekening Koran menjadi Buku Bank Excel

Aplikasi Streamlit untuk **rekening koran BCA berupa PDF teks** yang formatnya seperti e-statement Januari 2026. Aplikasi membaca mutasi, membuat Buku Bank, memberi usulan kode lawan akun, menyediakan koreksi manual, dan melakukan rekonsiliasi saldo. **Transaksi yang belum pasti klasifikasinya tidak dipaksakan menjadi final**.

## 1. Struktur folder di Mac

Simpan berkas proyek hasil ZIP ke:

```text
/Users/<nama-akun-mac>/Documents/convertRK/
├── app.py
├── bank_parser.py
├── xlsx_engine.py
├── coa_default.py
├── requirements.txt
├── README.md
├── .gitignore
└── .streamlit/
    └── config.toml
```

Opsional: simpan `Kode Akun.xlsx` dan `Bank(1).xlsx` **hanya di folder lokal**, bila ingin aplikasi memuat keduanya otomatis. Jangan memasukkan rekening koran atau file contoh nyata ke repositori publik.

## 2. Jalankan lokal di Mac

Setelah mengunduh ZIP, buka **Terminal**. Jika ZIP tersimpan di folder Downloads, jalankan:

```bash
mkdir -p ~/Documents/convertRK
unzip ~/Downloads/convertRK_streamlit.zip -d ~/Documents/convertRK
cd ~/Documents/convertRK
python3 -m venv .venv
source .venv/bin/activate
python3 -m pip install -r requirements.txt
python3 -m streamlit run app.py
```

Buka alamat lokal yang ditampilkan Terminal, biasanya `http://localhost:8501`.

Untuk menjalankan kembali:

```bash
cd ~/Documents/convertRK
source .venv/bin/activate
python3 -m streamlit run app.py
```

Untuk menghentikan: tekan **Ctrl+C** di Terminal.

## 3. Cara menggunakan

1. Periksa nama perusahaan dan akun bank (default `111 – Bank`).
2. **Kode Akun.xlsx** bersifat opsional karena 46 akun yang tersedia pada file referensi Anda sudah tersedia sebagai bawaan dalam `coa_default.py`. Anda dapat mengunggah file COA lain untuk menggantinya.
3. Bila menginginkan klasifikasi transaksi Januari **identik** dengan contoh, unggah **Bank(1).xlsx**. Aplikasi menggunakan tanggal, uraian, debit, dan kredit untuk mencocokkan transaksi secara tepat.
4. Unggah satu atau lebih PDF BCA. Anda boleh mengunggah file `.zip` berisi kumpulan PDF.
5. Klik **Proses rekening koran**.
6. Periksa ringkasan PDF dan saldo. Edit **Kode Lawan Akun**, **No. Bukti**, **Nama Vendor**. Centang **Verifikasi** untuk klasifikasi yang sudah benar.
7. Klik **Unduh Buku Bank Excel**.

File Excel berisi:

- **Bank** (gabungan transaksi per rekening, sesuai kolom contoh: Tanggal, No. Bukti, Kode Akun, Kode Lawan Akun, Nama Akun, Uraian, Debet, Kredit, Saldo, Nama Vendor).
- **YYYY-MM** (rincian setiap bulan, termasuk saldo awal dan saldo berjalan).
- **Rekonsiliasi** (perbandingan angka PDF dan angka olahan).
- **Perlu Review** (transaksi yang belum diverifikasi).
- **Kode Akun** (COA yang digunakan).

### Kebijakan klasifikasi

- Prioritas 1: transaksi **identik** dengan contoh `Bank(1).xlsx`.
- Prioritas 2: aturan kata kunci yang Anda tambahkan.
- Prioritas 3: transaksi yang relatif eksplisit (biaya administrasi bank → 901; cicilan BCA Finance → 203).
- `SETORAN TUNAI` diberi **usulan** 121, `TARIKAN` usulan 122, tetapi tetap **perlu verifikasi**. Data bulan lain pada template bisa berbeda klasifikasinya.
- `PENERIMAAN NEGARA` tanpa kecocokan transaksi identik **tidak otomatis menjadi satu kode pajak**.
- Jenis transaksi yang tidak bisa dipastikan diberi kode kosong dan ditandai **Perlu Review**.

### Rekonsiliasi

Per transaksi, aplikasi menghitung: **saldo awal + debet – kredit**. Untuk setiap PDF, hasilnya dibandingkan dengan saldo tercetak bila ada, mutasi CR/DB di ringkasan rekening koran, dan saldo akhirnya. Jika terdapat selisih, aplikasi menampilkan peringatan sebelum ekspor.

## 4. Deploy gratis ke Streamlit Community Cloud

1. Buat **repository GitHub** (disarankan **private**).
2. Unggah **hanya file kode dan konfigurasi** dari folder `convertRK`, tanpa PDF, contoh Buku Bank, atau rekening koran. File `.gitignore` membantu mencegah file keuangan ikut ter-commit.
3. Buka https://share.streamlit.io dan masuk menggunakan GitHub.
4. Pilih **Create app**, pilih repository dan branch Anda, lalu isi **Main file path: `app.py`**.
5. Pilih **Deploy**. Platform akan memasang pustaka dari `requirements.txt`.
6. Gunakan upload `Kode Akun.xlsx` / `Bank(1).xlsx` melalui antarmuka aplikasi bila diperlukan. Aplikasi menyimpan data unggahan di RAM sesi, bukan ke file database.

> **Penting:** Hosting pihak ketiga memproses dokumen di server. Untuk rekening koran yang sangat sensitif, jalankan aplikasi secara lokal di Mac. Jangan commit file bank atau contoh transaksi yang sebenarnya ke repository, khususnya bila publik.

Dokumentasi resmi Streamlit:
- https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/deploy
- https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/app-dependencies

## 5. Batasan versi ini

- Ditujukan untuk **format e-statement BCA serupa contoh**, bukan otomatis semua format bank.
- PDF harus memiliki teks yang bisa disalin/diseleksi. PDF scan/gambar **belum didukung OCR**.
- Maksimal 50 PDF, total 60 MB data PDF, dan 20 MB per PDF.
- Beberapa PDF dari **rekening dan periode yang sama** tidak digabung otomatis; gabungkan seluruh halamannya menjadi satu PDF lebih dahulu.
- Hasil klasifikasi akun perlu pengecekan terhadap bukti pendukung dan substansi transaksi.
- Aplikasi tidak perlu akses internet untuk parsing/ekspor ketika dijalankan lokal (setelah dependensi terpasang).
