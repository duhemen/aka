---

```markdown
<div align="center">

# ⚖️ AKA Validator Pro 2026
### *Agentic Knowledge Audit for Indonesian Public Procurement*

[![Python Version](https://img.shields.io/badge/Python-3.9%2B-blue?style=for-the-badge&logo=python)](https://python.org)
[![Streamlit](https://img.shields.io/badge/Streamlit-1.30%2B-FF4B4B?style=for-the-badge&logo=streamlit)](https://streamlit.io)
[![License](https://img.shields.io/badge/License-MIT-green?style=for-the-badge)](LICENSE)
[![Status](https://img.shields.io/badge/Status-Production%20Ready-success?style=for-the-badge)](https://github.com)

<br>

> **Transformasi Digital Pengadaan Publik.**  
> Dari RAB Penawar yang berantakan, menjadi Laporan Audit & Surat Dinas Resmi dalam hitungan detik.  
> **Sesuai Perpres 12/2021 & SE PUPR 07/2023.**

<br>

[🚀 Mulai Sekarang](#-cara-pakai-untuk-panitia---3-menit) &nbsp;·&nbsp; [📖 Dokumentasi](#-untuk-developer) &nbsp;·&nbsp; [📷 Demo](#-demo)

</div>

---

## ✨ Mengapa AKA Validator Pro?

Di dunia pengadaan, kesalahan aritmatika pada RAB sering kali menjadi biang keladi gugurnya penawaran atau kerugian negara. AKA hadir sebagai **AI Agent** yang bekerja di *background* Anda untuk menyelesaikan masalah ini secara instan.

| Fitur Unggulan | Deskripsi Fungsional |
| :--- | :--- |
| 🧠 **Agentic Layout Mapping** | Bisa membaca RAB meskipun format kolomnya acak (Vol vs Qty, Harga vs Unit Price). |
| 🔢 **Koreksi Aritmatik Otomatis** | Mendeteksi salah ketik rumus, item siluman (Volume 0 tapi ada harga), dan harga nol. |
| 🚦 **Evaluasi Kewajaran Harga (Traffic Light)** | Otomatis membandingkan dengan HPS. <80% (Merah), >110% (Oranye), Wajar (Hijau). |
| 📝 **Generator Surat Dinas (.docx)** | Bukan sekadar teks. Menghasilkan Surat Klarifikasi atau Surat Negosiasi siap tanda tangan. |
| 💎 **Preservasi Format Asli** | Mengubah warna merah pada sel yang salah, menambahkan *Excel Comment*, tanpa merusak tata letak asli penawar. |

---

## 🚀 Cara Pakai (Untuk Panitia - 3 Menit)

AKA dirancang agar **siapa pun bisa menggunakannya**, bahkan tanpa latar belakang pemrograman.

### 1. Persiapan Awal (Hanya Sekali)
Pastikan Python sudah terinstall di komputer Anda, lalu buka terminal/CMD dan jalankan:

```bash
pip install streamlit pandas openpyxl python-docx
```

### 2. Menjalankan Aplikasi
Buka folder tempat Anda menyimpan `app.py`, lalu jalankan perintah ini:

```bash
streamlit run app.py
```

### 3. Mulai Audit RAB (Alur Kerja)
1. 🌐 Browser akan terbuka otomatis di `http://localhost:8501`.
2. 📂 **Upload** file RAB Penawar (format `.xlsx`).
3. 💰 **Input** Total HPS Dinas (sesuai dokumen lelang).
4. 🤖 **Klik** tombol `🚀 Mulai Audit`.
5. 📥 **Download** hasil:

   - **Laporan Excel (BA + Audit)** → Berisi *Berita Acara*, *Audit Trail*, dan *RAB yang ditandai merah*.
   - **Surat Dinas (.docx)** → Akan muncul otomatis jika statusnya "Tidak Wajar" atau "Indikasi Kemahalan".

---

## 📊 Cara Membaca Hasil (Dashboard)

### 🟢 Status Kewajaran
Sistem akan menampilkan kartu warna sesuai aturan:

| Warna | Status | Tindak Lanjut |
| :--- | :--- | :--- |
| <span style="color: #EF4444;">🔴 Merah</span> | **TIDAK WAJAR** (< 80% HPS) | **WAJIB Klarifikasi.** Penawar harus membuktikan harga. Jika gagal, **GUGUR**. |
| <span style="color: #F59E0B;">🟠 Oranye</span> | **INDIKASI KEMAHALAN** (> 110% HPS) | **WAJIB Negosiasi** untuk efisiensi anggaran. |
| <span style="color: #F59E0B;">🟠 Oranye</span> | **WAJAR DENGAN CATATAN** | Ada temuan aritmatik. **Harga Kontrak harus menggunakan nilai terkoreksi**. |
| <span style="color: #10B981;">🟢 Hijau</span> | **WAJAR** | Tidak ada masalah. Lanjut ke evaluasi teknis. |

### 📋 Daftar Temuan
Di tabel rincian, baris yang berwarna merah muda adalah item yang error. Kolom **"Analisa"** akan menjelaskan akar masalahnya:
*   `SALAH KETIK/RUMUS` (Contoh: 2 x 50.000, tapi total tertulis 1.000.000).
*   `ITEM SILUMAN` (Volume 0 tapi ada harga jutaan rupiah).
*   `HARGA NOL` (Ada volume, tapi harga satuannya 0).

---

## ⚙️ Untuk Developer & Kontributor

AKA dibangun menggunakan arsitektur Python modern dan modular, sehingga sangat mudah untuk dikembangkan.

### Teknologi yang Digunakan
*   **Frontend & UI:** `Streamlit` (Interaktif & Ringan).
*   **Core Processing:** `Pandas` & `NumPy` (Analisis Data cepat).
*   **Excel Manipulation:** `openpyxl` (Mempertahankan styling & format asli).
*   **Document Generator:** `python-docx` (Pembuatan Surat Dinas resmi).

### Struktur Folder
```
aka/
├── app.py                  # Aplikasi Utama (v3.9.4)
├── requirements.txt        # Daftar dependency
└── README.md               # Dokumentasi ini
```

### Instalasi Developer
```bash
git clone https://github.com/username/aka-validator-pro.git
cd aka-validator-pro
pip install -r requirements.txt
```

---

## 📜 Basis Hukum

Fitur evaluasi harga pada AKA Validator Pro didasari oleh peraturan perundang-undangan yang berlaku di Indonesia:

*   **Perpres No. 12 Tahun 2021** tentang Perubahan atas Perpres No. 16 Tahun 2018 tentang Pengadaan Barang/Jasa Pemerintah (Pasal 48 & 60).
*   **SE Menteri PUPR No. 07/SE/M/2023** tentang Pedoman Pengadaan Barang/Jasa di Lingkungan Kementerian PUPR.

---

<div align="center">

**Dibuat dengan 💜 Ingat LUCA**  
*Solusi Cerdas untuk Pengadaan Publik yang Lebih Transparan*

[⬆ Kembali ke Atas](#-aka-validator-pro-2026)

</div>
```

---