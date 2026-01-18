# 📑 Faktur Pajak Tool Pro (Web Version)

Aplikasi berbasis web untuk otomasi pengolahan PDF Faktur Pajak. Didesain khusus untuk akuntan dan praktisi perpajakan agar dapat mengelola dokumen faktur dengan cepat, baik melalui PC maupun perangkat Mobile (Android/iOS).

## ✨ Fitur Utama
- **🔄 Smart Rename**: Mengubah nama file PDF secara otomatis berdasarkan nomor referensi (PJ) yang dideteksi di dalam dokumen, ditambah dengan 3 kata terakhir dari nama file asli.
- **📁 Folder Classification**: Mengelompokkan file PDF ke dalam folder-folder berdasarkan kategori warna yang ada di database Excel Anda. Hasilnya dibungkus dalam satu file ZIP yang rapi.
- **🔗 PDF Merger**: Menggabungkan banyak file PDF menjadi satu dokumen tunggal secara berurutan (A-Z).
- **📊 Live Preview & Status**: Melihat tabel pratinjau hasil ekstraksi secara real-time sebelum melakukan download, lengkap dengan indikator status ✅ Berhasil atau ❌ Gagal.
- **🚀 Mobile Friendly**: Layout responsif (Wide Mode) yang nyaman digunakan di browser HP.
- **🧹 Clear All**: Fitur sekali klik untuk membersihkan antrean file tanpa harus menghapus satu per satu.

## 🛠️ Teknologi yang Digunakan
- **Python** (Core Logic)
- **Streamlit** (Web Framework & UI)
- **pdfplumber** (PDF Text Extraction)
- **PyPDF2** (PDF Merging)
- **Openpyxl & Pandas** (Excel & Data Management)
