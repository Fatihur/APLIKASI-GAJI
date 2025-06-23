# Excel to PDF Converter

Aplikasi Python untuk mengkonversi file Excel ke PDF dengan fitur bulk conversion dan pemilihan sheet.

## Fitur Utama

- ✅ **Bulk Conversion**: Konversi multiple sheets sekaligus
- ✅ **Sheet Selection**: Pilih sheet mana saja yang ingin dikonversi
- ✅ **Preserve Formatting**: Mempertahankan format asli Excel (font, warna, alignment)
- ✅ **User-Friendly GUI**: Interface yang mudah digunakan
- ✅ **Progress Tracking**: Monitor progress konversi
- ✅ **Auto Layout**: Otomatis menyesuaikan orientasi halaman berdasarkan jumlah kolom
- ✅ **Output Location**: Pilih lokasi penyimpanan file PDF
- ✅ **Professional Layout**: Hasil PDF yang rapi dan mudah dibaca
- ✅ **Capture Method**: Capture Excel sheet persis seperti aslinya (format portrait, satu halaman)
- ✅ **Dual Conversion Mode**: Pilih antara Capture Method atau Table Conversion
- ✅ **Multiple Files Support**: Proses beberapa file Excel sekaligus
- ✅ **Folder Per File**: Setiap file Excel mendapat folder terpisah dengan nama yang bisa disesuaikan
- ✅ **Auto Sheet Filtering**: Otomatis mengabaikan sheet 1-9

## Persyaratan Sistem

- Python 3.7 atau lebih baru
- Windows (untuk Capture Method, memerlukan Microsoft Excel)
- Linux/macOS (hanya Table Conversion method)
- Packages yang diperlukan (lihat requirements.txt)

## Instalasi

1. Clone atau download repository ini
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Cara Penggunaan

1. Jalankan aplikasi:
   ```bash
   python main.py
   ```

2. **Add Excel Files**:
   - Klik tombol "Add Files" untuk memilih multiple file Excel (.xlsx atau .xls)
   - Gunakan "Remove Selected" untuk menghapus file yang dipilih
   - Gunakan "Clear All" untuk menghapus semua file
   - Klik "Set Folder Name" untuk mengatur nama folder custom per file

3. **Pilih Output Directory** (Opsional):
   - Klik tombol "Browse" di bagian Output Directory untuk memilih lokasi penyimpanan
   - Jika tidak dipilih, file akan disimpan di direktori yang sama dengan file Excel

4. **Pilih Sheet** (Opsional):
   - Klik pada file di daftar untuk melihat sheet-nya
   - Pilih sheet tertentu yang ingin dikonversi (multiple selection)
   - Jika tidak ada sheet yang dipilih, semua sheet dari semua file akan dikonversi
   - Sheet 1-9 otomatis diabaikan
   - Gunakan "Select All" untuk memilih semua sheet
   - Gunakan "Clear All" untuk membatalkan pilihan

5. **Atur Opsi Konversi**:
   - **Bulk Mode**: Jika dicentang, setiap sheet akan menjadi file PDF terpisah
   - **Preserve Original Formatting**: Mempertahankan format asli Excel
   - **Conversion Method**:
     - **Capture Method (Recommended)**: Capture Excel sheet persis seperti aslinya
     - **Table Conversion**: Konversi data ke tabel PDF

6. **Konversi**:
   - Klik tombol "Convert to PDF"
   - Monitor progress melalui progress bar
   - File PDF akan disimpan di lokasi yang telah ditentukan

## Struktur File Output

Aplikasi akan membuat folder terpisah untuk setiap file Excel:
```
output_directory/
├── [folder_name_file1]/
│   ├── [sheet1].pdf
│   ├── [sheet2].pdf
│   └── [sheet3].pdf
├── [folder_name_file2]/
│   ├── [sheet1].pdf
│   └── [sheet2].pdf
└── [folder_name_file3]/
    └── [sheet1].pdf
```

Contoh:
```
output/
├── sample_data/
│   ├── Data Karyawan.pdf
│   ├── Laporan Penjualan.pdf
│   └── Summary.pdf
├── company_report/
│   ├── Data Karyawan.pdf
│   └── Summary.pdf
└── financial_data/
    └── Balance Sheet.pdf
```

## Fitur Formatting yang Didukung

- ✅ Font bold
- ✅ Font size
- ✅ Background color
- ✅ Text alignment (left, center, right)
- ✅ Borders
- ✅ Merged cells (sebagian)
- ✅ Number formatting

## File Testing

Gunakan script `create_sample_excel.py` untuk membuat file Excel contoh:
```bash
python create_sample_excel.py
```

File `sample_data.xlsx` akan dibuat dengan 3 sheet:
1. **Data Karyawan**: Tabel data karyawan dengan formatting
2. **Laporan Penjualan**: Data penjualan dengan formula
3. **Summary**: Ringkasan dengan referensi ke sheet lain

### Test Konversi Otomatis

Untuk test konversi secara otomatis tanpa GUI:
```bash
python test_conversion.py
```

Script ini akan:
- Membaca file `sample_data.xlsx`
- Mengkonversi semua sheet ke PDF
- Menyimpan hasil di direktori `output_pdfs/`
- Menampilkan informasi ukuran file hasil

### Test Capture Method

Untuk test capture method secara khusus:
```bash
python test_capture.py
```

Script ini akan test:
- Excel capture functionality
- PDF converter dengan capture method
- Bulk conversion dengan capture method

## Struktur Project

```
├── main.py                 # File utama aplikasi GUI
├── excel_reader.py         # Modul untuk membaca file Excel
├── pdf_converter.py        # Modul untuk konversi ke PDF
├── create_sample_excel.py  # Script untuk membuat file testing
├── test_conversion.py      # Script untuk test konversi otomatis
├── requirements.txt        # Dependencies
├── sample_data.xlsx        # File Excel contoh
├── output_pdfs/           # Direktori hasil konversi (dibuat otomatis)
└── README.md              # Dokumentasi
```

## Troubleshooting

### Error saat install dependencies
Jika terjadi error saat install pandas, coba install tanpa pandas:
```bash
pip install openpyxl reportlab Pillow xlsxwriter
```

### File Excel tidak terbaca
- Pastikan file Excel tidak sedang dibuka di aplikasi lain
- Pastikan format file adalah .xlsx atau .xls
- Cek apakah file tidak corrupt

### PDF tidak sesuai layout
- Untuk tabel dengan banyak kolom, aplikasi otomatis menggunakan landscape orientation
- Jika masih tidak muat, pertimbangkan untuk membagi data ke multiple sheets

### Performance untuk file besar
- Untuk file Excel dengan data sangat besar (>10,000 rows), konversi mungkin memakan waktu lama
- Pertimbangkan untuk membagi data ke multiple sheets yang lebih kecil

## Kontribusi

Silakan buat issue atau pull request untuk perbaikan dan penambahan fitur.

## Lisensi

MIT License - Silakan gunakan dan modifikasi sesuai kebutuhan.
