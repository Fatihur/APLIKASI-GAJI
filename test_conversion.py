"""
Script untuk test konversi Excel ke PDF secara otomatis
"""

import os
from excel_reader import ExcelReader
from pdf_converter import PDFConverter

def test_conversion():
    """Test konversi file sample_data.xlsx"""
    
    excel_file = "sample_data.xlsx"
    
    if not os.path.exists(excel_file):
        print("File sample_data.xlsx tidak ditemukan!")
        print("Jalankan: python create_sample_excel.py")
        return
    
    try:
        # Baca file Excel
        print("📖 Membaca file Excel...")
        reader = ExcelReader(excel_file)
        sheets_info = reader.get_sheets_info()
        
        print(f"✅ Berhasil membaca {len(sheets_info)} sheets:")
        for sheet_name, info in sheets_info.items():
            print(f"   - {sheet_name}: {info['max_row']} rows, {info['max_col']} cols, {info['filled_cells']} filled cells")
        
        # Buat converter
        converter = PDFConverter(preserve_formatting=True, bulk_mode=True)
        
        # Buat direktori output
        output_dir = "output_pdfs"
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
            print(f"📁 Membuat direktori output: {output_dir}")
        
        # Konversi setiap sheet
        print("\n🔄 Memulai konversi...")
        for sheet_name in sheets_info.keys():
            print(f"   Converting: {sheet_name}...")
            
            # Nama file output
            safe_sheet_name = "".join(c for c in sheet_name if c.isalnum() or c in (' ', '-', '_')).rstrip()
            output_file = os.path.join(output_dir, f"sample_data_{safe_sheet_name}.pdf")
            
            # Konversi
            converter.convert_sheet_to_pdf(excel_file, sheet_name, output_file)
            print(f"   ✅ Berhasil: {output_file}")
        
        reader.close()
        
        print(f"\n🎉 Konversi selesai! File PDF tersimpan di direktori '{output_dir}'")
        print("\nFile yang dibuat:")
        for file in os.listdir(output_dir):
            if file.endswith('.pdf'):
                file_path = os.path.join(output_dir, file)
                file_size = os.path.getsize(file_path)
                print(f"   - {file} ({file_size:,} bytes)")
        
    except Exception as e:
        print(f"❌ Error: {str(e)}")

if __name__ == "__main__":
    test_conversion()
