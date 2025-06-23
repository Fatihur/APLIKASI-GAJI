"""
Excel Capture Module
Modul untuk capture Excel sheet sebagai gambar menggunakan xlwings
"""

import xlwings as xw
import os
import tempfile
from PIL import Image
import time

class ExcelCapture:
    def __init__(self):
        """Initialize Excel capture"""
        self.app = None
        self.workbook = None
        
    def open_excel_file(self, file_path):
        """
        Buka file Excel menggunakan xlwings
        
        Args:
            file_path (str): Path ke file Excel
        """
        try:
            # Buka Excel application (hidden)
            self.app = xw.App(visible=False, add_book=False)
            
            # Buka workbook
            self.workbook = self.app.books.open(file_path)
            
            return True
            
        except Exception as e:
            self.close()
            raise Exception(f"Error opening Excel file: {str(e)}")
    
    def get_sheet_names(self):
        """
        Dapatkan daftar nama sheet
        
        Returns:
            list: List nama sheet
        """
        if not self.workbook:
            raise Exception("No workbook opened")
            
        return [sheet.name for sheet in self.workbook.sheets]
    
    def capture_sheet_as_image(self, sheet_name, output_path=None):
        """
        Capture sheet sebagai gambar
        
        Args:
            sheet_name (str): Nama sheet yang akan di-capture
            output_path (str): Path output gambar (optional)
            
        Returns:
            str: Path ke file gambar hasil capture
        """
        if not self.workbook:
            raise Exception("No workbook opened")
            
        try:
            # Pilih sheet
            sheet = self.workbook.sheets[sheet_name]
            sheet.activate()
            
            # Tunggu sebentar untuk memastikan sheet ter-load
            time.sleep(0.5)
            
            # Dapatkan used range (area yang berisi data)
            used_range = sheet.used_range
            
            if not used_range:
                raise Exception(f"Sheet '{sheet_name}' is empty")
            
            # Set output path jika tidak diberikan
            if not output_path:
                temp_dir = tempfile.gettempdir()
                safe_sheet_name = "".join(c for c in sheet_name if c.isalnum() or c in (' ', '-', '_')).rstrip()
                output_path = os.path.join(temp_dir, f"excel_capture_{safe_sheet_name}.png")
            
            # Capture used range sebagai gambar
            used_range.api.CopyPicture(Format=2)  # xlBitmap format
            
            # Paste ke worksheet baru untuk export
            temp_sheet = self.workbook.sheets.add('TempCapture')
            temp_sheet.activate()
            
            # Paste gambar
            temp_sheet.api.Paste()
            
            # Export sebagai gambar
            temp_sheet.api.ExportAsFixedFormat(
                Type=0,  # xlTypePDF
                Filename=output_path.replace('.png', '.pdf'),
                Quality=0,  # xlQualityStandard
                IncludeDocProperties=True,
                IgnorePrintAreas=False,
                OpenAfterPublish=False
            )
            
            # Hapus temporary sheet
            temp_sheet.delete()
            
            return output_path.replace('.png', '.pdf')
            
        except Exception as e:
            raise Exception(f"Error capturing sheet '{sheet_name}': {str(e)}")
    
    def capture_sheet_as_png(self, sheet_name, output_path=None):
        """
        Capture sheet sebagai PNG menggunakan metode alternatif
        
        Args:
            sheet_name (str): Nama sheet yang akan di-capture
            output_path (str): Path output gambar (optional)
            
        Returns:
            str: Path ke file gambar hasil capture
        """
        if not self.workbook:
            raise Exception("No workbook opened")
            
        try:
            # Pilih sheet
            sheet = self.workbook.sheets[sheet_name]
            sheet.activate()
            
            # Tunggu sebentar untuk memastikan sheet ter-load
            time.sleep(0.5)
            
            # Dapatkan used range
            used_range = sheet.used_range
            
            if not used_range:
                raise Exception(f"Sheet '{sheet_name}' is empty")
            
            # Set output path jika tidak diberikan
            if not output_path:
                temp_dir = tempfile.gettempdir()
                safe_sheet_name = "".join(c for c in sheet_name if c.isalnum() or c in (' ', '-', '_')).rstrip()
                output_path = os.path.join(temp_dir, f"excel_capture_{safe_sheet_name}.png")
            
            # Fit sheet ke satu halaman
            sheet.api.PageSetup.Zoom = False
            sheet.api.PageSetup.FitToPagesWide = 1
            sheet.api.PageSetup.FitToPagesTall = 1
            sheet.api.PageSetup.Orientation = 1  # Portrait
            
            # Export sebagai PDF terlebih dahulu
            pdf_path = output_path.replace('.png', '.pdf')
            sheet.api.ExportAsFixedFormat(
                Type=0,  # xlTypePDF
                Filename=pdf_path,
                Quality=0,  # xlQualityStandard
                IncludeDocProperties=False,
                IgnorePrintAreas=False,
                OpenAfterPublish=False
            )
            
            return pdf_path
            
        except Exception as e:
            raise Exception(f"Error capturing sheet '{sheet_name}': {str(e)}")
    
    def capture_all_sheets(self, output_directory):
        """
        Capture semua sheet dalam workbook
        
        Args:
            output_directory (str): Direktori output
            
        Returns:
            dict: Dictionary dengan nama sheet sebagai key dan path file sebagai value
        """
        if not self.workbook:
            raise Exception("No workbook opened")
            
        # Buat direktori output jika belum ada
        if not os.path.exists(output_directory):
            os.makedirs(output_directory)
            
        results = {}
        sheet_names = self.get_sheet_names()
        
        for sheet_name in sheet_names:
            try:
                safe_sheet_name = "".join(c for c in sheet_name if c.isalnum() or c in (' ', '-', '_')).rstrip()
                output_path = os.path.join(output_directory, f"{safe_sheet_name}.pdf")
                
                captured_path = self.capture_sheet_as_png(sheet_name, output_path)
                results[sheet_name] = captured_path
                
            except Exception as e:
                print(f"Error capturing sheet '{sheet_name}': {str(e)}")
                results[sheet_name] = None
        
        return results
    
    def close(self):
        """Tutup Excel application"""
        try:
            if self.workbook:
                self.workbook.close()
                self.workbook = None
                
            if self.app:
                self.app.quit()
                self.app = None
                
        except Exception as e:
            print(f"Error closing Excel: {str(e)}")
    
    def __del__(self):
        """Destructor untuk memastikan Excel tertutup"""
        self.close()
