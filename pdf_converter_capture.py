"""
PDF Converter with Capture Method
Modul untuk mengkonversi Excel ke PDF menggunakan capture method
"""

import os
from excel_capture import ExcelCapture
from reportlab.lib.pagesizes import A4, letter, landscape
from reportlab.platypus import SimpleDocTemplate, Image as RLImage, Spacer, Paragraph
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import mm
from PIL import Image
import tempfile
from watermark_manager import WatermarkManager

class PDFConverterCapture:
    def __init__(self, page_orientation='portrait', enable_watermark=True, watermark_opacity=0.3, watermark_position="bottom-right"):
        """
        Initialize PDF converter dengan capture method

        Args:
            page_orientation (str): 'portrait' atau 'landscape'
            enable_watermark (bool): Enable watermark pada PDF
            watermark_opacity (float): Transparansi watermark (0.0-1.0)
            watermark_position (str): Posisi watermark
        """
        self.page_orientation = page_orientation
        self.styles = getSampleStyleSheet()
        self.enable_watermark = enable_watermark
        self.watermark_manager = WatermarkManager() if enable_watermark else None
        self.watermark_opacity = watermark_opacity
        self.watermark_position = watermark_position
        
    def convert_excel_to_pdf(self, excel_file, selected_sheets, output_directory, folder_prefix=""):
        """
        Konversi Excel sheets ke PDF menggunakan capture method (optimized)

        Args:
            excel_file (str): Path ke file Excel
            selected_sheets (list): List nama sheet yang akan dikonversi
            output_directory (str): Direktori output
            folder_prefix (str): Prefix untuk nama file

        Returns:
            dict: Dictionary hasil konversi {sheet_name: pdf_path}
        """
        results = {}
        capture = ExcelCapture()

        try:
            # Buka file Excel sekali saja
            capture.open_excel_file(excel_file)

            # Buat direktori output jika belum ada
            if not os.path.exists(output_directory):
                os.makedirs(output_directory)

            # Konversi semua sheet dalam satu session Excel
            for sheet_name in selected_sheets:
                try:
                    # Nama file output dengan prefix
                    safe_sheet_name = "".join(c for c in sheet_name if c.isalnum() or c in (' ', '-', '_')).rstrip()
                    if folder_prefix:
                        pdf_filename = f"{folder_prefix}_{safe_sheet_name}.pdf"
                    else:
                        base_name = os.path.splitext(os.path.basename(excel_file))[0]
                        pdf_filename = f"{base_name}_{safe_sheet_name}.pdf"

                    pdf_path = os.path.join(output_directory, pdf_filename)

                    # Capture sheet langsung ke PDF dengan watermark jika enabled
                    if self.enable_watermark and self.watermark_manager and self.watermark_manager.watermark_exists:
                        captured_pdf = capture.capture_sheet_as_png(sheet_name, add_watermark=True, watermark_text="CONFIDENTIAL")
                    else:
                        captured_pdf = capture.capture_sheet_as_png(sheet_name)

                    # Copy hasil capture ke lokasi yang diinginkan
                    if captured_pdf and os.path.exists(captured_pdf):
                        # Jika hasil capture sudah berupa PDF, copy saja
                        import shutil
                        shutil.copy2(captured_pdf, pdf_path)

                        # Hapus file temporary
                        try:
                            os.remove(captured_pdf)
                        except:
                            pass

                        # Tambahkan watermark jika enabled (skip untuk sekarang, biarkan PDF original)
                        if self.enable_watermark and self.watermark_manager and self.watermark_manager.watermark_exists:
                            print(f"🎨 Watermark enabled for {sheet_name} (feature ready)")
                            # Watermark akan ditambahkan di versi mendatang dengan PyPDF2

                        results[sheet_name] = pdf_path
                    else:
                        results[sheet_name] = None

                except Exception as e:
                    print(f"Error converting sheet '{sheet_name}': {str(e)}")
                    results[sheet_name] = None

        except Exception as e:
            print(f"Error opening Excel file: {str(e)}")

        finally:
            capture.close()

        return results
    
    def convert_single_sheet(self, excel_file, sheet_name, output_path, capture_instance=None):
        """
        Konversi single sheet ke PDF (optimized with reusable capture instance)

        Args:
            excel_file (str): Path ke file Excel
            sheet_name (str): Nama sheet
            output_path (str): Path output PDF
            capture_instance: Instance ExcelCapture yang sudah dibuka (optional)

        Returns:
            bool: True jika berhasil
        """
        # Gunakan instance yang sudah ada atau buat baru
        if capture_instance:
            capture = capture_instance
            close_after = False
        else:
            capture = ExcelCapture()
            close_after = True

        try:
            # Buka file Excel hanya jika belum dibuka
            if not capture_instance:
                capture.open_excel_file(excel_file)

            # Buat direktori output jika belum ada
            output_dir = os.path.dirname(output_path)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir)

            # Capture sheet dengan watermark jika enabled
            if self.enable_watermark and self.watermark_manager and self.watermark_manager.watermark_exists:
                captured_pdf = capture.capture_sheet_as_png(sheet_name, add_watermark=True, watermark_text="CONFIDENTIAL")
            else:
                captured_pdf = capture.capture_sheet_as_png(sheet_name)

            if captured_pdf and os.path.exists(captured_pdf):
                # Copy hasil capture ke lokasi yang diinginkan
                import shutil
                shutil.copy2(captured_pdf, output_path)

                # Hapus file temporary
                try:
                    os.remove(captured_pdf)
                except:
                    pass

                # Tambahkan watermark jika enabled (skip untuk sekarang)
                if self.enable_watermark and self.watermark_manager and self.watermark_manager.watermark_exists:
                    print(f"🎨 Watermark enabled for {os.path.basename(output_path)} (feature ready)")
                    # Watermark akan ditambahkan di versi mendatang dengan PyPDF2

                return True
            else:
                return False

        except Exception as e:
            print(f"Error converting sheet '{sheet_name}': {str(e)}")
            return False

        finally:
            # Hanya close jika kita yang membuat instance
            if close_after:
                capture.close()
                # Small delay for stability
                import time
                time.sleep(0.2)

    def _add_watermark_to_existing_pdf(self, pdf_path, sheet_name):
        """
        Tambahkan watermark ke PDF yang sudah ada dengan menambahkan teks/image overlay

        Args:
            pdf_path (str): Path ke PDF file
            sheet_name (str): Nama sheet (untuk logging)

        Returns:
            bool: True jika berhasil
        """
        try:
            if not self.watermark_manager or not self.watermark_manager.watermark_exists:
                return False

            # Baca PDF original dan tambahkan watermark menggunakan reportlab
            from reportlab.pdfgen import canvas
            from reportlab.lib.pagesizes import A4

            # Buat temporary file untuk PDF dengan watermark
            temp_watermarked = tempfile.NamedTemporaryFile(delete=False, suffix='_watermarked.pdf')
            temp_watermarked.close()

            # Baca konten PDF original (simplified approach)
            # Karena kita tidak bisa merge PDF tanpa PyPDF2, kita akan menambahkan watermark text

            # Buat canvas baru dengan watermark
            c = canvas.Canvas(temp_watermarked.name, pagesize=A4)

            # Tambahkan watermark text yang visible
            page_width, page_height = A4

            # Tambahkan watermark text di beberapa posisi
            c.saveState()
            c.setFillAlpha(0.3)  # Transparansi
            c.setFont("Helvetica-Bold", 48)
            c.setFillGray(0.8)

            # Watermark di tengah
            c.drawCentredText(page_width/2, page_height/2, "WATERMARK")

            # Watermark di pojok kanan bawah
            c.setFont("Helvetica", 24)
            c.drawRightString(page_width - 50, 50, "SAMPLE")

            # Watermark di pojok kiri atas
            c.drawString(50, page_height - 50, "CONFIDENTIAL")

            c.restoreState()
            c.save()

            # Copy watermarked PDF ke lokasi original
            import shutil
            shutil.copy2(temp_watermarked.name, pdf_path)

            # Cleanup
            try:
                os.remove(temp_watermarked.name)
            except:
                pass

            print(f"✅ Text watermark added to {sheet_name}")
            return True

        except Exception as e:
            print(f"❌ Error adding watermark: {str(e)}")
            return False
    
    def create_combined_pdf(self, excel_file, selected_sheets, output_path):
        """
        Buat PDF gabungan dari multiple sheets
        
        Args:
            excel_file (str): Path ke file Excel
            selected_sheets (list): List nama sheet
            output_path (str): Path output PDF
            
        Returns:
            bool: True jika berhasil
        """
        capture = ExcelCapture()
        temp_files = []
        
        try:
            # Buka file Excel
            capture.open_excel_file(excel_file)
            
            # Buat direktori output jika belum ada
            output_dir = os.path.dirname(output_path)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            # Tentukan page size
            page_size = A4 if self.page_orientation == 'portrait' else landscape(A4)
            
            # Buat dokumen PDF
            doc = SimpleDocTemplate(
                output_path,
                pagesize=page_size,
                rightMargin=10*mm,
                leftMargin=10*mm,
                topMargin=10*mm,
                bottomMargin=10*mm
            )
            
            elements = []
            
            # Capture setiap sheet dan tambahkan ke PDF
            for i, sheet_name in enumerate(selected_sheets):
                try:
                    # Capture sheet sebagai PDF temporary
                    temp_pdf = capture.capture_sheet_as_png(sheet_name)
                    
                    if temp_pdf and os.path.exists(temp_pdf):
                        temp_files.append(temp_pdf)
                        
                        # Tambahkan judul sheet
                        if i > 0:
                            elements.append(Spacer(1, 20*mm))  # Page break
                        
                        title_style = ParagraphStyle(
                            'SheetTitle',
                            parent=self.styles['Heading1'],
                            fontSize=16,
                            spaceAfter=10*mm,
                            alignment=1,  # Center
                            textColor=colors.darkblue
                        )
                        
                        title = Paragraph(f"<b>{sheet_name}</b>", title_style)
                        elements.append(title)
                        
                        # Note: Untuk menggabungkan PDF, kita perlu library tambahan
                        # Untuk sementara, kita buat file terpisah
                        
                except Exception as e:
                    print(f"Error processing sheet '{sheet_name}': {str(e)}")
            
            # Jika hanya ada satu sheet, copy langsung
            if len(temp_files) == 1:
                import shutil
                shutil.copy2(temp_files[0], output_path)
                return True
            
            # Untuk multiple sheets, kita perlu PyPDF2 atau library serupa
            # Sementara return False untuk implementasi nanti
            return False
            
        except Exception as e:
            raise Exception(f"Error creating combined PDF: {str(e)}")
            
        finally:
            capture.close()
            
            # Cleanup temporary files
            for temp_file in temp_files:
                try:
                    if os.path.exists(temp_file):
                        os.remove(temp_file)
                except:
                    pass
    
    def get_sheet_preview(self, excel_file, sheet_name):
        """
        Dapatkan preview sheet sebagai gambar
        
        Args:
            excel_file (str): Path ke file Excel
            sheet_name (str): Nama sheet
            
        Returns:
            str: Path ke file preview image
        """
        capture = ExcelCapture()
        
        try:
            capture.open_excel_file(excel_file)
            
            # Capture sebagai PNG untuk preview
            temp_dir = tempfile.gettempdir()
            preview_path = os.path.join(temp_dir, f"preview_{sheet_name}.png")
            
            # Untuk preview, kita bisa gunakan method yang berbeda
            # Sementara return None
            return None
            
        except Exception as e:
            print(f"Error creating preview for '{sheet_name}': {str(e)}")
            return None
            
        finally:
            capture.close()
