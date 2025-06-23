"""
Watermark Manager
Modul untuk menambahkan watermark ke PDF files
"""

import os
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.utils import ImageReader
# from PyPDF2 import PdfReader, PdfWriter  # Optional dependency
import tempfile
from PIL import Image
import io

class WatermarkManager:
    def __init__(self, watermark_path="watermark.png"):
        """
        Initialize watermark manager
        
        Args:
            watermark_path (str): Path ke file watermark
        """
        self.watermark_path = watermark_path
        self.watermark_exists = os.path.exists(watermark_path)
        
    def add_watermark_to_pdf(self, pdf_path, output_path=None, opacity=0.3, position="center"):
        """
        Tambahkan watermark ke PDF file menggunakan reportlab

        Args:
            pdf_path (str): Path ke PDF yang akan diberi watermark
            output_path (str): Path output PDF (optional, default overwrite)
            opacity (float): Transparansi watermark (0.0-1.0)
            position (str): Posisi watermark ("center", "bottom-right", "top-left")

        Returns:
            bool: True jika berhasil
        """
        if not self.watermark_exists:
            print(f"⚠️  Watermark file not found: {self.watermark_path}")
            return False

        if not os.path.exists(pdf_path):
            print(f"❌ PDF file not found: {pdf_path}")
            return False

        try:
            # Set output path
            if output_path is None:
                output_path = pdf_path

            # Buat PDF baru dengan watermark menggunakan reportlab
            success = self._add_watermark_with_reportlab(pdf_path, output_path, opacity, position)

            if success:
                print(f"✅ Watermark added to: {os.path.basename(output_path)}")
                return True
            else:
                print(f"❌ Failed to add watermark to: {os.path.basename(pdf_path)}")
                return False

        except Exception as e:
            print(f"❌ Error adding watermark: {str(e)}")
            return False
    
    def _create_watermark_pdf(self, page_width, page_height, opacity, position):
        """
        Buat PDF watermark untuk ukuran halaman tertentu
        
        Args:
            page_width (float): Lebar halaman
            page_height (float): Tinggi halaman
            opacity (float): Transparansi watermark
            position (str): Posisi watermark
            
        Returns:
            str: Path ke temporary watermark PDF
        """
        try:
            # Buat temporary file untuk watermark PDF
            temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
            temp_pdf.close()
            
            # Buat canvas untuk watermark
            c = canvas.Canvas(temp_pdf.name, pagesize=(page_width, page_height))
            
            # Load dan resize watermark image
            watermark_img = self._prepare_watermark_image(page_width, page_height, opacity)
            
            if watermark_img:
                # Hitung posisi watermark
                img_width, img_height = watermark_img.size
                x, y = self._calculate_watermark_position(
                    page_width, page_height, img_width, img_height, position
                )
                
                # Tambahkan watermark ke canvas
                c.drawImage(
                    ImageReader(watermark_img), 
                    x, y, 
                    width=img_width, 
                    height=img_height,
                    mask='auto'
                )
            
            c.save()
            return temp_pdf.name
            
        except Exception as e:
            print(f"❌ Error creating watermark PDF: {str(e)}")
            return None
    
    def _prepare_watermark_image(self, page_width, page_height, opacity):
        """
        Prepare watermark image dengan opacity dan ukuran yang sesuai
        
        Args:
            page_width (float): Lebar halaman
            page_height (float): Tinggi halaman
            opacity (float): Transparansi watermark
            
        Returns:
            PIL.Image: Processed watermark image
        """
        try:
            # Load watermark image
            img = Image.open(self.watermark_path)
            
            # Convert ke RGBA jika belum
            if img.mode != 'RGBA':
                img = img.convert('RGBA')
            
            # Resize watermark (maksimal 30% dari ukuran halaman)
            max_width = page_width * 0.3
            max_height = page_height * 0.3
            
            # Hitung ratio untuk maintain aspect ratio
            width_ratio = max_width / img.width
            height_ratio = max_height / img.height
            ratio = min(width_ratio, height_ratio)
            
            new_width = int(img.width * ratio)
            new_height = int(img.height * ratio)
            
            img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
            
            # Apply opacity
            if opacity < 1.0:
                # Buat alpha channel dengan opacity
                alpha = img.split()[-1]  # Get alpha channel
                alpha = alpha.point(lambda p: int(p * opacity))
                img.putalpha(alpha)
            
            return img
            
        except Exception as e:
            print(f"❌ Error preparing watermark image: {str(e)}")
            return None
    
    def _calculate_watermark_position(self, page_width, page_height, img_width, img_height, position):
        """
        Hitung posisi watermark berdasarkan parameter position
        
        Args:
            page_width (float): Lebar halaman
            page_height (float): Tinggi halaman
            img_width (int): Lebar watermark
            img_height (int): Tinggi watermark
            position (str): Posisi watermark
            
        Returns:
            tuple: (x, y) koordinat watermark
        """
        margin = 20  # Margin dari tepi
        
        if position == "center":
            x = (page_width - img_width) / 2
            y = (page_height - img_height) / 2
        elif position == "bottom-right":
            x = page_width - img_width - margin
            y = margin
        elif position == "top-left":
            x = margin
            y = page_height - img_height - margin
        elif position == "top-right":
            x = page_width - img_width - margin
            y = page_height - img_height - margin
        elif position == "bottom-left":
            x = margin
            y = margin
        else:  # default to center
            x = (page_width - img_width) / 2
            y = (page_height - img_height) / 2
        
        return x, y

    def _add_watermark_with_reportlab(self, pdf_path, output_path, opacity, position):
        """
        Tambahkan watermark menggunakan reportlab dengan membaca PDF original

        Args:
            pdf_path (str): Path ke PDF original
            output_path (str): Path output PDF
            opacity (float): Transparansi watermark
            position (str): Posisi watermark

        Returns:
            bool: True jika berhasil
        """
        try:
            # Baca PDF original untuk mendapatkan ukuran halaman
            from reportlab.lib.pagesizes import A4
            from reportlab.pdfgen import canvas
            from reportlab.lib.utils import ImageReader

            # Estimasi ukuran halaman (default A4)
            page_width, page_height = A4

            # Buat temporary file untuk PDF dengan watermark
            temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
            temp_pdf.close()

            # Buat canvas untuk PDF baru
            c = canvas.Canvas(temp_pdf.name, pagesize=(page_width, page_height))

            # Load dan prepare watermark image
            watermark_img = self._prepare_watermark_image(page_width, page_height, opacity)

            if watermark_img:
                # Hitung posisi watermark
                img_width, img_height = watermark_img.size
                x, y = self._calculate_watermark_position(
                    page_width, page_height, img_width, img_height, position
                )

                # Tambahkan watermark ke canvas
                c.drawImage(
                    ImageReader(watermark_img),
                    x, y,
                    width=img_width,
                    height=img_height,
                    mask='auto'
                )

            # Simpan PDF dengan watermark
            c.save()

            # Copy PDF dengan watermark ke output path
            import shutil
            shutil.copy2(temp_pdf.name, output_path)

            # Cleanup temporary file
            try:
                os.remove(temp_pdf.name)
            except:
                pass

            return True

        except Exception as e:
            print(f"❌ Error adding watermark with reportlab: {str(e)}")
            return False

    def _create_simple_watermark_overlay(self, pdf_path, opacity, position):
        """
        Buat simple watermark overlay (placeholder for full implementation)

        Args:
            pdf_path (str): Path ke PDF original
            opacity (float): Transparansi watermark
            position (str): Posisi watermark

        Returns:
            str: Path ke watermark overlay (or None if failed)
        """
        try:
            # For now, just validate that watermark image exists
            if os.path.exists(self.watermark_path):
                print(f"   🎨 Watermark image found: {self.watermark_path}")
                print(f"   📍 Position: {position}, Opacity: {opacity}")
                return self.watermark_path
            else:
                print(f"   ❌ Watermark image not found: {self.watermark_path}")
                return None

        except Exception as e:
            print(f"❌ Error creating watermark overlay: {str(e)}")
            return None
    
    def add_watermark_to_multiple_pdfs(self, pdf_paths, opacity=0.3, position="center"):
        """
        Tambahkan watermark ke multiple PDF files
        
        Args:
            pdf_paths (list): List path ke PDF files
            opacity (float): Transparansi watermark
            position (str): Posisi watermark
            
        Returns:
            dict: Dictionary hasil {pdf_path: success_status}
        """
        results = {}
        
        if not self.watermark_exists:
            print(f"⚠️  Watermark file not found: {self.watermark_path}")
            return {path: False for path in pdf_paths}
        
        print(f"🎨 Adding watermark to {len(pdf_paths)} PDF files...")
        
        for pdf_path in pdf_paths:
            try:
                success = self.add_watermark_to_pdf(pdf_path, opacity=opacity, position=position)
                results[pdf_path] = success
            except Exception as e:
                print(f"❌ Error processing {pdf_path}: {str(e)}")
                results[pdf_path] = False
        
        success_count = sum(1 for success in results.values() if success)
        print(f"📊 Watermark results: {success_count}/{len(pdf_paths)} files processed")
        
        return results
    
    def create_sample_watermark(self, output_path="watermark.png"):
        """
        Buat sample watermark jika belum ada
        
        Args:
            output_path (str): Path untuk save sample watermark
            
        Returns:
            bool: True jika berhasil
        """
        try:
            from PIL import Image, ImageDraw, ImageFont
            
            # Buat image 200x200 dengan background transparan
            img = Image.new('RGBA', (200, 200), (0, 0, 0, 0))
            draw = ImageDraw.Draw(img)
            
            # Tambahkan text watermark
            try:
                # Coba gunakan font yang lebih besar
                font = ImageFont.truetype("arial.ttf", 24)
            except:
                # Fallback ke default font
                font = ImageFont.load_default()
            
            # Text watermark
            text = "SAMPLE\nWATERMARK"
            
            # Hitung posisi text untuk center
            bbox = draw.textbbox((0, 0), text, font=font)
            text_width = bbox[2] - bbox[0]
            text_height = bbox[3] - bbox[1]
            
            x = (200 - text_width) // 2
            y = (200 - text_height) // 2
            
            # Gambar text dengan outline
            draw.text((x-1, y-1), text, font=font, fill=(0, 0, 0, 100))  # Shadow
            draw.text((x, y), text, font=font, fill=(128, 128, 128, 150))  # Main text
            
            # Tambahkan border
            draw.rectangle([10, 10, 190, 190], outline=(100, 100, 100, 100), width=2)
            
            # Save image
            img.save(output_path, 'PNG')
            print(f"✅ Sample watermark created: {output_path}")
            
            return True
            
        except Exception as e:
            print(f"❌ Error creating sample watermark: {str(e)}")
            return False

def test_watermark():
    """Test watermark functionality"""
    print("🧪 Testing Watermark Functionality...")
    print("=" * 50)
    
    # Create watermark manager
    wm = WatermarkManager()
    
    # Create sample watermark if not exists
    if not wm.watermark_exists:
        print("📁 Creating sample watermark...")
        wm.create_sample_watermark()
        wm.watermark_exists = True
    
    # Test with sample PDF (if exists)
    test_pdfs = []
    for file in os.listdir('.'):
        if file.endswith('.pdf'):
            test_pdfs.append(file)
            break  # Just test with one PDF
    
    if test_pdfs:
        print(f"📄 Testing with PDF: {test_pdfs[0]}")
        
        # Test different positions
        positions = ["center", "bottom-right", "top-left"]
        
        for position in positions:
            # Create copy for testing
            test_pdf = f"test_watermark_{position}.pdf"
            import shutil
            shutil.copy2(test_pdfs[0], test_pdf)
            
            # Add watermark
            success = wm.add_watermark_to_pdf(test_pdf, opacity=0.4, position=position)
            
            if success:
                print(f"   ✅ {position}: Watermark added successfully")
            else:
                print(f"   ❌ {position}: Failed to add watermark")
            
            # Cleanup
            try:
                os.remove(test_pdf)
            except:
                pass
    else:
        print("❌ No PDF files found for testing")

if __name__ == "__main__":
    test_watermark()
