"""
PDF Converter Module
Modul untuk mengkonversi Excel sheet ke PDF dengan mempertahankan formatting
"""

from reportlab.lib.pagesizes import letter, A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import inch, mm
from reportlab.pdfgen import canvas
from reportlab.platypus.tableofcontents import TableOfContents
import os
from excel_reader import ExcelReader

class PDFConverter:
    def __init__(self, preserve_formatting=True, bulk_mode=True):
        """
        Initialize PDF converter
        
        Args:
            preserve_formatting (bool): Apakah mempertahankan formatting Excel
            bulk_mode (bool): Apakah membuat file PDF terpisah untuk setiap sheet
        """
        self.preserve_formatting = preserve_formatting
        self.bulk_mode = bulk_mode
        self.styles = getSampleStyleSheet()
        
    def convert_sheet_to_pdf(self, excel_file, sheet_name, output_file):
        """
        Konversi sheet Excel ke PDF
        
        Args:
            excel_file (str): Path ke file Excel
            sheet_name (str): Nama sheet yang akan dikonversi
            output_file (str): Path output file PDF
        """
        try:
            # Baca data Excel
            reader = ExcelReader(excel_file)
            
            if self.preserve_formatting:
                sheet_data = reader.get_sheet_with_formatting(sheet_name)
                data = sheet_data['data']
                formatting = sheet_data['formatting']
            else:
                data = reader.get_sheet_data(sheet_name)
                formatting = None
            
            # Buat PDF
            self._create_pdf(data, output_file, sheet_name, formatting)
            
            reader.close()
            
        except Exception as e:
            raise Exception(f"Error converting sheet '{sheet_name}': {str(e)}")
    
    def _create_pdf(self, data, output_file, sheet_name, formatting=None):
        """
        Buat file PDF dari data

        Args:
            data (list): Data sheet dalam bentuk list of lists
            output_file (str): Path output file
            sheet_name (str): Nama sheet
            formatting (dict): Informasi formatting (optional)
        """
        # Buat direktori output jika belum ada
        output_dir = os.path.dirname(output_file)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)

        # Filter data kosong di akhir
        filtered_data = self._filter_empty_rows_cols(data)

        if not filtered_data:
            # Jika tidak ada data, buat PDF kosong dengan pesan
            self._create_empty_pdf(output_file, sheet_name)
            return

        # Tentukan orientasi berdasarkan jumlah kolom
        num_cols = len(filtered_data[0]) if filtered_data else 0
        use_landscape = num_cols > 5
        page_size = landscape(A4) if use_landscape else A4

        # Buat dokumen PDF
        doc = SimpleDocTemplate(
            output_file,
            pagesize=page_size,
            rightMargin=20*mm,
            leftMargin=20*mm,
            topMargin=20*mm,
            bottomMargin=20*mm
        )

        # Elemen-elemen yang akan ditambahkan ke PDF
        elements = []

        # Tambahkan judul
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=self.styles['Heading1'],
            fontSize=14,
            spaceAfter=12,
            alignment=1,  # Center alignment
            textColor=colors.darkblue
        )
        title = Paragraph(f"<b>{sheet_name}</b>", title_style)
        elements.append(title)
        elements.append(Spacer(1, 10))

        # Proses data untuk text yang lebih pendek
        processed_data = []
        for row in filtered_data:
            processed_row = []
            for cell in row:
                cell_str = str(cell) if cell is not None else ""
                # Batasi panjang cell untuk mencegah overflow
                if len(cell_str) > 50:
                    cell_str = cell_str[:47] + "..."
                processed_row.append(cell_str)
            processed_data.append(processed_row)

        # Buat tabel
        table = Table(processed_data)

        # Apply styling sederhana
        table_style = self._create_simple_table_style(len(processed_data))
        table.setStyle(table_style)

        # Atur lebar kolom secara merata
        if num_cols > 0:
            available_width = page_size[0] - 40*mm  # Total width minus margins
            col_width = available_width / num_cols
            table._argW = [col_width] * num_cols

        elements.append(table)

        # Build PDF
        doc.build(elements)
    
    def _create_table_style(self, data, formatting=None):
        """
        Buat style untuk tabel

        Args:
            data (list): Data tabel
            formatting (dict): Informasi formatting

        Returns:
            TableStyle: Style untuk tabel
        """
        style_commands = [
            # Header styling
            ('BACKGROUND', (0, 0), (-1, 0), colors.Color(0.2, 0.4, 0.6)),  # Professional blue header
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),  # White text on header
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),  # Center align headers
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  # Bold header font
            ('FONTSIZE', (0, 0), (-1, 0), 9),  # Header font size
            ('TOPPADDING', (0, 0), (-1, 0), 8),  # Header top padding
            ('BOTTOMPADDING', (0, 0), (-1, 0), 8),  # Header bottom padding

            # Data styling
            ('BACKGROUND', (0, 1), (-1, -1), colors.white),  # White background for data
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),  # Regular font for data
            ('FONTSIZE', (0, 1), (-1, -1), 8),  # Data font size
            ('ALIGN', (0, 1), (-1, -1), 'LEFT'),  # Left align data
            ('TOPPADDING', (0, 1), (-1, -1), 4),  # Data top padding
            ('BOTTOMPADDING', (0, 1), (-1, -1), 4),  # Data bottom padding
            ('LEFTPADDING', (0, 0), (-1, -1), 6),  # Left padding
            ('RIGHTPADDING', (0, 0), (-1, -1), 6),  # Right padding

            # Grid and borders
            ('GRID', (0, 0), (-1, -1), 0.5, colors.Color(0.7, 0.7, 0.7)),  # Light gray grid
            ('LINEBELOW', (0, 0), (-1, 0), 2, colors.Color(0.2, 0.4, 0.6)),  # Thick line below header
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),  # Top vertical alignment

            # Alternating row colors for better readability (manual implementation)
        ]

        # Add alternating row colors manually
        if len(data) > 1:
            for row_idx in range(1, len(data)):
                if row_idx % 2 == 0:  # Even rows (0-indexed, so actually odd data rows)
                    style_commands.append(
                        ('BACKGROUND', (0, row_idx), (-1, row_idx), colors.Color(0.95, 0.95, 0.95))
                    )

        # Tambahkan formatting khusus jika ada
        if formatting and self.preserve_formatting:
            style_commands.extend(self._apply_excel_formatting(data, formatting))

        return TableStyle(style_commands)

    def _create_simple_table_style(self, num_rows):
        """
        Buat style tabel yang sederhana dan stabil

        Args:
            num_rows (int): Jumlah baris data

        Returns:
            TableStyle: Style untuk tabel
        """
        style_commands = [
            # Header styling
            ('BACKGROUND', (0, 0), (-1, 0), colors.Color(0.2, 0.4, 0.6)),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('TOPPADDING', (0, 0), (-1, 0), 8),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 8),

            # Data styling
            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 9),
            ('ALIGN', (0, 1), (-1, -1), 'LEFT'),
            ('TOPPADDING', (0, 1), (-1, -1), 6),
            ('BOTTOMPADDING', (0, 1), (-1, -1), 6),
            ('LEFTPADDING', (0, 0), (-1, -1), 8),
            ('RIGHTPADDING', (0, 0), (-1, -1), 8),

            # Grid and borders
            ('GRID', (0, 0), (-1, -1), 0.5, colors.Color(0.7, 0.7, 0.7)),
            ('LINEBELOW', (0, 0), (-1, 0), 2, colors.Color(0.2, 0.4, 0.6)),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ]

        # Add alternating row colors manually
        for row_idx in range(1, num_rows):
            if row_idx % 2 == 0:
                style_commands.append(
                    ('BACKGROUND', (0, row_idx), (-1, row_idx), colors.Color(0.95, 0.95, 0.95))
                )

        return TableStyle(style_commands)

    def _apply_excel_formatting(self, data, formatting):
        """
        Apply formatting Excel ke tabel PDF
        
        Args:
            data (list): Data tabel
            formatting (dict): Informasi formatting Excel
            
        Returns:
            list: List command styling tambahan
        """
        additional_styles = []
        
        from openpyxl.utils import get_column_letter

        for row_idx, row in enumerate(data):
            for col_idx, cell in enumerate(row):
                # Koordinat sel dalam format Excel (A1, B1, etc.)
                cell_coord = f"{get_column_letter(col_idx + 1)}{row_idx + 1}"
                
                if cell_coord in formatting:
                    cell_format = formatting[cell_coord]
                    
                    # Apply bold font
                    if cell_format.get('font_bold'):
                        additional_styles.append(
                            ('FONTNAME', (col_idx, row_idx), (col_idx, row_idx), 'Helvetica-Bold')
                        )
                    
                    # Apply font size
                    font_size = cell_format.get('font_size', 8)
                    if font_size and font_size != 11:  # 11 adalah default Excel
                        pdf_font_size = max(6, min(font_size, 14))  # Batasi ukuran font
                        additional_styles.append(
                            ('FONTSIZE', (col_idx, row_idx), (col_idx, row_idx), pdf_font_size)
                        )
                    
                    # Apply background color
                    fill_color = cell_format.get('fill_color')
                    if fill_color and fill_color != 'FFFFFFFF':  # Bukan putih
                        try:
                            # Konversi hex color ke RGB
                            if len(fill_color) == 8:  # ARGB format
                                fill_color = fill_color[2:]  # Hapus alpha channel
                            
                            r = int(fill_color[0:2], 16) / 255.0
                            g = int(fill_color[2:4], 16) / 255.0
                            b = int(fill_color[4:6], 16) / 255.0
                            
                            additional_styles.append(
                                ('BACKGROUND', (col_idx, row_idx), (col_idx, row_idx), 
                                 colors.Color(r, g, b))
                            )
                        except:
                            pass  # Skip jika error parsing color
                    
                    # Apply text alignment
                    alignment = cell_format.get('alignment', {})
                    horizontal = alignment.get('horizontal')
                    if horizontal:
                        align_map = {
                            'center': 'CENTER',
                            'right': 'RIGHT',
                            'left': 'LEFT'
                        }
                        if horizontal in align_map:
                            additional_styles.append(
                                ('ALIGN', (col_idx, row_idx), (col_idx, row_idx), 
                                 align_map[horizontal])
                            )
        
        return additional_styles

    def _process_data_for_pdf(self, data):
        """
        Proses data untuk PDF dengan text wrapping dan formatting

        Args:
            data (list): Data asli

        Returns:
            list: Data yang sudah diproses
        """
        processed_data = []

        for row in data:
            processed_row = []
            for cell in row:
                cell_str = str(cell) if cell is not None else ""

                # Wrap text panjang
                if len(cell_str) > 30:
                    # Split long text into multiple lines
                    words = cell_str.split(' ')
                    lines = []
                    current_line = ""

                    for word in words:
                        if len(current_line + " " + word) <= 30:
                            current_line += (" " + word) if current_line else word
                        else:
                            if current_line:
                                lines.append(current_line)
                            current_line = word

                    if current_line:
                        lines.append(current_line)

                    cell_str = "\n".join(lines)

                processed_row.append(cell_str)
            processed_data.append(processed_row)

        return processed_data

    def _calculate_column_widths(self, data, available_width):
        """
        Hitung lebar kolom secara dinamis berdasarkan konten

        Args:
            data (list): Data tabel
            available_width (float): Lebar yang tersedia

        Returns:
            list: List lebar kolom
        """
        if not data:
            return []

        num_cols = len(data[0])
        col_max_lengths = [0] * num_cols

        # Hitung panjang maksimum untuk setiap kolom
        for row in data:
            for col_idx, cell in enumerate(row):
                if col_idx < num_cols:
                    cell_length = len(str(cell).split('\n')[0])  # Ambil baris pertama
                    col_max_lengths[col_idx] = max(col_max_lengths[col_idx], cell_length)

        # Hitung total length
        total_length = sum(col_max_lengths)

        if total_length == 0:
            # Jika semua kolom kosong, bagi rata
            return [available_width / num_cols] * num_cols

        # Distribusi lebar berdasarkan proporsi konten
        col_widths = []
        for length in col_max_lengths:
            proportion = length / total_length
            width = max(available_width * proportion, 20*mm)  # Minimum 20mm
            col_widths.append(width)

        # Normalisasi jika total melebihi available width
        total_width = sum(col_widths)
        if total_width > available_width:
            scale_factor = available_width / total_width
            col_widths = [width * scale_factor for width in col_widths]

        return col_widths

    def _filter_empty_rows_cols(self, data):
        """
        Filter baris dan kolom kosong di akhir
        
        Args:
            data (list): Data asli
            
        Returns:
            list: Data yang sudah difilter
        """
        if not data:
            return []
        
        # Hapus baris kosong di akhir
        while data and all(cell == "" or cell == "None" for cell in data[-1]):
            data.pop()
        
        if not data:
            return []
        
        # Hapus kolom kosong di akhir
        max_cols = len(data[0])
        for col_idx in range(max_cols - 1, -1, -1):
            if all(row[col_idx] == "" or row[col_idx] == "None" for row in data if len(row) > col_idx):
                for row in data:
                    if len(row) > col_idx:
                        row.pop(col_idx)
            else:
                break
        
        return data
    
    def _create_empty_pdf(self, output_file, sheet_name):
        """
        Buat PDF kosong dengan pesan
        
        Args:
            output_file (str): Path output file
            sheet_name (str): Nama sheet
        """
        doc = SimpleDocTemplate(output_file, pagesize=A4)
        elements = []
        
        title = Paragraph(f"Sheet: {sheet_name}", self.styles['Heading1'])
        message = Paragraph("Sheet ini kosong atau tidak mengandung data.", self.styles['Normal'])
        
        elements.append(title)
        elements.append(Spacer(1, 20))
        elements.append(message)
        
        doc.build(elements)
