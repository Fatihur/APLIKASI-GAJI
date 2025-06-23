"""
Excel to PDF Converter
Aplikasi untuk mengkonversi file Excel ke PDF dengan fitur bulk conversion
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from excel_reader import ExcelReader
from pdf_converter import PDFConverter
from pdf_converter_capture import PDFConverterCapture
import threading

class ExcelToPDFApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel to PDF Converter")
        self.root.geometry("900x700")
        self.root.configure(bg='#f0f0f0')
        self.root.minsize(800, 600)
        
        self.excel_file = None
        self.sheets_data = {}
        self.selected_sheets = []
        self.output_directory = ""

        self.setup_ui()
        
    def setup_ui(self):
        # Configure root grid
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        # Create main canvas and scrollbar for scrolling
        canvas = tk.Canvas(self.root, bg='#f0f0f0')
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Pack canvas and scrollbar
        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        # Main frame inside scrollable frame
        main_frame = ttk.Frame(scrollable_frame, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Title
        title_label = ttk.Label(main_frame, text="Excel to PDF Converter",
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 15))
        
        # File selection
        file_frame = ttk.LabelFrame(main_frame, text="Select Excel File", padding="10")
        file_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 20))

        self.file_path_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.file_path_var, width=50, state='readonly').grid(row=0, column=0, padx=(0, 10))
        ttk.Button(file_frame, text="Browse", command=self.browse_file).grid(row=0, column=1)

        # Output directory selection
        output_frame = ttk.LabelFrame(main_frame, text="Output Directory", padding="10")
        output_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 20))

        self.output_path_var = tk.StringVar()
        ttk.Entry(output_frame, textvariable=self.output_path_var, width=50, state='readonly').grid(row=0, column=0, padx=(0, 10))
        ttk.Button(output_frame, text="Browse", command=self.browse_output_directory).grid(row=0, column=1)

        # Default output directory info
        default_info = ttk.Label(output_frame, text="(Leave empty to save in same directory as Excel file)",
                                font=('Arial', 8), foreground='gray')
        default_info.grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=(5, 0))
        
        # Sheets selection
        sheets_frame = ttk.LabelFrame(main_frame, text="Select Sheets to Convert", padding="10")
        sheets_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 20))
        
        # Sheets listbox with scrollbar
        listbox_frame = ttk.Frame(sheets_frame)
        listbox_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S))

        self.sheets_listbox = tk.Listbox(listbox_frame, selectmode=tk.MULTIPLE, height=6)
        sheets_scrollbar = ttk.Scrollbar(listbox_frame, orient=tk.VERTICAL, command=self.sheets_listbox.yview)
        self.sheets_listbox.configure(yscrollcommand=sheets_scrollbar.set)

        self.sheets_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        sheets_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Selection buttons
        button_frame = ttk.Frame(sheets_frame)
        button_frame.grid(row=1, column=0, columnspan=3, pady=(10, 0))
        
        ttk.Button(button_frame, text="Select All", command=self.select_all_sheets).grid(row=0, column=0, padx=(0, 10))
        ttk.Button(button_frame, text="Clear All", command=self.clear_all_sheets).grid(row=0, column=1)
        
        # Conversion options
        options_frame = ttk.LabelFrame(main_frame, text="Conversion Options", padding="10")
        options_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 20))
        
        self.bulk_mode_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Bulk Mode (One PDF per sheet)", 
                       variable=self.bulk_mode_var).grid(row=0, column=0, sticky=tk.W)
        
        self.preserve_format_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Preserve Original Formatting",
                       variable=self.preserve_format_var).grid(row=1, column=0, sticky=tk.W)

        # Conversion method selection
        method_frame = ttk.Frame(options_frame)
        method_frame.grid(row=2, column=0, sticky=tk.W, pady=(10, 0))

        ttk.Label(method_frame, text="Conversion Method:").grid(row=0, column=0, sticky=tk.W)

        self.conversion_method_var = tk.StringVar(value="capture")
        ttk.Radiobutton(method_frame, text="Capture Method (Recommended)",
                       variable=self.conversion_method_var, value="capture").grid(row=1, column=0, sticky=tk.W, padx=(20, 0))
        ttk.Radiobutton(method_frame, text="Table Conversion",
                       variable=self.conversion_method_var, value="table").grid(row=2, column=0, sticky=tk.W, padx=(20, 0))

        # Method description
        desc_frame = ttk.Frame(options_frame)
        desc_frame.grid(row=3, column=0, sticky=tk.W, pady=(5, 0))

        self.method_desc_var = tk.StringVar(value="Capture method preserves exact Excel layout and formatting")
        ttk.Label(desc_frame, textvariable=self.method_desc_var,
                 font=('Arial', 8), foreground='gray').grid(row=0, column=0, sticky=tk.W, padx=(20, 0))

        # Bind method change to update description
        self.conversion_method_var.trace('w', self.update_method_description)
        
        # Convert button and progress
        convert_frame = ttk.LabelFrame(main_frame, text="Convert to PDF", padding="15")
        convert_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 15))

        # Convert button
        self.convert_button = ttk.Button(convert_frame, text="🔄 Convert to PDF",
                                        command=self.start_conversion, state='disabled')
        self.convert_button.grid(row=0, column=0, columnspan=2, pady=(0, 10), sticky=(tk.W, tk.E))

        # Progress bar
        progress_label = ttk.Label(convert_frame, text="Progress:")
        progress_label.grid(row=1, column=0, sticky=tk.W, padx=(0, 10))

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(convert_frame, variable=self.progress_var,
                                          maximum=100, length=400)
        self.progress_bar.grid(row=1, column=1, sticky=(tk.W, tk.E))

        # Status label
        self.status_var = tk.StringVar(value="Select an Excel file to begin")
        status_label = ttk.Label(convert_frame, textvariable=self.status_var,
                                font=('Arial', 9), foreground='blue')
        status_label.grid(row=2, column=0, columnspan=2, pady=(10, 0), sticky=tk.W)
        
        # Configure grid weights
        scrollable_frame.columnconfigure(0, weight=1)
        scrollable_frame.rowconfigure(0, weight=1)

        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(3, weight=1)  # Sheets frame gets extra space

        # File and output frames
        file_frame.columnconfigure(0, weight=1)
        output_frame.columnconfigure(0, weight=1)

        # Sheets frame
        sheets_frame.columnconfigure(0, weight=1)
        sheets_frame.rowconfigure(0, weight=1)
        listbox_frame.columnconfigure(0, weight=1)
        listbox_frame.rowconfigure(0, weight=1)

        # Convert frame
        convert_frame.columnconfigure(1, weight=1)

        # Bind mousewheel to canvas for scrolling
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")

        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        
        if file_path:
            self.file_path_var.set(file_path)
            self.excel_file = file_path
            self.load_sheets()

    def browse_output_directory(self):
        directory = filedialog.askdirectory(
            title="Select Output Directory"
        )

        if directory:
            self.output_path_var.set(directory)
            self.output_directory = directory

    def load_sheets(self):
        try:
            self.status_var.set("Loading sheets...")
            reader = ExcelReader(self.excel_file)
            self.sheets_data = reader.get_sheets_info()
            
            # Clear and populate listbox
            self.sheets_listbox.delete(0, tk.END)
            for sheet_name in self.sheets_data.keys():
                self.sheets_listbox.insert(tk.END, sheet_name)
                
            self.convert_button.config(state='normal')
            self.status_var.set(f"Loaded {len(self.sheets_data)} sheets")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load Excel file: {str(e)}")
            self.status_var.set("Error loading file")
            
    def select_all_sheets(self):
        self.sheets_listbox.select_set(0, tk.END)
        
    def clear_all_sheets(self):
        self.sheets_listbox.selection_clear(0, tk.END)

    def update_method_description(self, *args):
        """Update description based on selected conversion method"""
        method = self.conversion_method_var.get()
        if method == "capture":
            self.method_desc_var.set("Capture method preserves exact Excel layout and formatting (requires Excel installed)")
        else:
            self.method_desc_var.set("Table conversion method converts data to PDF tables (no Excel required)")
        
    def start_conversion(self):
        selected_indices = self.sheets_listbox.curselection()
        if not selected_indices:
            messagebox.showwarning("Warning", "Please select at least one sheet to convert")
            return
            
        self.selected_sheets = [self.sheets_listbox.get(i) for i in selected_indices]
        
        # Start conversion in separate thread
        self.convert_button.config(state='disabled')
        self.progress_var.set(0)
        
        thread = threading.Thread(target=self.convert_sheets)
        thread.daemon = True
        thread.start()
        
    def convert_sheets(self):
        try:
            conversion_method = self.conversion_method_var.get()
            total_sheets = len(self.selected_sheets)

            if conversion_method == "capture":
                # Gunakan capture method
                self.status_var.set("Initializing Excel capture...")

                # Tentukan output directory
                if self.output_directory:
                    output_dir = self.output_directory
                else:
                    output_dir = os.path.dirname(self.excel_file)

                converter = PDFConverterCapture()

                for i, sheet_name in enumerate(self.selected_sheets):
                    self.status_var.set(f"Capturing sheet: {sheet_name}")

                    # Convert sheet menggunakan capture method
                    output_path = self.get_output_filename(sheet_name)
                    success = converter.convert_single_sheet(
                        self.excel_file,
                        sheet_name,
                        output_path
                    )

                    if not success:
                        raise Exception(f"Failed to capture sheet: {sheet_name}")

                    # Update progress
                    progress = ((i + 1) / total_sheets) * 100
                    self.progress_var.set(progress)

            else:
                # Gunakan table conversion method
                converter = PDFConverter(
                    preserve_formatting=self.preserve_format_var.get(),
                    bulk_mode=self.bulk_mode_var.get()
                )

                for i, sheet_name in enumerate(self.selected_sheets):
                    self.status_var.set(f"Converting sheet: {sheet_name}")

                    # Convert sheet to PDF
                    converter.convert_sheet_to_pdf(
                        self.excel_file,
                        sheet_name,
                        self.get_output_filename(sheet_name)
                    )

                    # Update progress
                    progress = ((i + 1) / total_sheets) * 100
                    self.progress_var.set(progress)

            self.status_var.set(f"Successfully converted {total_sheets} sheets using {conversion_method} method")
            messagebox.showinfo("Success", f"Successfully converted {total_sheets} sheets to PDF")

        except Exception as e:
            messagebox.showerror("Error", f"Conversion failed: {str(e)}")
            self.status_var.set("Conversion failed")

        finally:
            self.convert_button.config(state='normal')
            self.progress_var.set(0)
            
    def get_output_filename(self, sheet_name):
        base_name = os.path.splitext(os.path.basename(self.excel_file))[0]
        safe_sheet_name = "".join(c for c in sheet_name if c.isalnum() or c in (' ', '-', '_')).rstrip()
        filename = f"{base_name}_{safe_sheet_name}.pdf"

        # Tentukan direktori output
        if self.output_directory:
            output_path = os.path.join(self.output_directory, filename)
        else:
            # Gunakan direktori yang sama dengan file Excel
            excel_dir = os.path.dirname(self.excel_file)
            output_path = os.path.join(excel_dir, filename)

        return output_path

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelToPDFApp(root)
    root.mainloop()
