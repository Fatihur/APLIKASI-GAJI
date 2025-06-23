"""
Excel to PDF Converter
Aplikasi untuk mengkonversi file Excel ke PDF dengan fitur bulk conversion
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
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
        
        self.excel_files = []  # Changed to support multiple files
        self.files_data = {}   # Data for all files
        self.selected_sheets = []
        self.output_directory = ""
        self.folder_names = {}  # Custom folder names per file

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
        file_frame = ttk.LabelFrame(main_frame, text="Select Excel Files", padding="10")
        file_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))

        # File list with scrollbar
        files_list_frame = ttk.Frame(file_frame)
        files_list_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))

        self.files_listbox = tk.Listbox(files_list_frame, height=4, selectmode=tk.SINGLE)
        files_scrollbar = ttk.Scrollbar(files_list_frame, orient=tk.VERTICAL, command=self.files_listbox.yview)
        self.files_listbox.configure(yscrollcommand=files_scrollbar.set)

        self.files_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E))
        files_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))

        # File control buttons
        file_buttons_frame = ttk.Frame(file_frame)
        file_buttons_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E))

        ttk.Button(file_buttons_frame, text="Add Files", command=self.browse_files).grid(row=0, column=0, padx=(0, 5))
        ttk.Button(file_buttons_frame, text="Remove Selected", command=self.remove_selected_file).grid(row=0, column=1, padx=(0, 5))
        ttk.Button(file_buttons_frame, text="Clear All", command=self.clear_all_files).grid(row=0, column=2, padx=(0, 5))
        ttk.Button(file_buttons_frame, text="Set Folder Name", command=self.set_folder_name).grid(row=0, column=3)

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
        sheets_frame = ttk.LabelFrame(main_frame, text="Select Sheets to Convert (Auto-ignores sheets 1-9)", padding="10")
        sheets_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 15))
        
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
        files_list_frame.columnconfigure(0, weight=1)
        output_frame.columnconfigure(0, weight=1)

        # Sheets frame
        sheets_frame.columnconfigure(0, weight=1)
        sheets_frame.rowconfigure(0, weight=1)
        listbox_frame.columnconfigure(0, weight=1)
        listbox_frame.rowconfigure(0, weight=1)

        # Convert frame
        convert_frame.columnconfigure(1, weight=1)

        # Bind file selection to update sheets
        self.files_listbox.bind('<<ListboxSelect>>', self.on_file_select)

        # Bind mousewheel to canvas for scrolling
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")

        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
    def browse_files(self):
        """Browse and add multiple Excel files"""
        file_paths = filedialog.askopenfilenames(
            title="Select Excel Files",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )

        for file_path in file_paths:
            if file_path not in self.excel_files:
                self.excel_files.append(file_path)

                # Add to listbox with just filename
                filename = os.path.basename(file_path)
                self.files_listbox.insert(tk.END, filename)

                # Set default folder name (filename without extension)
                base_name = os.path.splitext(filename)[0]
                self.folder_names[file_path] = base_name

                # Load sheets data for this file
                try:
                    reader = ExcelReader(file_path)
                    sheets_info = reader.get_sheets_info()

                    # Filter out sheets 1-9
                    filtered_sheets = {}
                    for sheet_name, info in sheets_info.items():
                        if not self.is_sheet_ignored(sheet_name):
                            filtered_sheets[sheet_name] = info

                    self.files_data[file_path] = filtered_sheets
                    reader.close()

                except Exception as e:
                    print(f"Error loading {filename}: {str(e)}")
                    self.files_data[file_path] = {}

        self.update_convert_button_state()

    def is_sheet_ignored(self, sheet_name):
        """Check if sheet should be ignored based on specific sheet names"""
        # List of sheet names to ignore
        ignored_sheets = [
            'Payroll adjust',
            'Database',
            'Summary Amman',
            'Summary Karyawan',
            'PPh 21',
            'Payroll',
            'Tarif TER',
            'hr_libur',
            'jm_istrht'
        ]

        # Case-insensitive comparison
        sheet_name_lower = sheet_name.strip().lower()

        for ignored_sheet in ignored_sheets:
            if sheet_name_lower == ignored_sheet.lower():
                return True

        return False

    def remove_selected_file(self):
        """Remove selected file from list"""
        selection = self.files_listbox.curselection()
        if selection:
            index = selection[0]
            if index < len(self.excel_files):
                file_path = self.excel_files[index]

                # Remove from all data structures
                self.excel_files.pop(index)
                self.files_listbox.delete(index)

                if file_path in self.files_data:
                    del self.files_data[file_path]
                if file_path in self.folder_names:
                    del self.folder_names[file_path]

                # Clear sheets if this was the selected file
                self.sheets_listbox.delete(0, tk.END)

        self.update_convert_button_state()

    def clear_all_files(self):
        """Clear all files"""
        self.excel_files.clear()
        self.files_data.clear()
        self.folder_names.clear()
        self.files_listbox.delete(0, tk.END)
        self.sheets_listbox.delete(0, tk.END)
        self.update_convert_button_state()

    def set_folder_name(self):
        """Set custom folder name for selected file"""
        selection = self.files_listbox.curselection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a file first")
            return

        index = selection[0]
        if index >= len(self.excel_files):
            return

        file_path = self.excel_files[index]
        current_name = self.folder_names.get(file_path, "")

        # Simple input dialog
        new_name = tk.simpledialog.askstring(
            "Folder Name",
            f"Enter folder name for {os.path.basename(file_path)}:",
            initialvalue=current_name
        )

        if new_name:
            # Sanitize folder name
            safe_name = "".join(c for c in new_name if c.isalnum() or c in (' ', '-', '_')).rstrip()
            self.folder_names[file_path] = safe_name

    def on_file_select(self, event):
        """Handle file selection to show its sheets"""
        selection = self.files_listbox.curselection()
        if not selection:
            return

        index = selection[0]
        if index >= len(self.excel_files):
            return

        file_path = self.excel_files[index]

        # Clear current sheets
        self.sheets_listbox.delete(0, tk.END)

        # Load sheets for selected file
        if file_path in self.files_data:
            sheets_data = self.files_data[file_path]
            for sheet_name in sheets_data.keys():
                self.sheets_listbox.insert(tk.END, sheet_name)

        self.status_var.set(f"Showing sheets for: {os.path.basename(file_path)}")

    def update_convert_button_state(self):
        """Update convert button state based on files"""
        if self.excel_files:
            self.convert_button.config(state='normal')
            self.status_var.set(f"Ready to convert {len(self.excel_files)} file(s)")
        else:
            self.convert_button.config(state='disabled')
            self.status_var.set("Add Excel files to begin")

    def browse_output_directory(self):
        directory = filedialog.askdirectory(
            title="Select Output Directory"
        )

        if directory:
            self.output_path_var.set(directory)
            self.output_directory = directory

    def load_sheets_for_file(self, file_path):
        """Load sheets for a specific file"""
        try:
            reader = ExcelReader(file_path)
            sheets_info = reader.get_sheets_info()

            # Filter out sheets 1-9
            filtered_sheets = {}
            for sheet_name, info in sheets_info.items():
                if not self.is_sheet_ignored(sheet_name):
                    filtered_sheets[sheet_name] = info

            self.files_data[file_path] = filtered_sheets
            reader.close()

            return filtered_sheets

        except Exception as e:
            print(f"Error loading {os.path.basename(file_path)}: {str(e)}")
            return {}
            
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
        if not self.excel_files:
            messagebox.showwarning("Warning", "Please add Excel files first")
            return

        # Check if we have a file selected for sheet selection
        file_selection = self.files_listbox.curselection()
        if file_selection:
            # Convert only selected sheets from selected file
            selected_indices = self.sheets_listbox.curselection()
            if not selected_indices:
                messagebox.showwarning("Warning", "Please select at least one sheet to convert")
                return

            self.selected_sheets = [self.sheets_listbox.get(i) for i in selected_indices]
            selected_file_index = file_selection[0]
            self.selected_file = self.excel_files[selected_file_index]
        else:
            # Convert all sheets from all files
            self.selected_sheets = []
            self.selected_file = None

        # Start conversion in separate thread
        self.convert_button.config(state='disabled')
        self.progress_var.set(0)

        thread = threading.Thread(target=self.convert_sheets)
        thread.daemon = True
        thread.start()
        
    def convert_sheets(self):
        try:
            conversion_method = self.conversion_method_var.get()

            # Determine base output directory
            base_output_dir = self.output_directory if self.output_directory else os.getcwd()

            total_files = 0
            total_sheets_converted = 0

            if hasattr(self, 'selected_file') and self.selected_file:
                # Convert only selected sheets from selected file
                files_to_process = [(self.selected_file, self.selected_sheets)]
                total_files = 1
            else:
                # Convert all sheets from all files
                files_to_process = []
                for file_path in self.excel_files:
                    if file_path in self.files_data:
                        sheets = list(self.files_data[file_path].keys())
                        if sheets:  # Only add if file has sheets
                            files_to_process.append((file_path, sheets))
                total_files = len(files_to_process)

            if not files_to_process:
                messagebox.showwarning("Warning", "No sheets to convert")
                return

            # Calculate total sheets for progress
            total_sheets = sum(len(sheets) for _, sheets in files_to_process)
            current_sheet = 0

            for file_path, sheets_to_convert in files_to_process:
                # Create folder for this file
                folder_name = self.folder_names.get(file_path, os.path.splitext(os.path.basename(file_path))[0])
                file_output_dir = os.path.join(base_output_dir, folder_name)

                if not os.path.exists(file_output_dir):
                    os.makedirs(file_output_dir)

                self.status_var.set(f"Processing file: {os.path.basename(file_path)}")

                if conversion_method == "capture":
                    converter = PDFConverterCapture()

                    for sheet_name in sheets_to_convert:
                        self.status_var.set(f"Capturing: {sheet_name} from {os.path.basename(file_path)}")

                        # Create output path in file's folder
                        safe_sheet_name = "".join(c for c in sheet_name if c.isalnum() or c in (' ', '-', '_')).rstrip()
                        output_path = os.path.join(file_output_dir, f"{safe_sheet_name}.pdf")

                        success = converter.convert_single_sheet(file_path, sheet_name, output_path)

                        if success:
                            total_sheets_converted += 1
                        else:
                            print(f"Failed to capture sheet: {sheet_name}")

                        current_sheet += 1
                        progress = (current_sheet / total_sheets) * 100
                        self.progress_var.set(progress)

                else:
                    # Table conversion method
                    converter = PDFConverter(
                        preserve_formatting=self.preserve_format_var.get(),
                        bulk_mode=self.bulk_mode_var.get()
                    )

                    for sheet_name in sheets_to_convert:
                        self.status_var.set(f"Converting: {sheet_name} from {os.path.basename(file_path)}")

                        # Create output path in file's folder
                        safe_sheet_name = "".join(c for c in sheet_name if c.isalnum() or c in (' ', '-', '_')).rstrip()
                        output_path = os.path.join(file_output_dir, f"{safe_sheet_name}.pdf")

                        try:
                            converter.convert_sheet_to_pdf(file_path, sheet_name, output_path)
                            total_sheets_converted += 1
                        except Exception as e:
                            print(f"Failed to convert sheet {sheet_name}: {str(e)}")

                        current_sheet += 1
                        progress = (current_sheet / total_sheets) * 100
                        self.progress_var.set(progress)

            self.status_var.set(f"Conversion completed: {total_sheets_converted}/{total_sheets} sheets from {total_files} file(s)")
            messagebox.showinfo("Success",
                              f"Successfully converted {total_sheets_converted}/{total_sheets} sheets from {total_files} file(s)\n"
                              f"Output saved to: {base_output_dir}")

        except Exception as e:
            messagebox.showerror("Error", f"Conversion failed: {str(e)}")
            self.status_var.set("Conversion failed")

        finally:
            self.convert_button.config(state='normal')
            self.progress_var.set(0)
            
    def get_files_summary(self):
        """Get summary of loaded files"""
        if not self.excel_files:
            return "No files loaded"

        total_sheets = 0
        for file_path in self.excel_files:
            if file_path in self.files_data:
                total_sheets += len(self.files_data[file_path])

        return f"{len(self.excel_files)} file(s) loaded, {total_sheets} sheets total (sheets 1-9 ignored)"

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelToPDFApp(root)
    root.mainloop()
