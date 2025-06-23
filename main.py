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
import ttkbootstrap as tb
from ttkbootstrap.constants import *

class ExcelToPDFApp:
    def __init__(self, root):
        self.root = root
        self.root.title("SLIP GAJI PDF GENERATE")
        self.root.geometry("1000x800")
        self.root.configure(bg='#1a1a1a')  # Dark background
        self.root.minsize(900, 700)

        # Set default options (hidden from UI but functional)
        self.bulk_mode_var = tk.BooleanVar(value=True)
        self.preserve_format_var = tk.BooleanVar(value=True)
        self.enable_watermark_var = tk.BooleanVar(value=True)
        self.conversion_method_var = tk.StringVar(value="capture")
        
        self.excel_files = []  # Changed to support multiple files
        self.files_data = {}   # Data for all files
        self.selected_sheets = []
        self.output_directory = ""
        self.folder_names = {}  # Custom folder names per file
        self.file_checkboxes = {}  # Checkbox variables per file
        self.file_checkbox_widgets = {}  # Checkbox widgets per file

        self.setup_ui()

    def setup_ui(self):
        # Configure root grid for dark mode
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        self.root.configure(bg='#1a1a1a')  # Dark background

        # Create main canvas and scrollbar for dark mode
        canvas = tb.Canvas(self.root, background='#1a1a1a', highlightthickness=0)
        scrollbar = tb.Scrollbar(self.root, orient="vertical", command=canvas.yview, bootstyle="light-round")
        scrollable_frame = tb.Frame(canvas, bootstyle="dark")

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Pack canvas and scrollbar
        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        # Main frame inside scrollable frame with dark styling
        main_frame = tb.Frame(scrollable_frame, padding=(35, 30, 35, 30), bootstyle="dark")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Title with dark mode styling
        title_label = tb.Label(main_frame, text=" SLIP GAJI PDF GENERATE",
                               font=('Segoe UI', 24, 'bold'), bootstyle="light",
                               background='#1a1a1a', foreground='#ffffff')
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 30))

        # File selection with dark mode styling
        file_frame = tb.LabelFrame(main_frame, text="📁 Select Excel Files", padding="20", bootstyle="info")
        file_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 25))

        # File list with checkboxes
        files_list_frame = tb.Frame(file_frame)
        files_list_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 8))

        # Simple frame for file checkboxes
        self.files_container = tb.Frame(files_list_frame)
        self.files_container.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=5, pady=5)

        # Add "Select All Files" checkbox
        self.select_all_files_var = tk.BooleanVar()
        select_all_cb = tb.Checkbutton(files_list_frame, text="Select All Files",
                                       variable=self.select_all_files_var,
                                       command=self.toggle_all_files)
        select_all_cb.grid(row=1, column=0, sticky=tk.W, pady=(5, 0))

        # File control buttons
        file_buttons_frame = tb.Frame(file_frame)
        file_buttons_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(8, 0))

        tb.Button(file_buttons_frame, text="Add Files", command=self.browse_files, width=14).grid(row=0, column=0, padx=(0, 8))
        tb.Button(file_buttons_frame, text="Remove Selected", command=self.remove_selected_file, width=16).grid(row=0, column=1, padx=(0, 8))
        tb.Button(file_buttons_frame, text="Clear All", command=self.clear_all_files, width=10).grid(row=0, column=2, padx=(0, 8))
        tb.Button(file_buttons_frame, text="Set Folder Name", command=self.set_folder_name, width=16).grid(row=0, column=3)

        # Output directory selection with dark mode styling
        output_frame = tb.LabelFrame(main_frame, text="📂 Output Directory", padding="20", bootstyle="secondary")
        output_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 25))

        self.output_path_var = tk.StringVar()
        tb.Entry(output_frame, textvariable=self.output_path_var, width=50, state='readonly').grid(row=0, column=0, padx=(0, 10))
        tb.Button(output_frame, text="Browse", command=self.browse_output_directory, width=10).grid(row=0, column=1)

        # Default output directory info
        default_info = tb.Label(output_frame, text="(Leave empty to save in same directory as Excel file)",
                                font=('Arial', 8), foreground='gray')
        default_info.grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=(5, 0))

        # Sheets selection with dark mode styling
        sheets_frame = tb.LabelFrame(main_frame, text="📋 Select Sheets to Convert", padding="20", bootstyle="success")
        sheets_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 25))

        # Add info label for auto-ignore feature
        info_label = tb.Label(sheets_frame,
                              text="ℹ️ Auto-ignores: Payroll adjust, Database, Summary sheets, etc.",
                              style='Info.TLabel')
        info_label.grid(row=0, column=0, columnspan=3, sticky=tk.W, pady=(0, 8))

        # Sheets listbox with scrollbar
        listbox_frame = tb.Frame(sheets_frame)
        listbox_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S))

        self.sheets_listbox = tk.Listbox(listbox_frame, selectmode=tk.MULTIPLE, height=6, font=('Segoe UI', 10))
        sheets_scrollbar = tb.Scrollbar(listbox_frame, orient=tk.VERTICAL, command=self.sheets_listbox.yview)
        self.sheets_listbox.configure(yscrollcommand=sheets_scrollbar.set)

        self.sheets_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        sheets_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))

        # Selection buttons
        button_frame = tb.Frame(sheets_frame)
        button_frame.grid(row=2, column=0, columnspan=3, pady=(12, 0))

        tb.Button(button_frame, text="Select All", command=self.select_all_sheets, width=12).grid(row=0, column=0, padx=(0, 8))
        tb.Button(button_frame, text="Clear All", command=self.clear_all_sheets, width=12).grid(row=0, column=1)

        # Conversion options are hidden but functional (set in __init__)

        # Convert button and progress with dark mode styling
        convert_frame = tb.LabelFrame(main_frame, text="🚀 Convert to PDF", padding="25", bootstyle="warning")
        convert_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(15, 25))

        # Convert button with dark mode styling
        self.convert_button = tb.Button(convert_frame, text="🚀 Convert to PDF",
                                        command=self.start_conversion, state='disabled',
                                        bootstyle="success-outline", width=30)
        self.convert_button.grid(row=0, column=0, columnspan=3, pady=(0, 20), sticky=(tk.W, tk.E))

        # Progress bar with dark mode styling
        progress_label = tb.Label(convert_frame, text="📊 Progress:",
                                  font=('Segoe UI', 12, 'bold'), bootstyle="light")
        progress_label.grid(row=1, column=0, sticky=tk.W, padx=(0, 15), pady=(8, 0))

        self.progress_var = tk.DoubleVar()
        self.progress_bar = tb.Progressbar(convert_frame, variable=self.progress_var,
                                          maximum=100, length=500, mode='determinate', bootstyle="success-striped")
        self.progress_bar.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=(8, 0))

        # Progress percentage label
        self.progress_percent_var = tk.StringVar(value="0%")
        progress_percent_label = tb.Label(convert_frame, textvariable=self.progress_percent_var,
                                         font=('Segoe UI', 12, 'bold'), bootstyle="success")
        progress_percent_label.grid(row=1, column=2, padx=(15, 0), pady=(8, 0))

        # Current sheet being processed with dark mode styling
        self.current_sheet_var = tk.StringVar(value="")
        current_sheet_label = tb.Label(convert_frame, textvariable=self.current_sheet_var,
                                       font=('Segoe UI', 10, 'bold'), bootstyle="success")
        current_sheet_label.grid(row=2, column=0, columnspan=3, pady=(12, 8), sticky=tk.W)

        # Status label with dark mode styling
        self.status_var = tk.StringVar(value="🌙 Add Excel files to begin")
        status_label = tb.Label(convert_frame, textvariable=self.status_var,
                                font=('Segoe UI', 10), bootstyle="info")
        status_label.grid(row=3, column=0, columnspan=3, pady=(8, 0), sticky=tk.W)

        # Responsive: all frames and widgets expand horizontally & vertically
        responsive_frames = [main_frame, file_frame, files_list_frame, output_frame, sheets_frame, listbox_frame, button_frame, convert_frame]
        for frame in responsive_frames:
            frame.grid(sticky="nsew")
            for col in range(4):
                try:
                    frame.grid_columnconfigure(col, weight=1)
                except:
                    pass
            for row in range(4):
                try:
                    frame.grid_rowconfigure(row, weight=1)
                except:
                    pass

        # Pastikan root dan scrollable_frame juga responsif
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        scrollable_frame.grid_rowconfigure(0, weight=1)
        scrollable_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_rowconfigure(3, weight=1)
        main_frame.grid_columnconfigure(0, weight=1)

        # Listbox dan scrollbar expand
        self.sheets_listbox.grid(sticky="nsew")
        sheets_scrollbar.grid(sticky="ns")
        listbox_frame.grid_rowconfigure(0, weight=1)
        listbox_frame.grid_columnconfigure(0, weight=1)

        # Progress bar expands
        self.progress_bar.grid(sticky="ew")
        convert_frame.grid_columnconfigure(1, weight=1)

        # Bind mousewheel to canvas for scrolling
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")

        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
    def browse_files(self):
        """Browse and add multiple Excel files"""
        print("🔍 Browse files called")  # Debug

        file_paths = filedialog.askopenfilenames(
            title="Select Excel Files",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )

        print(f"📁 Selected files: {file_paths}")  # Debug

        if not file_paths:
            print("❌ No files selected")
            return

        for file_path in file_paths:
            print(f"📄 Processing file: {file_path}")  # Debug

            if file_path not in self.excel_files:
                self.excel_files.append(file_path)

                # Set default folder name (filename without extension)
                filename = os.path.basename(file_path)
                base_name = os.path.splitext(filename)[0]
                self.folder_names[file_path] = base_name

                print(f"✅ Added file: {filename}")  # Debug

                # Create checkbox for this file
                try:
                    self.add_file_checkbox(file_path, filename)
                    print(f"✅ Created checkbox for: {filename}")  # Debug
                except Exception as e:
                    print(f"❌ Error creating checkbox: {str(e)}")

                # Load sheets data for this file
                try:
                    reader = ExcelReader(file_path)
                    sheets_info = reader.get_sheets_info()

                    # Filter out ignored sheets
                    filtered_sheets = {}
                    for sheet_name, info in sheets_info.items():
                        if not self.is_sheet_ignored(sheet_name):
                            filtered_sheets[sheet_name] = info

                    self.files_data[file_path] = filtered_sheets
                    reader.close()

                    print(f"✅ Loaded {len(filtered_sheets)} sheets from {filename}")  # Debug

                except Exception as e:
                    print(f"❌ Error loading {filename}: {str(e)}")
                    self.files_data[file_path] = {}
            else:
                print(f"⚠️  File already added: {file_path}")

        self.update_convert_button_state()
        print(f"📊 Total files: {len(self.excel_files)}")  # Debug

    def add_file_checkbox(self, file_path, filename):
        """Add checkbox for a file"""
        # Create checkbox variable
        var = tk.BooleanVar(value=True)  # Default checked
        self.file_checkboxes[file_path] = var

        # Create checkbox directly in container
        checkbox = tb.Checkbutton(self.files_container, text=filename, variable=var,
                                  command=self.update_convert_button_state)
        checkbox.pack(anchor=tk.W, padx=5, pady=2)

        # Store widget reference
        self.file_checkbox_widgets[file_path] = {
            'checkbox': checkbox,
            'var': var
        }

        # Bind click to show sheets
        checkbox.bind('<Button-1>', lambda e, fp=file_path: self.show_sheets_for_file(fp))

    def show_sheets_for_file(self, file_path):
        """Show sheets for clicked file"""
        # Clear current sheets
        self.sheets_listbox.delete(0, tk.END)

        # Load sheets for selected file
        if file_path in self.files_data:
            sheets_data = self.files_data[file_path]
            for sheet_name in sheets_data.keys():
                self.sheets_listbox.insert(tk.END, sheet_name)

        filename = os.path.basename(file_path)
        self.status_var.set(f"Showing sheets for: {filename}")

    def toggle_all_files(self):
        """Toggle all file checkboxes"""
        select_all = self.select_all_files_var.get()

        for file_path in self.excel_files:
            if file_path in self.file_checkboxes:
                self.file_checkboxes[file_path].set(select_all)

        self.update_convert_button_state()

    def is_sheet_ignored(self, sheet_name):
        """Check if sheet should be ignored based on keywords"""
        # List of exact matches and keywords to ignore (case-insensitive)
        ignored_exact = [
            'payroll adjust',
            'database',
            'summary amman',
            'summary karyawan',
            'pph 21',
            'payroll',
            'payrol',
            'tarif ter',
            'hr_libur',
            'jm_istrht'
        ]

        # Keywords that should match as whole words or at word boundaries
        ignored_keywords = [
            'adjust',  # But not "adjustment"
        ]

        # Case-insensitive comparison
        sheet_name_lower = sheet_name.strip().lower()

        # Check exact matches first
        for exact_match in ignored_exact:
            if exact_match in sheet_name_lower:
                return True

        # Check keyword matches (whole word boundaries)
        import re
        for keyword in ignored_keywords:
            # Use word boundary regex to match whole words only
            pattern = r'\b' + re.escape(keyword) + r'\b'
            if re.search(pattern, sheet_name_lower):
                return True

        return False

    def remove_selected_file(self):
        """Remove selected file from list"""
        # Find checked files to remove
        files_to_remove = []
        for file_path in self.excel_files:
            if file_path in self.file_checkboxes and self.file_checkboxes[file_path].get():
                files_to_remove.append(file_path)

        if not files_to_remove:
            messagebox.showwarning("Warning", "Please check files to remove")
            return

        for file_path in files_to_remove:
            # Remove from data structures
            self.excel_files.remove(file_path)

            if file_path in self.files_data:
                del self.files_data[file_path]
            if file_path in self.folder_names:
                del self.folder_names[file_path]

            # Remove checkbox widget
            if file_path in self.file_checkbox_widgets:
                self.file_checkbox_widgets[file_path]['checkbox'].destroy()
                del self.file_checkbox_widgets[file_path]
                del self.file_checkboxes[file_path]

        # Clear sheets
        self.sheets_listbox.delete(0, tk.END)
        self.update_convert_button_state()

    def clear_all_files(self):
        """Clear all files"""
        # Remove all checkbox widgets
        for file_path in list(self.file_checkbox_widgets.keys()):
            self.file_checkbox_widgets[file_path]['checkbox'].destroy()

        # Clear all data structures
        self.excel_files.clear()
        self.files_data.clear()
        self.folder_names.clear()
        self.file_checkboxes.clear()
        self.file_checkbox_widgets.clear()
        self.sheets_listbox.delete(0, tk.END)
        self.select_all_files_var.set(False)
        self.update_convert_button_state()

    def set_folder_name(self):
        """Set custom folder name for checked files"""
        # Find checked files
        checked_files = []
        for file_path in self.excel_files:
            if file_path in self.file_checkboxes and self.file_checkboxes[file_path].get():
                checked_files.append(file_path)

        if not checked_files:
            messagebox.showwarning("Warning", "Please check files to set folder names")
            return

        if len(checked_files) == 1:
            # Single file - direct input
            file_path = checked_files[0]
            current_name = self.folder_names.get(file_path, "")

            new_name = simpledialog.askstring(
                "Folder Name",
                f"Enter folder name for {os.path.basename(file_path)}:",
                initialvalue=current_name
            )

            if new_name:
                safe_name = "".join(c for c in new_name if c.isalnum() or c in (' ', '-', '_')).rstrip()
                self.folder_names[file_path] = safe_name
        else:
            # Multiple files - show dialog for each
            for file_path in checked_files:
                current_name = self.folder_names.get(file_path, "")

                new_name = simpledialog.askstring(
                    "Folder Name",
                    f"Enter folder name for {os.path.basename(file_path)}:",
                    initialvalue=current_name
                )

                if new_name:
                    safe_name = "".join(c for c in new_name if c.isalnum() or c in (' ', '-', '_')).rstrip()
                    self.folder_names[file_path] = safe_name



    def update_convert_button_state(self):
        """Update convert button state based on checked files"""
        checked_files = []
        for file_path in self.excel_files:
            if file_path in self.file_checkboxes and self.file_checkboxes[file_path].get():
                checked_files.append(file_path)

        if checked_files:
            self.convert_button.config(state='normal')
            total_sheets = sum(len(self.files_data.get(f, {})) for f in checked_files)
            self.status_var.set(f"Ready to convert {len(checked_files)} file(s), {total_sheets} sheets")
        else:
            self.convert_button.config(state='disabled')
            if self.excel_files:
                self.status_var.set("Check files to convert")
            else:
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


        
    def start_conversion(self):
        if not self.excel_files:
            messagebox.showwarning("Warning", "Please add Excel files first")
            return

        # Check if any files are checked
        checked_files = [f for f in self.excel_files if self.file_checkboxes.get(f, tk.BooleanVar()).get()]
        if not checked_files:
            messagebox.showwarning("Warning", "Please check at least one file to convert")
            return

        # Auto-select all valid sheets (not ignored) from checked files
        self.selected_sheets = []
        self.selected_file = None

        # Start conversion in separate thread
        self.convert_button.config(state='disabled')
        self.progress_var.set(0)
        self.progress_percent_var.set("0%")
        self.current_sheet_var.set("")

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

            # Convert all valid sheets from checked files
            files_to_process = []
            for file_path in self.excel_files:
                # Only process checked files
                if (file_path in self.file_checkboxes and
                    self.file_checkboxes[file_path].get()):

                    # Get all sheets and filter out ignored ones
                    if file_path in self.files_data:
                        sheets = list(self.files_data[file_path].keys())
                    else:
                        # Load sheets if not already loaded
                        try:
                            reader = ExcelReader(file_path)
                            all_sheets = reader.get_sheet_names()
                            sheets = [s for s in all_sheets if not self.is_sheet_ignored(s)]
                            reader.close()
                        except:
                            sheets = []

                    if sheets:  # Only add if file has valid sheets
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

                file_display_name = os.path.basename(file_path)
                self.status_var.set(f"Processing file: {file_display_name}")

                if conversion_method == "direct":
                    # Direct method - fastest, no Excel app needed
                    from pdf_converter_direct import PDFConverterDirect

                    # Get folder prefix for file naming
                    folder_prefix = self.folder_names.get(file_path, "")

                    # Use direct conversion (no Excel app)
                    self.current_sheet_var.set(f"📄 Reading: {file_display_name}")
                    self.status_var.set(f"File {files_to_process.index((file_path, sheets_to_convert)) + 1}/{len(files_to_process)}")

                    converter = PDFConverterDirect(
                        enable_watermark=self.enable_watermark_var.get(),
                        watermark_opacity=0.3,
                        watermark_position="bottom-right"
                    )

                    # Convert all sheets directly (fastest)
                    results = converter.convert_excel_to_pdf_direct(file_path, sheets_to_convert, file_output_dir, folder_prefix)

                    # Update progress for each sheet
                    for sheet_name in sheets_to_convert:
                        self.current_sheet_var.set(f"📄 Converting: {sheet_name} (from {file_display_name})")

                        if results.get(sheet_name):
                            total_sheets_converted += 1
                            self.current_sheet_var.set(f"✅ Completed: {sheet_name}")
                        else:
                            print(f"Failed to convert sheet: {sheet_name}")
                            self.current_sheet_var.set(f"❌ Failed: {sheet_name}")

                        current_sheet += 1
                        progress = (current_sheet / total_sheets) * 100
                        self.progress_var.set(progress)
                        self.progress_percent_var.set(f"{progress:.1f}%")

                        # Small delay to show progress
                        import time
                        time.sleep(0.05)  # Faster delay for direct method

                elif conversion_method == "capture":
                    # Get folder prefix for file naming
                    folder_prefix = self.folder_names.get(file_path, "")

                    # Use optimized bulk conversion
                    self.current_sheet_var.set(f"📄 Opening: {file_display_name}")
                    self.status_var.set(f"File {files_to_process.index((file_path, sheets_to_convert)) + 1}/{len(files_to_process)}")

                    converter = PDFConverterCapture(
                        enable_watermark=self.enable_watermark_var.get(),
                        watermark_opacity=0.3,
                        watermark_position="bottom-right"
                    )

                    # Convert all sheets in one Excel session (faster)
                    results = converter.convert_excel_to_pdf(file_path, sheets_to_convert, file_output_dir, folder_prefix)

                    # Update progress for each sheet
                    for sheet_name in sheets_to_convert:
                        self.current_sheet_var.set(f"📄 Converting: {sheet_name} (from {file_display_name})")

                        if results.get(sheet_name):
                            total_sheets_converted += 1
                            self.current_sheet_var.set(f"✅ Completed: {sheet_name}")
                        else:
                            print(f"Failed to capture sheet: {sheet_name}")
                            self.current_sheet_var.set(f"❌ Failed: {sheet_name}")

                        current_sheet += 1
                        progress = (current_sheet / total_sheets) * 100
                        self.progress_var.set(progress)
                        self.progress_percent_var.set(f"{progress:.1f}%")

                        # Small delay to show progress
                        import time
                        time.sleep(0.1)

                else:
                    # Table conversion method
                    converter = PDFConverter(
                        preserve_formatting=self.preserve_format_var.get(),
                        bulk_mode=self.bulk_mode_var.get()
                    )

                    # Get folder prefix for file naming
                    folder_prefix = self.folder_names.get(file_path, "")

                    for sheet_name in sheets_to_convert:
                        # Update current sheet display
                        self.current_sheet_var.set(f"📄 Converting: {sheet_name} (from {file_display_name})")
                        self.status_var.set(f"File {files_to_process.index((file_path, sheets_to_convert)) + 1}/{len(files_to_process)}")

                        # Create output path in file's folder with prefix
                        safe_sheet_name = "".join(c for c in sheet_name if c.isalnum() or c in (' ', '-', '_')).rstrip()
                        if folder_prefix:
                            pdf_filename = f"{folder_prefix}_{safe_sheet_name}.pdf"
                        else:
                            pdf_filename = f"{safe_sheet_name}.pdf"
                        output_path = os.path.join(file_output_dir, pdf_filename)

                        try:
                            converter.convert_sheet_to_pdf(file_path, sheet_name, output_path)
                            total_sheets_converted += 1
                            self.current_sheet_var.set(f"✅ Completed: {sheet_name}")
                        except Exception as e:
                            print(f"Failed to convert sheet {sheet_name}: {str(e)}")
                            self.current_sheet_var.set(f"❌ Failed: {sheet_name}")

                        current_sheet += 1
                        progress = (current_sheet / total_sheets) * 100
                        self.progress_var.set(progress)
                        self.progress_percent_var.set(f"{progress:.1f}%")

                        # Small delay to show progress
                        import time
                        time.sleep(0.1)

            self.status_var.set(f"Conversion completed: {total_sheets_converted}/{total_sheets} sheets from {total_files} file(s)")
            self.current_sheet_var.set(f"🎉 All done! Converted {total_sheets_converted} sheets successfully")
            messagebox.showinfo("Success",
                              f"Successfully converted {total_sheets_converted}/{total_sheets} sheets from {total_files} file(s)\n"
                              f"Output saved to: {base_output_dir}")

        except Exception as e:
            messagebox.showerror("Error", f"Conversion failed: {str(e)}")
            self.status_var.set("Conversion failed")
            self.current_sheet_var.set("❌ Conversion failed")

        finally:
            self.convert_button.config(state='normal')
            self.progress_var.set(0)
            self.progress_percent_var.set("0%")
            # Keep the final message for a few seconds, then clear
            self.root.after(5000, lambda: self.current_sheet_var.set(""))
            
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
    import ttkbootstrap as tb
    import tkinter as tk
    import os, threading

    # Use dark theme
    root = tb.Window(themename="darkly")  # Dark mode theme
    app = ExcelToPDFApp(root)
    root.mainloop()
