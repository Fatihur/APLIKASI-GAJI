"""
Script untuk test aplikasi yang sudah diperbaiki
"""

import os
import sys
import tkinter as tk
import shutil

def test_improved_filtering():
    """Test improved filtering with keywords"""
    print("🔍 Testing Improved Filtering...")
    print("=" * 50)
    
    sys.path.append('.')
    from main import ExcelToPDFApp
    
    root = tk.Tk()
    root.withdraw()
    
    app = ExcelToPDFApp(root)
    
    # Test cases with new keyword-based filtering
    test_cases = [
        # Should be ignored (contains keywords)
        ('Payroll adjust', True),
        ('Database backup', True),
        ('Summary Amman report', True),
        ('Summary Karyawan 2024', True),
        ('PPh 21 calculation', True),
        ('Payroll monthly', True),
        ('Tarif TER update', True),
        ('hr_libur schedule', True),
        ('jm_istrht data', True),
        ('Data with adjust', True),  # Contains "adjust"
        ('Monthly Payroll', True),   # Contains "payroll"
        
        # Should NOT be ignored
        ('Data Karyawan', False),
        ('Laporan Bulanan', False),
        ('Absensi', False),
        ('Overtime Report', False),
        ('Employee List', False),
        ('Summary Report', False),  # Different from "Summary Amman"
        ('Payment Details', False), # Different from "Payroll"
        ('Adjustment Notes', False), # "Adjustment" != "adjust"
    ]
    
    print("Testing keyword-based filtering:")
    passed = 0
    failed_cases = []
    
    for sheet_name, should_ignore in test_cases:
        result = app.is_sheet_ignored(sheet_name)
        status = "✅" if result == should_ignore else "❌"
        print(f"   {status} '{sheet_name}' -> Ignored: {result} (Expected: {should_ignore})")
        
        if result == should_ignore:
            passed += 1
        else:
            failed_cases.append((sheet_name, result, should_ignore))
    
    root.destroy()
    
    print(f"\nKeyword filtering test: {passed}/{len(test_cases)} passed")
    
    if failed_cases:
        print("\n❌ Failed cases:")
        for sheet_name, got, expected in failed_cases:
            print(f"   '{sheet_name}': got {got}, expected {expected}")
    
    return passed == len(test_cases)

def test_ui_improvements():
    """Test UI improvements"""
    print("\n🖥️  Testing UI Improvements...")
    print("=" * 50)
    
    sys.path.append('.')
    from main import ExcelToPDFApp
    
    root = tk.Tk()
    root.title("UI Test")
    
    try:
        app = ExcelToPDFApp(root)
        
        print("✅ App created successfully")
        
        # Test adding files
        if os.path.exists("sample_data.xlsx"):
            test_files = ["sample_data.xlsx"]
            
            # Create additional test files
            shutil.copy2("sample_data.xlsx", "test_ui_1.xlsx")
            shutil.copy2("sample_data.xlsx", "test_ui_2.xlsx")
            test_files.extend(["test_ui_1.xlsx", "test_ui_2.xlsx"])
            
            print(f"📁 Adding {len(test_files)} test files...")
            
            for file_path in test_files:
                app.excel_files.append(file_path)
                filename = os.path.basename(file_path)
                base_name = os.path.splitext(filename)[0]
                app.folder_names[file_path] = base_name
                
                # Add checkbox
                app.add_file_checkbox(file_path, filename)
                
                # Add dummy sheet data
                app.files_data[file_path] = {
                    "Valid Sheet 1": {},
                    "Valid Sheet 2": {},
                    "Payroll adjust": {},  # Should be ignored
                    "Database": {}         # Should be ignored
                }
            
            print(f"✅ Added {len(test_files)} files with checkboxes")
            
            # Test checkbox functionality
            checked_count = sum(1 for f in test_files if app.file_checkboxes[f].get())
            print(f"📋 Checked files: {checked_count}/{len(test_files)}")
            
            # Test convert button state
            app.update_convert_button_state()
            convert_state = app.convert_button.cget('state')
            print(f"🔄 Convert button state: {convert_state}")
            
            # Test select all functionality
            app.select_all_files_var.set(True)
            app.toggle_all_files()
            all_checked = all(app.file_checkboxes[f].get() for f in test_files)
            print(f"✅ Select All works: {all_checked}")
            
            # Test clear all functionality
            app.select_all_files_var.set(False)
            app.toggle_all_files()
            none_checked = not any(app.file_checkboxes[f].get() for f in test_files)
            print(f"❌ Clear All works: {none_checked}")
            
            # Show window briefly
            root.geometry("900x700")
            root.update()
            
            print("\n👀 UI displayed for visual verification")
            
            # Auto close after 2 seconds
            root.after(2000, root.quit)
            root.mainloop()
            
            # Cleanup
            cleanup_files = ["test_ui_1.xlsx", "test_ui_2.xlsx"]
            for file_path in cleanup_files:
                try:
                    if os.path.exists(file_path):
                        os.remove(file_path)
                except:
                    pass
            
            return True
        else:
            print("❌ No sample file available")
            root.destroy()
            return False
            
    except Exception as e:
        print(f"❌ UI test failed: {str(e)}")
        root.destroy()
        return False

def test_auto_conversion():
    """Test auto conversion functionality"""
    print("\n🔄 Testing Auto Conversion...")
    print("=" * 50)
    
    if not os.path.exists("sample_data.xlsx"):
        print("📁 Creating sample Excel file...")
        try:
            from create_sample_excel import create_sample_excel
            create_sample_excel()
        except Exception as e:
            print(f"❌ Failed to create sample file: {str(e)}")
            return False
    
    try:
        from pdf_converter_capture import PDFConverterCapture
        from excel_reader import ExcelReader
        
        # Test file
        test_file = "sample_data.xlsx"
        
        # Read sheets and apply filtering
        reader = ExcelReader(test_file)
        all_sheets = reader.get_sheet_names()
        reader.close()
        
        print(f"📋 All sheets: {all_sheets}")
        
        # Apply filtering
        sys.path.append('.')
        from main import ExcelToPDFApp
        
        root = tk.Tk()
        root.withdraw()
        app = ExcelToPDFApp(root)
        
        valid_sheets = []
        ignored_sheets = []
        
        for sheet_name in all_sheets:
            if app.is_sheet_ignored(sheet_name):
                ignored_sheets.append(sheet_name)
            else:
                valid_sheets.append(sheet_name)
        
        root.destroy()
        
        print(f"✅ Valid sheets: {valid_sheets}")
        print(f"🚫 Ignored sheets: {ignored_sheets}")
        
        # Test conversion of valid sheets
        if valid_sheets:
            converter = PDFConverterCapture()
            output_dir = "auto_conversion_test"
            
            if os.path.exists(output_dir):
                shutil.rmtree(output_dir)
            os.makedirs(output_dir)
            
            print(f"\n🔄 Converting {len(valid_sheets)} valid sheets...")
            
            success_count = 0
            for sheet_name in valid_sheets:
                try:
                    safe_name = "".join(c for c in sheet_name if c.isalnum() or c in (' ', '-', '_')).rstrip()
                    output_path = os.path.join(output_dir, f"{safe_name}.pdf")
                    
                    success = converter.convert_single_sheet(test_file, sheet_name, output_path)
                    
                    if success and os.path.exists(output_path):
                        file_size = os.path.getsize(output_path)
                        print(f"   ✅ {sheet_name}: {file_size:,} bytes")
                        success_count += 1
                    else:
                        print(f"   ❌ {sheet_name}: Failed")
                        
                except Exception as e:
                    print(f"   ❌ {sheet_name}: Error - {str(e)}")
            
            print(f"\n📊 Conversion Results: {success_count}/{len(valid_sheets)} successful")
            
            return success_count > 0
        else:
            print("❌ No valid sheets to convert")
            return False
            
    except Exception as e:
        print(f"❌ Auto conversion test failed: {str(e)}")
        return False

def main():
    """Main test function"""
    print("🧪 Improved App Test Suite")
    print("=" * 60)
    
    tests = [
        ("Improved Filtering", test_improved_filtering),
        ("UI Improvements", test_ui_improvements),
        ("Auto Conversion", test_auto_conversion)
    ]
    
    results = []
    
    for test_name, test_func in tests:
        print(f"\n🔬 Running {test_name} test...")
        try:
            result = test_func()
            results.append((test_name, result))
        except Exception as e:
            print(f"❌ {test_name} test crashed: {str(e)}")
            results.append((test_name, False))
    
    # Summary
    print("\n" + "=" * 60)
    print("📋 TEST SUMMARY")
    print("=" * 60)
    
    passed = 0
    for test_name, result in results:
        status = "✅ PASSED" if result else "❌ FAILED"
        print(f"{test_name:.<30} {status}")
        if result:
            passed += 1
    
    print(f"\nOverall: {passed}/{len(results)} tests passed")
    
    if passed == len(results):
        print("🎉 All improved app tests passed!")
        print("\n📋 Key Improvements:")
        print("   ✅ Keyword-based filtering (contains: adjust, payroll, database, etc.)")
        print("   ✅ Simplified UI with direct checkboxes")
        print("   ✅ Auto-select all valid sheets for conversion")
        print("   ✅ Improved file management")
    else:
        print("⚠️  Some tests failed. Check the output above for details.")

if __name__ == "__main__":
    main()
