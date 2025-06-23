"""
Script untuk test konversi dengan checkbox functionality
"""

import os
import shutil
from excel_capture import ExcelCapture
from pdf_converter_capture import PDFConverterCapture

def test_checkbox_conversion():
    """Test konversi dengan checkbox functionality"""
    print("🔄 Testing Checkbox Conversion...")
    print("=" * 50)
    
    # Pastikan ada file test
    if not os.path.exists("sample_data.xlsx"):
        print("📁 Creating sample Excel file...")
        from create_sample_excel import create_sample_excel
        create_sample_excel()
    
    # Test files
    test_files = ["sample_data.xlsx"]
    
    # Copy untuk membuat multiple files
    if os.path.exists("sample_data.xlsx"):
        shutil.copy2("sample_data.xlsx", "test_payroll.xlsx")
        shutil.copy2("sample_data.xlsx", "test_hr.xlsx")
        test_files.extend(["test_payroll.xlsx", "test_hr.xlsx"])
    
    print(f"📋 Test files: {test_files}")
    
    try:
        # Simulate checkbox selection (all files checked)
        checked_files = test_files.copy()
        
        # Test conversion
        converter = PDFConverterCapture()
        base_output_dir = "checkbox_test_output"
        
        # Clean up previous test
        if os.path.exists(base_output_dir):
            shutil.rmtree(base_output_dir)
        
        os.makedirs(base_output_dir)
        
        total_files_processed = 0
        total_sheets_converted = 0
        
        for file_path in checked_files:
            print(f"\n📄 Processing checked file: {file_path}")
            
            # Create folder name for this file
            base_name = os.path.splitext(os.path.basename(file_path))[0]
            folder_name = f"folder_{base_name}"
            file_output_dir = os.path.join(base_output_dir, folder_name)
            
            # Read sheets from file and apply filtering
            capture = ExcelCapture()
            try:
                capture.open_excel_file(file_path)
                sheet_names = capture.get_sheet_names()
                
                # Apply new filtering logic
                from main import ExcelToPDFApp
                import tkinter as tk
                root = tk.Tk()
                root.withdraw()
                app = ExcelToPDFApp(root)
                
                filtered_sheets = []
                for sheet_name in sheet_names:
                    if not app.is_sheet_ignored(sheet_name):
                        filtered_sheets.append(sheet_name)
                
                root.destroy()
                
                print(f"   📋 Total sheets: {len(sheet_names)}")
                print(f"   📋 After filtering: {len(filtered_sheets)}")
                print(f"   📋 Sheets to convert: {filtered_sheets}")
                
                if not os.path.exists(file_output_dir):
                    os.makedirs(file_output_dir)
                
                # Convert each filtered sheet
                file_sheets_converted = 0
                for sheet_name in filtered_sheets:
                    try:
                        safe_sheet_name = "".join(c for c in sheet_name if c.isalnum() or c in (' ', '-', '_')).rstrip()
                        output_path = os.path.join(file_output_dir, f"{safe_sheet_name}.pdf")
                        
                        print(f"      Converting: {sheet_name}")
                        success = converter.convert_single_sheet(file_path, sheet_name, output_path)
                        
                        if success and os.path.exists(output_path):
                            file_size = os.path.getsize(output_path)
                            print(f"      ✅ Success: {output_path} ({file_size:,} bytes)")
                            file_sheets_converted += 1
                            total_sheets_converted += 1
                        else:
                            print(f"      ❌ Failed: {sheet_name}")
                            
                    except Exception as e:
                        print(f"      ❌ Error converting {sheet_name}: {str(e)}")
                
                print(f"   📊 File summary: {file_sheets_converted}/{len(filtered_sheets)} sheets converted")
                total_files_processed += 1
                
            except Exception as e:
                print(f"   ❌ Error processing file: {str(e)}")
                
            finally:
                capture.close()
        
        # Final summary
        print(f"\n🎯 CHECKBOX CONVERSION SUMMARY:")
        print(f"   Files processed: {total_files_processed}/{len(checked_files)}")
        print(f"   Total sheets converted: {total_sheets_converted}")
        print(f"   Output directory: {base_output_dir}")
        
        # List output structure
        if os.path.exists(base_output_dir):
            print(f"\n📁 Output Structure:")
            for item in os.listdir(base_output_dir):
                item_path = os.path.join(base_output_dir, item)
                if os.path.isdir(item_path):
                    pdf_files = [f for f in os.listdir(item_path) if f.endswith('.pdf')]
                    print(f"   📂 {item}/ ({len(pdf_files)} PDF files)")
                    for pdf_file in pdf_files:
                        pdf_path = os.path.join(item_path, pdf_file)
                        size = os.path.getsize(pdf_path)
                        print(f"      📄 {pdf_file} ({size:,} bytes)")
        
        return total_files_processed > 0 and total_sheets_converted > 0
        
    except Exception as e:
        print(f"❌ Checkbox conversion test failed: {str(e)}")
        return False
        
    finally:
        # Cleanup test files
        cleanup_files = ["test_payroll.xlsx", "test_hr.xlsx"]
        for file_path in cleanup_files:
            try:
                if os.path.exists(file_path):
                    os.remove(file_path)
                    print(f"   🧹 Cleaned up: {file_path}")
            except Exception as e:
                print(f"   ❌ Error cleaning {file_path}: {str(e)}")

def test_selective_checkbox():
    """Test selective checkbox (some checked, some unchecked)"""
    print("\n✅ Testing Selective Checkbox...")
    print("=" * 50)
    
    # Create test files
    test_files = []
    if os.path.exists("sample_data.xlsx"):
        shutil.copy2("sample_data.xlsx", "checked_file1.xlsx")
        shutil.copy2("sample_data.xlsx", "unchecked_file2.xlsx")
        shutil.copy2("sample_data.xlsx", "checked_file3.xlsx")
        test_files = ["checked_file1.xlsx", "unchecked_file2.xlsx", "checked_file3.xlsx"]
    
    # Simulate checkbox selection (only file1 and file3 checked)
    checked_files = ["checked_file1.xlsx", "checked_file3.xlsx"]
    unchecked_files = ["unchecked_file2.xlsx"]
    
    print(f"📋 All files: {test_files}")
    print(f"✅ Checked files: {checked_files}")
    print(f"❌ Unchecked files: {unchecked_files}")
    
    try:
        converter = PDFConverterCapture()
        base_output_dir = "selective_test_output"
        
        # Clean up previous test
        if os.path.exists(base_output_dir):
            shutil.rmtree(base_output_dir)
        
        os.makedirs(base_output_dir)
        
        # Only process checked files
        processed_files = []
        for file_path in checked_files:
            print(f"\n📄 Processing checked file: {file_path}")
            
            base_name = os.path.splitext(os.path.basename(file_path))[0]
            folder_name = f"output_{base_name}"
            file_output_dir = os.path.join(base_output_dir, folder_name)
            
            if not os.path.exists(file_output_dir):
                os.makedirs(file_output_dir)
            
            # Convert one sheet as test
            capture = ExcelCapture()
            try:
                capture.open_excel_file(file_path)
                sheet_names = capture.get_sheet_names()
                
                # Get first valid sheet
                from main import ExcelToPDFApp
                import tkinter as tk
                root = tk.Tk()
                root.withdraw()
                app = ExcelToPDFApp(root)
                
                valid_sheet = None
                for sheet_name in sheet_names:
                    if not app.is_sheet_ignored(sheet_name):
                        valid_sheet = sheet_name
                        break
                
                root.destroy()
                
                if valid_sheet:
                    output_path = os.path.join(file_output_dir, f"{valid_sheet}.pdf")
                    success = converter.convert_single_sheet(file_path, valid_sheet, output_path)
                    
                    if success:
                        processed_files.append(file_path)
                        print(f"   ✅ Converted: {valid_sheet}")
                    else:
                        print(f"   ❌ Failed: {valid_sheet}")
                
            finally:
                capture.close()
        
        # Verify unchecked files were not processed
        for file_path in unchecked_files:
            base_name = os.path.splitext(os.path.basename(file_path))[0]
            folder_name = f"output_{base_name}"
            folder_path = os.path.join(base_output_dir, folder_name)
            
            if not os.path.exists(folder_path):
                print(f"   ✅ Correctly ignored unchecked file: {file_path}")
            else:
                print(f"   ❌ Unexpectedly processed unchecked file: {file_path}")
        
        print(f"\n📊 Selective Test Results:")
        print(f"   Expected to process: {len(checked_files)}")
        print(f"   Actually processed: {len(processed_files)}")
        print(f"   Success: {len(processed_files) == len(checked_files)}")
        
        return len(processed_files) == len(checked_files)
        
    except Exception as e:
        print(f"❌ Selective checkbox test failed: {str(e)}")
        return False
        
    finally:
        # Cleanup
        cleanup_files = ["checked_file1.xlsx", "unchecked_file2.xlsx", "checked_file3.xlsx"]
        for file_path in cleanup_files:
            try:
                if os.path.exists(file_path):
                    os.remove(file_path)
            except:
                pass

def main():
    """Main test function"""
    print("🧪 Checkbox Conversion Test Suite")
    print("=" * 60)
    
    tests = [
        ("Checkbox Conversion", test_checkbox_conversion),
        ("Selective Checkbox", test_selective_checkbox)
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
        print("🎉 All checkbox conversion tests passed!")
    else:
        print("⚠️  Some tests failed. Check the output above for details.")

if __name__ == "__main__":
    main()
