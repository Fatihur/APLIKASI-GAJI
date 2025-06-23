"""
Script untuk test multiple files functionality
"""

import os
import shutil
from excel_capture import ExcelCapture
from pdf_converter_capture import PDFConverterCapture
from create_sample_excel import create_sample_excel

def create_test_files():
    """Create multiple test Excel files"""
    print("📁 Creating test Excel files...")
    
    # Create original sample file
    create_sample_excel()
    
    # Create additional test files by copying and modifying
    test_files = []
    
    # File 1: Original sample_data.xlsx
    if os.path.exists("sample_data.xlsx"):
        test_files.append("sample_data.xlsx")
        print("   ✅ sample_data.xlsx")
    
    # File 2: Copy as test_file_2.xlsx
    if os.path.exists("sample_data.xlsx"):
        shutil.copy2("sample_data.xlsx", "test_file_2.xlsx")
        test_files.append("test_file_2.xlsx")
        print("   ✅ test_file_2.xlsx")
    
    # File 3: Copy as company_report.xlsx
    if os.path.exists("sample_data.xlsx"):
        shutil.copy2("sample_data.xlsx", "company_report.xlsx")
        test_files.append("company_report.xlsx")
        print("   ✅ company_report.xlsx")
    
    return test_files

def test_sheet_filtering():
    """Test sheet filtering (ignore sheets 1-9)"""
    print("\n🔍 Testing Sheet Filtering...")
    print("=" * 50)
    
    # Import the filtering function
    import sys
    sys.path.append('.')
    from main import ExcelToPDFApp
    
    # Create a dummy app to test filtering
    import tkinter as tk
    root = tk.Tk()
    root.withdraw()  # Hide window
    
    app = ExcelToPDFApp(root)
    
    # Test cases
    test_cases = [
        ("1", True),           # Should be ignored
        ("2", True),           # Should be ignored
        ("9", True),           # Should be ignored
        ("10", False),         # Should NOT be ignored
        ("1 Data", True),      # Should be ignored (starts with 1)
        ("2-Summary", True),   # Should be ignored (starts with 2)
        ("Data Karyawan", False),  # Should NOT be ignored
        ("Summary", False),    # Should NOT be ignored
        ("Sheet1", False),     # Should NOT be ignored (not just number)
    ]
    
    print("Testing sheet name filtering:")
    passed = 0
    for sheet_name, should_ignore in test_cases:
        result = app.is_sheet_ignored(sheet_name)
        status = "✅" if result == should_ignore else "❌"
        print(f"   {status} '{sheet_name}' -> Ignored: {result} (Expected: {should_ignore})")
        if result == should_ignore:
            passed += 1
    
    root.destroy()
    
    print(f"\nSheet filtering test: {passed}/{len(test_cases)} passed")
    return passed == len(test_cases)

def test_multiple_file_conversion():
    """Test converting multiple files with folder per file"""
    print("\n📦 Testing Multiple File Conversion...")
    print("=" * 50)
    
    # Create test files
    test_files = create_test_files()
    
    if not test_files:
        print("❌ No test files created")
        return False
    
    try:
        converter = PDFConverterCapture()
        base_output_dir = "multi_file_test_output"
        
        # Clean up previous test
        if os.path.exists(base_output_dir):
            shutil.rmtree(base_output_dir)
        
        os.makedirs(base_output_dir)
        
        total_files_processed = 0
        total_sheets_converted = 0
        
        for file_path in test_files:
            print(f"\n📄 Processing: {file_path}")
            
            # Create folder name for this file
            base_name = os.path.splitext(os.path.basename(file_path))[0]
            folder_name = f"converted_{base_name}"
            file_output_dir = os.path.join(base_output_dir, folder_name)
            
            # Read sheets from file
            capture = ExcelCapture()
            try:
                capture.open_excel_file(file_path)
                sheet_names = capture.get_sheet_names()
                
                # Filter sheets (ignore 1-9)
                filtered_sheets = []
                for sheet_name in sheet_names:
                    # Simple filtering logic
                    if not (sheet_name.strip() in ['1', '2', '3', '4', '5', '6', '7', '8', '9']):
                        filtered_sheets.append(sheet_name)
                
                print(f"   📋 Found {len(sheet_names)} sheets, {len(filtered_sheets)} after filtering")
                print(f"   📋 Sheets to convert: {filtered_sheets}")
                
                if not os.path.exists(file_output_dir):
                    os.makedirs(file_output_dir)
                
                # Convert each sheet
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
        print(f"\n🎯 CONVERSION SUMMARY:")
        print(f"   Files processed: {total_files_processed}/{len(test_files)}")
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
        print(f"❌ Multiple file conversion test failed: {str(e)}")
        return False

def cleanup_test_files():
    """Clean up test files"""
    print("\n🧹 Cleaning up test files...")
    
    test_files = ["test_file_2.xlsx", "company_report.xlsx"]
    for file_path in test_files:
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
                print(f"   ✅ Removed: {file_path}")
        except Exception as e:
            print(f"   ❌ Error removing {file_path}: {str(e)}")

def main():
    """Main test function"""
    print("🧪 Multiple Files Test Suite")
    print("=" * 60)
    
    tests = [
        ("Sheet Filtering", test_sheet_filtering),
        ("Multiple File Conversion", test_multiple_file_conversion)
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
    
    # Cleanup
    cleanup_test_files()
    
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
        print("🎉 All multiple files tests passed!")
    else:
        print("⚠️  Some tests failed. Check the output above for details.")

if __name__ == "__main__":
    main()
