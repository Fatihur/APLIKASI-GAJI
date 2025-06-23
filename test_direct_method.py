"""
Script untuk test direct method conversion (tanpa Excel app)
"""

import os
import time
import shutil
from pdf_converter_direct import PDFConverterDirect

def test_direct_conversion():
    """Test direct conversion tanpa Excel app"""
    print("🚀 Testing Direct Conversion (No Excel App)...")
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
        # Test file
        test_file = "sample_data.xlsx"
        
        # Get valid sheets
        from excel_reader import ExcelReader
        reader = ExcelReader(test_file)
        all_sheets = reader.get_sheet_names()
        reader.close()
        
        # Filter valid sheets
        import sys
        sys.path.append('.')
        from main import ExcelToPDFApp
        import tkinter as tk
        
        root = tk.Tk()
        root.withdraw()
        app = ExcelToPDFApp(root)
        
        valid_sheets = [s for s in all_sheets if not app.is_sheet_ignored(s)]
        root.destroy()
        
        print(f"📋 Valid sheets to convert: {valid_sheets}")
        
        if not valid_sheets:
            print("❌ No valid sheets found")
            return False
        
        # Test direct conversion
        output_dir = "direct_test_output"
        if os.path.exists(output_dir):
            shutil.rmtree(output_dir)
        os.makedirs(output_dir)
        
        converter = PDFConverterDirect()
        
        # Test with custom prefix
        folder_prefix = "DIRECT_2024"
        
        print(f"\n🚀 Starting direct conversion with prefix: {folder_prefix}")
        print("   ⚡ No Excel application will be opened!")
        start_time = time.time()
        
        results = converter.convert_excel_to_pdf_direct(
            test_file, 
            valid_sheets, 
            output_dir, 
            folder_prefix
        )
        
        end_time = time.time()
        conversion_time = end_time - start_time
        
        print(f"⏱️  Direct conversion completed in {conversion_time:.2f} seconds")
        
        # Check results
        success_count = sum(1 for result in results.values() if result is not None)
        print(f"📊 Results: {success_count}/{len(valid_sheets)} sheets converted")
        
        # Check file names with prefix
        print(f"\n📁 Generated files:")
        for sheet_name, pdf_path in results.items():
            if pdf_path and os.path.exists(pdf_path):
                filename = os.path.basename(pdf_path)
                file_size = os.path.getsize(pdf_path)
                has_prefix = filename.startswith(folder_prefix)
                prefix_status = "✅" if has_prefix else "❌"
                print(f"   {prefix_status} {filename} ({file_size:,} bytes)")
            else:
                print(f"   ❌ {sheet_name}: Failed")
        
        return success_count > 0 and conversion_time < 10  # Should be very fast
        
    except Exception as e:
        print(f"❌ Direct conversion test failed: {str(e)}")
        return False

def test_speed_comparison_direct():
    """Test speed comparison: Direct vs Capture method"""
    print("\n⚡ Testing Speed Comparison: Direct vs Capture...")
    print("=" * 50)
    
    if not os.path.exists("sample_data.xlsx"):
        print("❌ No sample file available")
        return False
    
    try:
        # Get valid sheets (take first 2 for speed test)
        from excel_reader import ExcelReader
        reader = ExcelReader("sample_data.xlsx")
        all_sheets = reader.get_sheet_names()
        reader.close()
        
        # Filter valid sheets
        import sys
        sys.path.append('.')
        from main import ExcelToPDFApp
        import tkinter as tk
        
        root = tk.Tk()
        root.withdraw()
        app = ExcelToPDFApp(root)
        
        valid_sheets = [s for s in all_sheets if not app.is_sheet_ignored(s)][:2]  # Only 2 sheets
        root.destroy()
        
        if len(valid_sheets) < 2:
            print("❌ Need at least 2 valid sheets for comparison")
            return False
        
        print(f"📋 Testing with sheets: {valid_sheets}")
        
        # Test 1: Direct method (no Excel app)
        print(f"\n🚀 Testing direct method (no Excel app)...")
        direct_output_dir = "speed_test_direct"
        if os.path.exists(direct_output_dir):
            shutil.rmtree(direct_output_dir)
        os.makedirs(direct_output_dir)
        
        start_time = time.time()
        
        direct_converter = PDFConverterDirect()
        direct_results = direct_converter.convert_excel_to_pdf_direct(
            "sample_data.xlsx", 
            valid_sheets, 
            direct_output_dir, 
            "SPEED_DIRECT"
        )
        
        direct_time = time.time() - start_time
        
        # Test 2: Capture method (with Excel app)
        print(f"🐌 Testing capture method (with Excel app)...")
        capture_output_dir = "speed_test_capture"
        if os.path.exists(capture_output_dir):
            shutil.rmtree(capture_output_dir)
        os.makedirs(capture_output_dir)
        
        start_time = time.time()
        
        from pdf_converter_capture import PDFConverterCapture
        capture_converter = PDFConverterCapture()
        capture_results = capture_converter.convert_excel_to_pdf(
            "sample_data.xlsx", 
            valid_sheets, 
            capture_output_dir, 
            "SPEED_CAPTURE"
        )
        
        capture_time = time.time() - start_time
        
        # Compare results
        print(f"\n📊 Speed Comparison Results:")
        print(f"   Direct method: {direct_time:.2f} seconds (no Excel app)")
        print(f"   Capture method: {capture_time:.2f} seconds (with Excel app)")
        
        if direct_time < capture_time:
            improvement = ((capture_time - direct_time) / capture_time) * 100
            print(f"   🚀 Direct method is {improvement:.1f}% faster!")
        else:
            print(f"   ⚠️  Direct method is slower")
        
        # Check file quality
        direct_success = sum(1 for r in direct_results.values() if r and os.path.exists(r))
        capture_success = sum(1 for r in capture_results.values() if r and os.path.exists(r))
        
        print(f"   Direct method success: {direct_success}/{len(valid_sheets)}")
        print(f"   Capture method success: {capture_success}/{len(valid_sheets)}")
        
        # Cleanup
        if os.path.exists(direct_output_dir):
            shutil.rmtree(direct_output_dir)
        if os.path.exists(capture_output_dir):
            shutil.rmtree(capture_output_dir)
        
        return direct_time < capture_time and direct_success >= capture_success
        
    except Exception as e:
        print(f"❌ Speed comparison test failed: {str(e)}")
        return False

def test_direct_quality():
    """Test quality of direct conversion output"""
    print("\n📋 Testing Direct Conversion Quality...")
    print("=" * 50)
    
    if not os.path.exists("sample_data.xlsx"):
        print("❌ No sample file available")
        return False
    
    try:
        # Get one valid sheet for testing
        from excel_reader import ExcelReader
        reader = ExcelReader("sample_data.xlsx")
        all_sheets = reader.get_sheet_names()
        reader.close()
        
        # Filter valid sheets
        import sys
        sys.path.append('.')
        from main import ExcelToPDFApp
        import tkinter as tk
        
        root = tk.Tk()
        root.withdraw()
        app = ExcelToPDFApp(root)
        
        valid_sheets = [s for s in all_sheets if not app.is_sheet_ignored(s)][:1]  # Only 1 sheet
        root.destroy()
        
        if not valid_sheets:
            print("❌ No valid sheets found")
            return False
        
        test_sheet = valid_sheets[0]
        print(f"📋 Testing quality with sheet: {test_sheet}")
        
        # Test direct conversion
        output_dir = "quality_test_direct"
        if os.path.exists(output_dir):
            shutil.rmtree(output_dir)
        os.makedirs(output_dir)
        
        converter = PDFConverterDirect()
        
        results = converter.convert_excel_to_pdf_direct(
            "sample_data.xlsx", 
            [test_sheet], 
            output_dir, 
            "QUALITY_TEST"
        )
        
        if results.get(test_sheet):
            pdf_path = results[test_sheet]
            file_size = os.path.getsize(pdf_path)
            filename = os.path.basename(pdf_path)
            
            print(f"✅ PDF created successfully:")
            print(f"   📄 File: {filename}")
            print(f"   📊 Size: {file_size:,} bytes")
            
            # Check if file is valid PDF
            try:
                with open(pdf_path, 'rb') as f:
                    header = f.read(4)
                    is_pdf = header == b'%PDF'
                    print(f"   📋 Valid PDF: {'✅' if is_pdf else '❌'}")
            except:
                print(f"   📋 Valid PDF: ❌")
                is_pdf = False
            
            # Check file size is reasonable
            size_ok = 1000 < file_size < 1000000  # Between 1KB and 1MB
            print(f"   📊 Reasonable size: {'✅' if size_ok else '❌'}")
            
            # Cleanup
            if os.path.exists(output_dir):
                shutil.rmtree(output_dir)
            
            return is_pdf and size_ok
        else:
            print(f"❌ Failed to create PDF")
            return False
        
    except Exception as e:
        print(f"❌ Quality test failed: {str(e)}")
        return False

def main():
    """Main test function"""
    print("🧪 Direct Method Test Suite")
    print("=" * 60)
    
    tests = [
        ("Direct Conversion", test_direct_conversion),
        ("Speed Comparison", test_speed_comparison_direct),
        ("Direct Quality", test_direct_quality)
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
        print("🎉 All direct method tests passed!")
        print("\n🚀 Direct Method Benefits:")
        print("   ✅ No Excel application needed")
        print("   ✅ Fastest conversion speed")
        print("   ✅ Works without Excel installed")
        print("   ✅ Better resource usage")
        print("   ✅ More stable (no Excel crashes)")
    else:
        print("⚠️  Some tests failed. Check the output above for details.")

if __name__ == "__main__":
    main()
