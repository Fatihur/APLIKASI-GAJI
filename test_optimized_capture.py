"""
Script untuk test optimized capture method
"""

import os
import time
import shutil
from pdf_converter_capture import PDFConverterCapture

def test_optimized_capture():
    """Test optimized capture method (single Excel session)"""
    print("📸 Testing Optimized Capture Method...")
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
        
        # Test optimized capture
        output_dir = "optimized_capture_test"
        if os.path.exists(output_dir):
            shutil.rmtree(output_dir)
        os.makedirs(output_dir)
        
        converter = PDFConverterCapture()
        
        # Test with custom prefix
        folder_prefix = "CAPTURE_2024"
        
        print(f"\n📸 Starting optimized capture with prefix: {folder_prefix}")
        print("   ⚡ Excel will open once (hidden) for all sheets")
        start_time = time.time()
        
        results = converter.convert_excel_to_pdf(
            test_file, 
            valid_sheets, 
            output_dir, 
            folder_prefix
        )
        
        end_time = time.time()
        conversion_time = end_time - start_time
        
        print(f"⏱️  Optimized capture completed in {conversion_time:.2f} seconds")
        
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
        
        return success_count > 0 and conversion_time < 30  # Should be reasonably fast
        
    except Exception as e:
        print(f"❌ Optimized capture test failed: {str(e)}")
        return False

def test_capture_vs_old_method():
    """Test speed comparison: Optimized vs Old capture method"""
    print("\n⚡ Testing Capture Method: Optimized vs Old...")
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
        
        # Test 1: Optimized method (single Excel session)
        print(f"\n⚡ Testing optimized capture (single Excel session)...")
        optimized_output_dir = "speed_test_optimized"
        if os.path.exists(optimized_output_dir):
            shutil.rmtree(optimized_output_dir)
        os.makedirs(optimized_output_dir)
        
        start_time = time.time()
        
        converter = PDFConverterCapture()
        optimized_results = converter.convert_excel_to_pdf(
            "sample_data.xlsx", 
            valid_sheets, 
            optimized_output_dir, 
            "OPTIMIZED"
        )
        
        optimized_time = time.time() - start_time
        
        # Test 2: Old method (separate Excel sessions)
        print(f"🐌 Testing old capture method (separate Excel sessions)...")
        old_output_dir = "speed_test_old"
        if os.path.exists(old_output_dir):
            shutil.rmtree(old_output_dir)
        os.makedirs(old_output_dir)
        
        start_time = time.time()
        
        old_results = {}
        for sheet_name in valid_sheets:
            output_path = os.path.join(old_output_dir, f"OLD_{sheet_name}.pdf")
            success = converter.convert_single_sheet("sample_data.xlsx", sheet_name, output_path)
            old_results[sheet_name] = output_path if success else None
        
        old_time = time.time() - start_time
        
        # Compare results
        print(f"\n📊 Capture Method Comparison:")
        print(f"   Optimized method: {optimized_time:.2f} seconds (single Excel session)")
        print(f"   Old method: {old_time:.2f} seconds (separate Excel sessions)")
        
        if optimized_time < old_time:
            improvement = ((old_time - optimized_time) / old_time) * 100
            print(f"   ⚡ Optimized method is {improvement:.1f}% faster!")
        else:
            print(f"   ⚠️  Optimized method is slower")
        
        # Check file quality
        optimized_success = sum(1 for r in optimized_results.values() if r and os.path.exists(r))
        old_success = sum(1 for r in old_results.values() if r and os.path.exists(r))
        
        print(f"   Optimized method success: {optimized_success}/{len(valid_sheets)}")
        print(f"   Old method success: {old_success}/{len(valid_sheets)}")
        
        # Cleanup
        if os.path.exists(optimized_output_dir):
            shutil.rmtree(optimized_output_dir)
        if os.path.exists(old_output_dir):
            shutil.rmtree(old_output_dir)
        
        return optimized_time <= old_time and optimized_success >= old_success
        
    except Exception as e:
        print(f"❌ Capture comparison test failed: {str(e)}")
        return False

def test_capture_quality():
    """Test quality of optimized capture output"""
    print("\n📋 Testing Optimized Capture Quality...")
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
        
        # Test optimized capture
        output_dir = "quality_test_capture"
        if os.path.exists(output_dir):
            shutil.rmtree(output_dir)
        os.makedirs(output_dir)
        
        converter = PDFConverterCapture()
        
        results = converter.convert_excel_to_pdf(
            "sample_data.xlsx", 
            [test_sheet], 
            output_dir, 
            "QUALITY_CAPTURE"
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
            
            # Check file size is reasonable (capture files are usually larger)
            size_ok = 10000 < file_size < 10000000  # Between 10KB and 10MB
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
    print("🧪 Optimized Capture Method Test Suite")
    print("=" * 60)
    
    tests = [
        ("Optimized Capture", test_optimized_capture),
        ("Capture Comparison", test_capture_vs_old_method),
        ("Capture Quality", test_capture_quality)
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
        print("🎉 All optimized capture tests passed!")
        print("\n📸 Optimized Capture Benefits:")
        print("   ✅ Single Excel session (faster)")
        print("   ✅ Hidden Excel app (no visual distraction)")
        print("   ✅ Preserves exact Excel layout")
        print("   ✅ Custom file prefixes")
        print("   ✅ Better resource management")
        print("   ✅ More stable conversion")
    else:
        print("⚠️  Some tests failed. Check the output above for details.")

if __name__ == "__main__":
    main()
