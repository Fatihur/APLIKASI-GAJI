"""
Script untuk test optimized conversion dengan prefix nama file
"""

import os
import time
import shutil
from pdf_converter_capture import PDFConverterCapture

def test_optimized_conversion():
    """Test optimized conversion (single Excel session)"""
    print("⚡ Testing Optimized Conversion...")
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
        
        root = tk.Tk()
        root.withdraw()
        app = ExcelToPDFApp(root)
        
        valid_sheets = [s for s in all_sheets if not app.is_sheet_ignored(s)]
        root.destroy()
        
        print(f"📋 Valid sheets to convert: {valid_sheets}")
        
        if not valid_sheets:
            print("❌ No valid sheets found")
            return False
        
        # Test optimized conversion
        output_dir = "optimized_test_output"
        if os.path.exists(output_dir):
            shutil.rmtree(output_dir)
        os.makedirs(output_dir)
        
        converter = PDFConverterCapture()
        
        # Test with custom prefix
        folder_prefix = "PAYROLL_2024"
        
        print(f"\n⚡ Starting optimized conversion with prefix: {folder_prefix}")
        start_time = time.time()
        
        results = converter.convert_excel_to_pdf(
            test_file, 
            valid_sheets, 
            output_dir, 
            folder_prefix
        )
        
        end_time = time.time()
        conversion_time = end_time - start_time
        
        print(f"⏱️  Conversion completed in {conversion_time:.2f} seconds")
        
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
        
        return success_count > 0 and conversion_time < 30  # Should be fast
        
    except Exception as e:
        print(f"❌ Optimized conversion test failed: {str(e)}")
        return False

def test_speed_comparison():
    """Test speed comparison between old and new method"""
    print("\n🏃 Testing Speed Comparison...")
    print("=" * 50)
    
    if not os.path.exists("sample_data.xlsx"):
        print("❌ No sample file available")
        return False
    
    try:
        # Get valid sheets
        from excel_reader import ExcelReader
        reader = ExcelReader("sample_data.xlsx")
        all_sheets = reader.get_sheet_names()
        reader.close()
        
        # Filter valid sheets (take first 2 for speed test)
        import sys
        sys.path.append('.')
        from main import ExcelToPDFApp
        import tkinter as tk
        
        root = tk.Tk()
        root.withdraw()
        app = ExcelToPDFApp(root)
        
        valid_sheets = [s for s in all_sheets if not app.is_sheet_ignored(s)][:2]  # Only 2 sheets for speed test
        root.destroy()
        
        if len(valid_sheets) < 2:
            print("❌ Need at least 2 valid sheets for comparison")
            return False
        
        print(f"📋 Testing with sheets: {valid_sheets}")
        
        converter = PDFConverterCapture()
        
        # Test 1: Old method (separate Excel sessions)
        print(f"\n🐌 Testing old method (separate Excel sessions)...")
        old_output_dir = "speed_test_old"
        if os.path.exists(old_output_dir):
            shutil.rmtree(old_output_dir)
        os.makedirs(old_output_dir)
        
        start_time = time.time()
        
        old_results = {}
        for sheet_name in valid_sheets:
            output_path = os.path.join(old_output_dir, f"{sheet_name}.pdf")
            success = converter.convert_single_sheet("sample_data.xlsx", sheet_name, output_path)
            old_results[sheet_name] = output_path if success else None
        
        old_time = time.time() - start_time
        
        # Test 2: New method (single Excel session)
        print(f"⚡ Testing new method (single Excel session)...")
        new_output_dir = "speed_test_new"
        if os.path.exists(new_output_dir):
            shutil.rmtree(new_output_dir)
        os.makedirs(new_output_dir)
        
        start_time = time.time()
        
        new_results = converter.convert_excel_to_pdf(
            "sample_data.xlsx", 
            valid_sheets, 
            new_output_dir, 
            "SPEED_TEST"
        )
        
        new_time = time.time() - start_time
        
        # Compare results
        print(f"\n📊 Speed Comparison Results:")
        print(f"   Old method: {old_time:.2f} seconds")
        print(f"   New method: {new_time:.2f} seconds")
        
        if new_time < old_time:
            improvement = ((old_time - new_time) / old_time) * 100
            print(f"   ⚡ Improvement: {improvement:.1f}% faster!")
        else:
            print(f"   ⚠️  New method is slower")
        
        # Check file quality
        old_success = sum(1 for r in old_results.values() if r and os.path.exists(r))
        new_success = sum(1 for r in new_results.values() if r and os.path.exists(r))
        
        print(f"   Old method success: {old_success}/{len(valid_sheets)}")
        print(f"   New method success: {new_success}/{len(valid_sheets)}")
        
        # Cleanup
        if os.path.exists(old_output_dir):
            shutil.rmtree(old_output_dir)
        if os.path.exists(new_output_dir):
            shutil.rmtree(new_output_dir)
        
        return new_time < old_time and new_success >= old_success
        
    except Exception as e:
        print(f"❌ Speed comparison test failed: {str(e)}")
        return False

def test_prefix_naming():
    """Test prefix naming functionality"""
    print("\n📝 Testing Prefix Naming...")
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
        print(f"📋 Testing with sheet: {test_sheet}")
        
        # Test different prefixes
        test_cases = [
            ("", f"sample_data_{test_sheet}.pdf"),  # No prefix
            ("PAYROLL_2024", f"PAYROLL_2024_{test_sheet}.pdf"),  # With prefix
            ("HR-DATA", f"HR-DATA_{test_sheet}.pdf"),  # With dash
            ("Company Report", f"Company Report_{test_sheet}.pdf"),  # With space
        ]
        
        converter = PDFConverterCapture()
        
        for prefix, expected_filename in test_cases:
            output_dir = f"prefix_test_{len(prefix) if prefix else 'empty'}"
            if os.path.exists(output_dir):
                shutil.rmtree(output_dir)
            os.makedirs(output_dir)
            
            print(f"\n🏷️  Testing prefix: '{prefix}'")
            
            results = converter.convert_excel_to_pdf(
                "sample_data.xlsx", 
                [test_sheet], 
                output_dir, 
                prefix
            )
            
            if results.get(test_sheet):
                actual_filename = os.path.basename(results[test_sheet])
                expected_clean = "".join(c for c in expected_filename if c.isalnum() or c in (' ', '-', '_', '.')).rstrip()
                actual_clean = actual_filename
                
                # Check if filename matches expected pattern
                if prefix and prefix in actual_filename:
                    print(f"   ✅ Prefix included: {actual_filename}")
                elif not prefix and not any(p in actual_filename for p in ["PAYROLL", "HR-DATA", "Company"]):
                    print(f"   ✅ No prefix: {actual_filename}")
                else:
                    print(f"   ❌ Unexpected filename: {actual_filename}")
                    print(f"      Expected pattern: {expected_clean}")
            else:
                print(f"   ❌ Conversion failed")
            
            # Cleanup
            if os.path.exists(output_dir):
                shutil.rmtree(output_dir)
        
        return True
        
    except Exception as e:
        print(f"❌ Prefix naming test failed: {str(e)}")
        return False

def main():
    """Main test function"""
    print("🧪 Optimized Conversion Test Suite")
    print("=" * 60)
    
    # Import tkinter for tests
    import tkinter as tk
    
    tests = [
        ("Optimized Conversion", test_optimized_conversion),
        ("Speed Comparison", test_speed_comparison),
        ("Prefix Naming", test_prefix_naming)
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
        print("🎉 All optimization tests passed!")
        print("\n⚡ Optimization Benefits:")
        print("   ✅ Faster conversion (single Excel session)")
        print("   ✅ Custom file prefixes")
        print("   ✅ Reduced Excel app overhead")
        print("   ✅ Better resource management")
    else:
        print("⚠️  Some tests failed. Check the output above for details.")

if __name__ == "__main__":
    main()
