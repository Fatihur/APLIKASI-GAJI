"""
Script untuk test capture functionality
"""

import os
from excel_capture import ExcelCapture
from pdf_converter_capture import PDFConverterCapture

def test_excel_capture():
    """Test Excel capture functionality"""
    
    excel_file = "sample_data.xlsx"
    
    if not os.path.exists(excel_file):
        print("❌ File sample_data.xlsx tidak ditemukan!")
        print("Jalankan: python create_sample_excel.py")
        return False
    
    print("🔍 Testing Excel Capture Functionality...")
    print("=" * 50)
    
    capture = ExcelCapture()
    
    try:
        # Test 1: Open Excel file
        print("📖 Test 1: Opening Excel file...")
        capture.open_excel_file(excel_file)
        print("✅ Excel file opened successfully")
        
        # Test 2: Get sheet names
        print("\n📋 Test 2: Getting sheet names...")
        sheet_names = capture.get_sheet_names()
        print(f"✅ Found {len(sheet_names)} sheets:")
        for i, sheet_name in enumerate(sheet_names, 1):
            print(f"   {i}. {sheet_name}")
        
        # Test 3: Capture individual sheets
        print("\n📸 Test 3: Capturing individual sheets...")
        output_dir = "capture_test_output"
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
            
        for sheet_name in sheet_names:
            try:
                print(f"   Capturing: {sheet_name}...")
                safe_name = "".join(c for c in sheet_name if c.isalnum() or c in (' ', '-', '_')).rstrip()
                output_path = os.path.join(output_dir, f"capture_{safe_name}.pdf")
                
                result = capture.capture_sheet_as_png(sheet_name, output_path)
                
                if result and os.path.exists(result):
                    file_size = os.path.getsize(result)
                    print(f"   ✅ Success: {result} ({file_size:,} bytes)")
                else:
                    print(f"   ❌ Failed: {sheet_name}")
                    
            except Exception as e:
                print(f"   ❌ Error capturing {sheet_name}: {str(e)}")
        
        print("\n🎉 Capture test completed!")
        return True
        
    except Exception as e:
        print(f"❌ Test failed: {str(e)}")
        return False
        
    finally:
        capture.close()

def test_pdf_converter_capture():
    """Test PDF converter dengan capture method"""
    
    excel_file = "sample_data.xlsx"
    
    if not os.path.exists(excel_file):
        print("❌ File sample_data.xlsx tidak ditemukan!")
        return False
    
    print("\n🔄 Testing PDF Converter with Capture Method...")
    print("=" * 50)
    
    try:
        # Test converter
        converter = PDFConverterCapture()
        
        # Test sheets
        test_sheets = ["Data Karyawan", "Laporan Penjualan"]
        output_dir = "converter_capture_test"
        
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        print(f"📋 Testing conversion of {len(test_sheets)} sheets...")
        
        for sheet_name in test_sheets:
            try:
                print(f"   Converting: {sheet_name}...")
                
                safe_name = "".join(c for c in sheet_name if c.isalnum() or c in (' ', '-', '_')).rstrip()
                output_path = os.path.join(output_dir, f"converted_{safe_name}.pdf")
                
                success = converter.convert_single_sheet(excel_file, sheet_name, output_path)
                
                if success and os.path.exists(output_path):
                    file_size = os.path.getsize(output_path)
                    print(f"   ✅ Success: {output_path} ({file_size:,} bytes)")
                else:
                    print(f"   ❌ Failed: {sheet_name}")
                    
            except Exception as e:
                print(f"   ❌ Error converting {sheet_name}: {str(e)}")
        
        print("\n🎉 PDF Converter test completed!")
        return True
        
    except Exception as e:
        print(f"❌ Converter test failed: {str(e)}")
        return False

def test_bulk_conversion():
    """Test bulk conversion"""
    
    excel_file = "sample_data.xlsx"
    
    if not os.path.exists(excel_file):
        print("❌ File sample_data.xlsx tidak ditemukan!")
        return False
    
    print("\n📦 Testing Bulk Conversion...")
    print("=" * 50)
    
    try:
        converter = PDFConverterCapture()
        
        # Get all sheets
        capture = ExcelCapture()
        capture.open_excel_file(excel_file)
        all_sheets = capture.get_sheet_names()
        capture.close()
        
        output_dir = "bulk_conversion_test"
        
        print(f"📋 Converting all {len(all_sheets)} sheets...")
        
        results = converter.convert_excel_to_pdf(excel_file, all_sheets, output_dir)
        
        print("\n📊 Conversion Results:")
        success_count = 0
        for sheet_name, pdf_path in results.items():
            if pdf_path and os.path.exists(pdf_path):
                file_size = os.path.getsize(pdf_path)
                print(f"   ✅ {sheet_name}: {pdf_path} ({file_size:,} bytes)")
                success_count += 1
            else:
                print(f"   ❌ {sheet_name}: Failed")
        
        print(f"\n🎯 Summary: {success_count}/{len(all_sheets)} sheets converted successfully")
        return success_count == len(all_sheets)
        
    except Exception as e:
        print(f"❌ Bulk conversion test failed: {str(e)}")
        return False

def main():
    """Main test function"""
    print("🧪 Excel to PDF Capture Method - Test Suite")
    print("=" * 60)
    
    # Check if Excel is available
    try:
        import xlwings as xw
        print("✅ xlwings library available")
        
        # Try to create Excel app
        app = xw.App(visible=False, add_book=False)
        app.quit()
        print("✅ Excel application accessible")
        
    except Exception as e:
        print(f"❌ Excel not accessible: {str(e)}")
        print("   Make sure Microsoft Excel is installed")
        return
    
    # Run tests
    tests = [
        ("Excel Capture", test_excel_capture),
        ("PDF Converter Capture", test_pdf_converter_capture),
        ("Bulk Conversion", test_bulk_conversion)
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
        print("🎉 All tests passed! Capture method is working correctly.")
    else:
        print("⚠️  Some tests failed. Check the output above for details.")

if __name__ == "__main__":
    main()
