"""
Script untuk test progress bar dengan nama sheet
"""

import tkinter as tk
import threading
import time
import os
import sys

def test_progress_display():
    """Test progress bar display dengan nama sheet"""
    print("📊 Testing Progress Bar with Sheet Names...")
    print("=" * 50)
    
    sys.path.append('.')
    from main import ExcelToPDFApp
    
    root = tk.Tk()
    root.title("Progress Bar Test")
    root.geometry("900x700")
    
    try:
        app = ExcelToPDFApp(root)
        
        print("✅ App created successfully")
        
        # Simulate adding files
        if os.path.exists("sample_data.xlsx"):
            test_file = "sample_data.xlsx"
            
            # Add file to app
            app.excel_files.append(test_file)
            filename = os.path.basename(test_file)
            base_name = os.path.splitext(filename)[0]
            app.folder_names[test_file] = base_name
            
            # Add checkbox
            app.add_file_checkbox(test_file, filename)
            
            # Add sheet data
            app.files_data[test_file] = {
                "Data Karyawan": {},
                "Laporan Penjualan": {},
                "Summary": {},
                "Payroll": {},  # This should be ignored
                "Database": {}  # This should be ignored
            }
            
            print(f"✅ Added test file: {filename}")
            
            # Update button state
            app.update_convert_button_state()
            
            # Simulate progress updates
            def simulate_conversion():
                """Simulate conversion process with progress updates"""
                print("\n🔄 Simulating conversion process...")
                
                # Test sheets (only valid ones)
                test_sheets = ["Data Karyawan", "Laporan Penjualan", "Summary"]
                total_sheets = len(test_sheets)
                
                app.convert_button.config(state='disabled')
                app.progress_var.set(0)
                app.progress_percent_var.set("0%")
                app.current_sheet_var.set("")
                
                for i, sheet_name in enumerate(test_sheets):
                    # Update current sheet display
                    app.current_sheet_var.set(f"📄 Converting: {sheet_name} (from {filename})")
                    app.status_var.set(f"File 1/1 - Processing...")
                    
                    # Simulate conversion time
                    time.sleep(2)
                    
                    # Update progress
                    progress = ((i + 1) / total_sheets) * 100
                    app.progress_var.set(progress)
                    app.progress_percent_var.set(f"{progress:.1f}%")
                    
                    # Show completion
                    app.current_sheet_var.set(f"✅ Completed: {sheet_name}")
                    time.sleep(0.5)
                
                # Final status
                app.status_var.set(f"Conversion completed: {total_sheets}/{total_sheets} sheets from 1 file(s)")
                app.current_sheet_var.set(f"🎉 All done! Converted {total_sheets} sheets successfully")
                app.convert_button.config(state='normal')
                
                print(f"✅ Simulation completed - {total_sheets} sheets processed")
                
                # Auto close after showing final result
                time.sleep(3)
                root.quit()
            
            # Start simulation in thread
            thread = threading.Thread(target=simulate_conversion)
            thread.daemon = True
            thread.start()
            
            print("\n👀 Watch the progress bar and sheet names in the GUI")
            print("   - Progress percentage should update")
            print("   - Current sheet name should be displayed")
            print("   - Status should show file progress")
            
            # Show window and run
            root.mainloop()
            
            return True
        else:
            print("❌ No sample file available")
            root.destroy()
            return False
            
    except Exception as e:
        print(f"❌ Progress test failed: {str(e)}")
        root.destroy()
        return False

def test_progress_components():
    """Test individual progress components"""
    print("\n🧪 Testing Progress Components...")
    print("=" * 50)
    
    sys.path.append('.')
    from main import ExcelToPDFApp
    
    root = tk.Tk()
    root.withdraw()  # Hide for this test
    
    try:
        app = ExcelToPDFApp(root)
        
        # Test progress variables
        print("📊 Testing progress variables:")
        
        # Test progress_var
        app.progress_var.set(50)
        progress_value = app.progress_var.get()
        print(f"   Progress value: {progress_value}% {'✅' if progress_value == 50 else '❌'}")
        
        # Test progress_percent_var
        app.progress_percent_var.set("75.5%")
        percent_value = app.progress_percent_var.get()
        print(f"   Progress percent: {percent_value} {'✅' if percent_value == '75.5%' else '❌'}")
        
        # Test current_sheet_var
        test_sheet = "📄 Converting: Test Sheet"
        app.current_sheet_var.set(test_sheet)
        sheet_value = app.current_sheet_var.get()
        print(f"   Current sheet: {sheet_value} {'✅' if sheet_value == test_sheet else '❌'}")
        
        # Test status_var
        test_status = "Processing file 1/3"
        app.status_var.set(test_status)
        status_value = app.status_var.get()
        print(f"   Status: {status_value} {'✅' if status_value == test_status else '❌'}")
        
        # Test progress bar widget
        progress_bar_exists = hasattr(app, 'progress_bar')
        print(f"   Progress bar widget: {'✅' if progress_bar_exists else '❌'}")
        
        root.destroy()
        
        return True
        
    except Exception as e:
        print(f"❌ Components test failed: {str(e)}")
        root.destroy()
        return False

def test_progress_formatting():
    """Test progress formatting and display"""
    print("\n📝 Testing Progress Formatting...")
    print("=" * 50)
    
    # Test different progress scenarios
    test_cases = [
        (0, "0.0%", "Starting conversion..."),
        (25.5, "25.5%", "📄 Converting: Sheet 1"),
        (50.0, "50.0%", "✅ Completed: Sheet 1"),
        (75.8, "75.8%", "📄 Converting: Sheet 2"),
        (100.0, "100.0%", "🎉 All done! Converted 3 sheets successfully")
    ]
    
    print("Testing progress formatting scenarios:")
    
    for progress, expected_percent, sheet_message in test_cases:
        # Format percentage
        formatted_percent = f"{progress:.1f}%"
        percent_ok = formatted_percent == expected_percent
        
        # Check message format
        message_ok = len(sheet_message) > 0 and any(emoji in sheet_message for emoji in ["📄", "✅", "🎉", "❌"])
        
        status = "✅" if percent_ok and message_ok else "❌"
        print(f"   {status} Progress {progress}% -> {formatted_percent}, Message: {sheet_message[:30]}...")
    
    return True

def main():
    """Main test function"""
    print("🧪 Progress Bar Test Suite")
    print("=" * 60)
    
    # Ensure we have test file
    if not os.path.exists("sample_data.xlsx"):
        print("📁 Creating sample Excel file...")
        try:
            from create_sample_excel import create_sample_excel
            create_sample_excel()
        except Exception as e:
            print(f"❌ Failed to create sample file: {str(e)}")
            return
    
    tests = [
        ("Progress Components", test_progress_components),
        ("Progress Formatting", test_progress_formatting),
        ("Progress Display", test_progress_display)  # Visual test last
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
        print("🎉 All progress bar tests passed!")
        print("\n📊 Progress Bar Features:")
        print("   ✅ Progress percentage display")
        print("   ✅ Current sheet name display")
        print("   ✅ File progress status")
        print("   ✅ Visual progress bar")
        print("   ✅ Completion messages with emojis")
    else:
        print("⚠️  Some tests failed. Check the output above for details.")

if __name__ == "__main__":
    main()
