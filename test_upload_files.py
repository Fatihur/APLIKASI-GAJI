"""
Script untuk test upload files functionality
"""

import tkinter as tk
from tkinter import ttk
import os
import sys

def test_file_upload_gui():
    """Test file upload GUI functionality"""
    print("🧪 Testing File Upload GUI...")
    print("=" * 50)
    
    # Import main app
    sys.path.append('.')
    from main import ExcelToPDFApp
    
    # Create root window
    root = tk.Tk()
    root.title("Test File Upload")
    
    try:
        # Create app
        app = ExcelToPDFApp(root)
        
        # Test initial state
        print("✅ App created successfully")
        print(f"📊 Initial files count: {len(app.excel_files)}")
        print(f"📊 Initial checkboxes count: {len(app.file_checkboxes)}")
        
        # Simulate adding files programmatically
        test_files = []
        if os.path.exists("sample_data.xlsx"):
            test_files.append("sample_data.xlsx")
        
        if test_files:
            print(f"\n📁 Simulating file upload with: {test_files}")
            
            for file_path in test_files:
                if file_path not in app.excel_files:
                    app.excel_files.append(file_path)
                    
                    # Set default folder name
                    filename = os.path.basename(file_path)
                    base_name = os.path.splitext(filename)[0]
                    app.folder_names[file_path] = base_name
                    
                    print(f"📄 Adding file: {filename}")
                    
                    # Create checkbox
                    try:
                        app.add_file_checkbox(file_path, filename)
                        print(f"✅ Checkbox created for: {filename}")
                    except Exception as e:
                        print(f"❌ Error creating checkbox: {str(e)}")
                    
                    # Load sheets data
                    try:
                        from excel_reader import ExcelReader
                        reader = ExcelReader(file_path)
                        sheets_info = reader.get_sheets_info()
                        
                        # Filter sheets
                        filtered_sheets = {}
                        for sheet_name, info in sheets_info.items():
                            if not app.is_sheet_ignored(sheet_name):
                                filtered_sheets[sheet_name] = info
                        
                        app.files_data[file_path] = filtered_sheets
                        reader.close()
                        
                        print(f"✅ Loaded {len(filtered_sheets)} sheets: {list(filtered_sheets.keys())}")
                        
                    except Exception as e:
                        print(f"❌ Error loading sheets: {str(e)}")
                        app.files_data[file_path] = {}
            
            # Update button state
            app.update_convert_button_state()
            
            # Check final state
            print(f"\n📊 Final state:")
            print(f"   Files count: {len(app.excel_files)}")
            print(f"   Checkboxes count: {len(app.file_checkboxes)}")
            print(f"   Files data count: {len(app.files_data)}")
            
            # Test checkbox functionality
            print(f"\n🔘 Testing checkbox functionality:")
            for file_path in app.excel_files:
                if file_path in app.file_checkboxes:
                    checkbox_state = app.file_checkboxes[file_path].get()
                    print(f"   {os.path.basename(file_path)}: {'✅ Checked' if checkbox_state else '❌ Unchecked'}")
            
            # Test convert button state
            convert_state = app.convert_button.cget('state')
            print(f"\n🔄 Convert button state: {convert_state}")
            
            # Show window briefly for visual verification
            root.geometry("900x700")
            root.update()
            
            print("\n👀 GUI window displayed for visual verification")
            print("   Check if files are visible in the interface")
            
            # Keep window open for a few seconds
            root.after(3000, root.quit)  # Auto close after 3 seconds
            root.mainloop()
            
            return True
        else:
            print("❌ No test files available")
            root.destroy()
            return False
            
    except Exception as e:
        print(f"❌ Test failed: {str(e)}")
        root.destroy()
        return False

def test_checkbox_operations():
    """Test checkbox operations"""
    print("\n🔘 Testing Checkbox Operations...")
    print("=" * 50)
    
    sys.path.append('.')
    from main import ExcelToPDFApp
    
    root = tk.Tk()
    root.withdraw()  # Hide window for this test
    
    try:
        app = ExcelToPDFApp(root)
        
        # Add test files
        test_files = []
        if os.path.exists("sample_data.xlsx"):
            test_files.append("sample_data.xlsx")
        
        # Create copies for testing
        import shutil
        if test_files:
            shutil.copy2(test_files[0], "test_file_1.xlsx")
            shutil.copy2(test_files[0], "test_file_2.xlsx")
            test_files.extend(["test_file_1.xlsx", "test_file_2.xlsx"])
        
        if not test_files:
            print("❌ No test files available")
            root.destroy()
            return False
        
        # Add files to app
        for file_path in test_files:
            app.excel_files.append(file_path)
            filename = os.path.basename(file_path)
            base_name = os.path.splitext(filename)[0]
            app.folder_names[file_path] = base_name
            app.add_file_checkbox(file_path, filename)
            app.files_data[file_path] = {"Sheet1": {}, "Sheet2": {}}
        
        print(f"✅ Added {len(test_files)} test files")
        
        # Test 1: Check all files initially checked
        all_checked = all(app.file_checkboxes[f].get() for f in test_files)
        print(f"📋 All files initially checked: {'✅' if all_checked else '❌'}")
        
        # Test 2: Uncheck one file
        if test_files:
            app.file_checkboxes[test_files[0]].set(False)
            app.update_convert_button_state()
            
            checked_count = sum(1 for f in test_files if app.file_checkboxes[f].get())
            print(f"📋 After unchecking one file: {checked_count}/{len(test_files)} checked")
        
        # Test 3: Select all functionality
        app.select_all_files_var.set(True)
        app.toggle_all_files()
        
        all_checked_after_select_all = all(app.file_checkboxes[f].get() for f in test_files)
        print(f"📋 After 'Select All': {'✅' if all_checked_after_select_all else '❌'}")
        
        # Test 4: Clear all functionality
        app.select_all_files_var.set(False)
        app.toggle_all_files()
        
        none_checked = not any(app.file_checkboxes[f].get() for f in test_files)
        print(f"📋 After 'Clear All': {'✅' if none_checked else '❌'}")
        
        # Test 5: Convert button state
        app.update_convert_button_state()
        convert_disabled = app.convert_button.cget('state') == 'disabled'
        print(f"🔄 Convert button disabled when no files checked: {'✅' if convert_disabled else '❌'}")
        
        # Cleanup
        cleanup_files = ["test_file_1.xlsx", "test_file_2.xlsx"]
        for file_path in cleanup_files:
            try:
                if os.path.exists(file_path):
                    os.remove(file_path)
            except:
                pass
        
        root.destroy()
        
        # All tests should pass
        tests_passed = all([
            all_checked,
            checked_count == len(test_files) - 1,
            all_checked_after_select_all,
            none_checked,
            convert_disabled
        ])
        
        print(f"\n📊 Checkbox operations test: {'✅ PASSED' if tests_passed else '❌ FAILED'}")
        return tests_passed
        
    except Exception as e:
        print(f"❌ Checkbox operations test failed: {str(e)}")
        root.destroy()
        return False

def main():
    """Main test function"""
    print("🧪 File Upload Test Suite")
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
        ("File Upload GUI", test_file_upload_gui),
        ("Checkbox Operations", test_checkbox_operations)
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
        print("🎉 All file upload tests passed!")
    else:
        print("⚠️  Some tests failed. Check the output above for details.")

if __name__ == "__main__":
    main()
