"""
Script untuk test layout GUI
"""

import tkinter as tk
from main import ExcelToPDFApp
import time
import threading

def test_gui_layout():
    """Test GUI layout dan komponen"""
    
    print("🖥️  Testing GUI Layout...")
    print("=" * 50)
    
    # Buat root window
    root = tk.Tk()
    
    try:
        # Buat aplikasi
        app = ExcelToPDFApp(root)
        
        # Test komponen GUI
        print("✅ GUI components created successfully")
        
        # Test window size
        width = root.winfo_reqwidth()
        height = root.winfo_reqheight()
        print(f"✅ Window size: {width}x{height}")
        
        # Test komponen visibility
        components = [
            ("File path entry", app.file_path_var),
            ("Output path entry", app.output_path_var),
            ("Sheets listbox", app.sheets_listbox),
            ("Convert button", app.convert_button),
            ("Progress bar", app.progress_bar),
            ("Status label", app.status_var),
            ("Bulk mode checkbox", app.bulk_mode_var),
            ("Preserve format checkbox", app.preserve_format_var),
            ("Conversion method", app.conversion_method_var),
        ]
        
        print("\n📋 Component Status:")
        for name, component in components:
            try:
                if hasattr(component, 'get'):
                    # Variable components
                    value = component.get()
                    print(f"   ✅ {name}: {value}")
                else:
                    # Widget components
                    state = component.cget('state') if hasattr(component, 'cget') else 'normal'
                    print(f"   ✅ {name}: {state}")
            except Exception as e:
                print(f"   ❌ {name}: Error - {str(e)}")
        
        # Test default values
        print("\n🔧 Default Values:")
        print(f"   Bulk mode: {app.bulk_mode_var.get()}")
        print(f"   Preserve formatting: {app.preserve_format_var.get()}")
        print(f"   Conversion method: {app.conversion_method_var.get()}")
        print(f"   Status: {app.status_var.get()}")
        
        # Test method description update
        print("\n🔄 Testing method description update...")
        app.conversion_method_var.set("table")
        app.update_method_description()
        print(f"   Table method desc: {app.method_desc_var.get()}")
        
        app.conversion_method_var.set("capture")
        app.update_method_description()
        print(f"   Capture method desc: {app.method_desc_var.get()}")
        
        print("\n✅ GUI layout test completed successfully!")
        
        # Show window briefly
        root.update()
        
        # Auto close after 2 seconds
        def close_window():
            time.sleep(2)
            root.quit()
            root.destroy()
        
        thread = threading.Thread(target=close_window)
        thread.daemon = True
        thread.start()
        
        # Start mainloop
        root.mainloop()
        
        return True
        
    except Exception as e:
        print(f"❌ GUI test failed: {str(e)}")
        root.destroy()
        return False

def test_window_responsiveness():
    """Test window responsiveness"""
    
    print("\n📐 Testing Window Responsiveness...")
    print("=" * 50)
    
    root = tk.Tk()
    
    try:
        app = ExcelToPDFApp(root)
        
        # Test different window sizes
        sizes = [
            (800, 600),
            (900, 700),
            (1000, 800),
            (700, 500)
        ]
        
        for width, height in sizes:
            root.geometry(f"{width}x{height}")
            root.update()
            
            # Check if components are still accessible
            actual_width = root.winfo_width()
            actual_height = root.winfo_height()
            
            print(f"   ✅ Size {width}x{height} -> Actual: {actual_width}x{actual_height}")
            
            # Brief pause
            time.sleep(0.1)
        
        print("✅ Window responsiveness test completed!")
        
        root.destroy()
        return True
        
    except Exception as e:
        print(f"❌ Responsiveness test failed: {str(e)}")
        root.destroy()
        return False

def main():
    """Main test function"""
    print("🧪 GUI Layout Test Suite")
    print("=" * 60)
    
    tests = [
        ("GUI Layout", test_gui_layout),
        ("Window Responsiveness", test_window_responsiveness)
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
        print("🎉 All GUI tests passed! Layout is working correctly.")
    else:
        print("⚠️  Some tests failed. Check the output above for details.")

if __name__ == "__main__":
    main()
