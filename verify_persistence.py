import tkinter as tk
import os
import json
import importlib.util
import sys

# Import the app module handling the hyphen in filename
file_path = "screw_input_app1-6.py"
spec = importlib.util.spec_from_file_location("screw_input_app", file_path)
module = importlib.util.module_from_spec(spec)
sys.modules["screw_input_app"] = module
spec.loader.exec_module(module)
ScrewInputApp = module.ScrewInputApp

def test():
    print("Starting verification...")
    # Setup: ensure clean state (optional, but good for reliable test)
    # But user might have existing data, so let's backup if valid
    csv_exists = os.path.exists("screw_data.json")
    if csv_exists:
        os.rename("screw_data.json", "screw_data.json.bak")
    
    try:
        root = tk.Tk()
        # Prevent window from showing up effectively
        root.withdraw() 
        
        print("Initializing app first time...")
        app = ScrewInputApp(root)
        
        # 1. Verify default data loaded and saved
        key_count = len(app.screw_data)
        print(f"Initial keys count: {key_count}")
        
        if not os.path.exists("screw_data.json"):
            print("FAIL: screw_data.json was not created on init")
            return
        else:
            print("PASS: screw_data.json created")
            
        # 2. Add new data programmatically
        new_screw_name = "TEST_SCREW_999"
        new_data = {'type': 'OUTER', 'gauge': 'NORMAL', 'customer': 'SM', 'major': {'min': '1.0', 'max': '2.0'}}
        
        print(f"Adding {new_screw_name}...")
        if new_screw_name not in app.screw_data:
            app.screw_data[new_screw_name] = []
        app.screw_data[new_screw_name].append(new_data)
        app.save_screw_data()
        
        root.destroy()
        
        # 3. Restart and verify persistence
        print("Initializing app second time...")
        root2 = tk.Tk()
        root2.withdraw()
        app2 = ScrewInputApp(root2)
        
        if new_screw_name in app2.screw_data:
            print(f"PASS: {new_screw_name} found after restart")
        else:
            print(f"FAIL: {new_screw_name} NOT found after restart")
            
        root2.destroy()
        
    except Exception as e:
        print(f"ERROR: {e}")
    finally:
        # Cleanup / Restore
        if os.path.exists("screw_data.json"):
            os.remove("screw_data.json")
        if csv_exists:
            os.rename("screw_data.json.bak", "screw_data.json")
            print("Restored original screw_data.json")

if __name__ == "__main__":
    test()
