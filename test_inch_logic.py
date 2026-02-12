import unittest
import tkinter as tk
import sys
import os

# Add current directory to path to import the app
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# Import the module. Since the filename has a hyphen, we need to import it dynamically or rename it.
# Python doesn't like hyphens in filenames for import.
# Using __import__ or importlib
import importlib.util
spec = importlib.util.spec_from_file_location("ScrewInputApp", "screw_input_app1-6.py")
screw_module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(screw_module)

class TestInchConversion(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.root = tk.Tk()
        
    @classmethod
    def tearDownClass(cls):
        cls.root.destroy()
        
    def setUp(self):
        self.app = screw_module.ScrewInputApp(self.root)
        
    def test_convert_mm_to_inch(self):
        # 25.4mm -> 1.00000 inch
        val = self.app.convert_mm_to_inch("25.4")
        self.assertEqual(val, "1")
        
        # 12.7mm -> 0.50000 inch
        val = self.app.convert_mm_to_inch("12.7")
        self.assertEqual(val, "0.5")
        
        # 0mm -> 0
        val = self.app.convert_mm_to_inch("0")
        self.assertEqual(val, "0")
        
        # 1mm -> 0.03937 inch
        val = self.app.convert_mm_to_inch("1")
        self.assertEqual(val, "0.03937")

    def test_apply_plating_value_logic(self):
        # Test logic: Min = Normal Min - Plating/1000
        # Max = (Min+Max)/2 - Plating*2/1000
        
        normal_data = {'min': '10.000', 'max': '10.100'}
        plating_um = 10 # 0.01 mm
        
        # Expected:
        # Min = 10.000 - 0.010 = 9.990
        # Avg = 10.050
        # Max = 10.050 - (0.010 * 2) = 10.030
        
        calc_data = self.app.apply_plating_value(normal_data, 'OUTER', plating_um)
        
        self.assertEqual(calc_data['min'], '9.99')
        self.assertEqual(calc_data['max'], '10.03')
        
    def test_format_value_inch(self):
        value_dict = {'min': '25.4', 'max': '50.8'}
        # 1 inch ~ 2 inch
        res = self.app.format_value_inch(value_dict, "Major")
        self.assertEqual(res, 'Major : 1" ~ 2"')
        
        # Outer Minor -> Max suffix
        value_dict_minor = {'max': '25.4'}
        res = self.app.format_value_inch(value_dict_minor, "Minor", "OUTER")
        self.assertEqual(res, 'Minor : 1" Max.')

if __name__ == '__main__':
    unittest.main()
