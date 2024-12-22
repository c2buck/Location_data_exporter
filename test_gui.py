import unittest
import tkinter as tk
from location_data_v1 import root, browse_file, browse_folder, run

class TestGUI(unittest.TestCase):

    def setUp(self):
        self.root = root

    def test_gui_elements(self):
        # Test that all GUI elements are displayed correctly
        self.assertIsInstance(self.root, tk.Tk)
        self.assertIsNotNone(self.root.nametowidget('.!entry'))
        self.assertIsNotNone(self.root.nametowidget('.!button'))
        self.assertIsNotNone(self.root.nametowidget('.!label'))
        self.assertIsNotNone(self.root.nametowidget('.!checkbutton'))
        self.assertIsNotNone(self.root.nametowidget('.!radiobutton'))
        self.assertIsNotNone(self.root.nametowidget('.!combobox'))
        self.assertIsNotNone(self.root.nametowidget('.!progressbar'))
        self.assertIsNotNone(self.root.nametowidget('.!text'))

    def test_browse_file(self):
        # Test that the browse file button works
        browse_file()
        self.assertNotEqual(self.root.nametowidget('.!entry').get(), '')

    def test_browse_folder(self):
        # Test that the browse folder button works
        browse_folder()
        self.assertNotEqual(self.root.nametowidget('.!entry2').get(), '')

    def test_run(self):
        # Test that the run button works
        run()
        self.assertTrue(self.root.nametowidget('.!progressbar').get() > 0)

if __name__ == '__main__':
    unittest.main()
