import unittest
from sheet import *
from openpyxl import load_workbook

class TestSheet(unittest.TestCase):
    def setUp(self):
        # testfile is a spreadsheet containing the first four rows of MOCK_DATA.xlsx
        testfile = 'test.xlsx'
        wb = load_workbook(testfile, read_only=True)
        self.sheet = Sheet(wb.active)
    
    def test_lookup_unique_entry(self):
        result = [['1', 'Sigismund','Enbury', 'senbury0@bluehost.com', 'Male', '143.78.8.105']]
        self.assertEqual(self.sheet.lookup('id', '1'), result)
        
    def test_lookup_nonunique_entry(self):
        result = [['1', 'Sigismund','Enbury', 'senbury0@bluehost.com', 'Male', '143.78.8.105'],
                  ['2', 'Gaby', 'Blowing', 'gblowing1@apache.org', 'Male', '179.189.205.66'],
                  ['3', 'Delainey', 'Odde', 'dodde2@lulu.com', 'Male', '18.11.227.140']
                 ]
        self.assertEqual(self.sheet.lookup('gender', 'Male'), result)

    def test_lookup_no_key(self):
        with self.assertRaises(KeyError) as cm:
            self.sheet.lookup('notKey', 0)
        err_message = str(cm.exception)
        self.assertEqual(err_message, "'No field named: notKey'")

    def test_lookup_no_value(self):
        with self.assertRaises(KeyError) as cm:
            self.sheet.lookup('id', 'purple')
        err_message = str(cm.exception)
        self.assertEqual(err_message, "'purple does not exist'")
        
    def test_lookup_ID(self):
        result = ['1', 'Sigismund','Enbury', 'senbury0@bluehost.com', 'Male', '143.78.8.105']
        self.assertEqual(self.sheet.lookup_ID('1'), result)
        
if __name__ == '__main__':
    unittest.main()