import unittest
import export_pdf as exp
import os
import pathlib
import xlslib

class Tests(unittest.TestCase):

    def test_make_out_path(self):
        self.assertEqual(exp.make_out_path(os.path.abspath('a/b/c'), None, 'out'), str(pathlib.Path('out/a/b/c')))
        self.assertEqual(exp.make_out_path('a/b/c', None, 'out'), str(pathlib.Path('out/a/b/c')))
        self.assertEqual(exp.make_out_path('a/b/c', 'a', 'out'), str(pathlib.Path('out/b/c')))
        self.assertEqual(exp.make_out_path('a/b/c', 'a\\', 'out'), str(pathlib.Path('out/b/c')))

    def test_make_pdf_path(self):
        self.assertEqual(exp.make_pdf_path(os.path.abspath('a/b/c.txt'), None, 'out'), str(pathlib.Path('out/a/b/c.pdf')))
        self.assertEqual(exp.make_pdf_path('a/b/c.txt', None, 'out'), str(pathlib.Path('out/a/b/c.pdf')))
        self.assertEqual(exp.make_pdf_path('a/b/c.txt', 'a', 'out'), str(pathlib.Path('out/b/c.pdf')))
        self.assertEqual(exp.make_pdf_path('a/b/c.txt', 'a\\', 'out'), str(pathlib.Path('out/b/c.pdf')))

    def test_xlsrange(self):
        rg = xlslib.XlsRange("D2:E3")
        self.assertEqual(rg.entire_row, False)
        self.assertEqual(rg.entire_column, False)
        self.assertEqual(rg.start_row, 2)
        self.assertEqual(rg.end_row, 3)
        self.assertEqual(rg.start_col, 4)
        self.assertEqual(rg.end_col, 5)

        rg = xlslib.XlsRange("D:E")
        self.assertEqual(rg.entire_row, False)
        self.assertEqual(rg.entire_column, True)
        self.assertEqual(rg.start_row, 1)
        self.assertEqual(rg.end_row, 0x100_000)
        self.assertEqual(rg.start_col, 4)
        self.assertEqual(rg.end_col, 5)

        rg = xlslib.XlsRange("2:3")
        self.assertEqual(rg.entire_row, True)
        self.assertEqual(rg.entire_column, False)
        self.assertEqual(rg.start_row, 2)
        self.assertEqual(rg.end_row, 3)
        self.assertEqual(rg.start_col, 1)
        self.assertEqual(rg.end_col, 0x4000)

        rg = xlslib.XlsRange("D2")
        self.assertEqual(rg.entire_row, False)
        self.assertEqual(rg.entire_column, False)
        self.assertEqual(rg.start_row, 2)
        self.assertEqual(rg.end_row, 2)
        self.assertEqual(rg.start_col, 4)
        self.assertEqual(rg.end_col, 4)

if __name__ == '__main__':
    unittest.main()
