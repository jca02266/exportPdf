import unittest
import export_pdf
import os


class Test_export_pdf(unittest.TestCase):
    def test_make_pdf_path(self):
        self.assertEqual(export_pdf.make_pdf_path(
            os.getcwd() + r"\foo.xlsx", None, "."), r"foo.pdf")

        self.assertEqual(export_pdf.make_pdf_path(
            os.getcwd() + r"\bar\foo.xlsx", None, "."), r"bar\foo.pdf")

        self.assertEqual(export_pdf.make_pdf_path(
            os.getcwd() + r"\bar\foo.xlsx", None, r"qux\baz"), r"qux\baz\bar\foo.pdf")

        self.assertEqual(export_pdf.make_pdf_path(
            os.getcwd() + r"\bar\foo.xlsx", os.getcwd() + r"\bar", r"qux\baz"), r"qux\baz\foo.pdf")


if __name__ == '__main__':
    unittest.main()
