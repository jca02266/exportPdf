import win32com.client as win32
import os
from typing import List


def __make_pdf_path(path: str):
    basename, _ = os.path.splitext(os.path.basename(path))
    return basename + ".pdf"


class Excel:
    def __init__(self, visible=False):
        self.xl = win32.gencache.EnsureDispatch("Excel.Application")
        if visible:
            self.xl.Visible = visible

    def __enter__(self):
        return self.xl

    def __exit__(self, exception_type, exception_value, traceback):
        if not self.xl.Visible:
            self.xl.Quit()


class Word:
    def __init__(self, visible=False):
        self.wd = win32.gencache.EnsureDispatch("Word.Application")
        if visible:
            self.wd.Visible = visible

    def __enter__(self):
        return self.wd

    def __exit__(self, exception_type, exception_value, traceback):
        if not self.wd.Visible:
            self.wd.Quit()


def export_pdf_word(path: str, pdf_path: str = None, title: str = None,
                    visible: bool = False):
    if pdf_path is None:
        pdf_path = __make_pdf_path(path)

    with Word(visible) as wd:
        wd.DisplayAlerts = False

        doc = wd.Documents.Open(path, ReadOnly=True)

        if title:
            # Not Bult*i*n , But Built*I*n
            doc.BuiltInDocumentProperties("Title").Value = title

        doc.ExportAsFixedFormat(
            OutputFileName=os.path.join(doc.Path, pdf_path),
            ExportFormat=win32.constants.wdExportFormatPDF,
            OpenAfterExport=False,
            OptimizeFor=win32.constants.wdExportOptimizeForPrint,
            Range=win32.constants.wdExportAllDocument,
            IncludeDocProps=False,
            KeepIRM=False,
            CreateBookmarks=win32.constants.wdExportCreateNoBookmarks,
            DocStructureTags=False,
            BitmapMissingFonts=True,
            UseISO19005_1=False)

        print("%d Pages" % doc.Content.Information(
            win32.constants.wdNumberOfPagesInDocument))
        doc.Saved = True
        doc.Close()


def export_pdf_excel(path: str, pdf_path: str = None, title: str = None,
                     visible: bool = False, target_sheets: List[str] = []):
    if pdf_path is None:
        pdf_path = __make_pdf_path(path)

    with Excel(visible) as xl:
        xl.DisplayAlerts = False

        wb = xl.Workbooks.Open(path, ReadOnly=True)

        if title:
            # Not Bult*I*n , But Built*i*n
            wb.BuiltinDocumentProperties("Title").Value = title

        if len(target_sheets) == 0:
            for ws in wb.Worksheets:
                # Visible and Sheet tab color is not Black
                if ws.Visible and ws.Tab.ColorIndex != 1:
                    print(ws.Name, ws.Tab.ColorIndex)
                    target_sheets.append(ws.Name)

        wb.Worksheets(target_sheets).Select()

        wb.ActiveSheet.ExportAsFixedFormat(
            Type=win32.constants.xlTypePDF,
            Filename=os.path.join(wb.Path, pdf_path),
            Quality=win32.constants.xlQualityStandard,
            IncludeDocProperties=True,
            IgnorePrintAreas=False,
            OpenAfterPublish=False)
        wb.Saved = True
        wb.Close()


def export_pdf(path: str, **opt):
    if path.endswith(".docx"):
        export_pdf_word(path, **opt)
    elif path.endswith(".xlsx"):
        export_pdf_excel(path, **opt)


if __name__ == '__main__':
    import sys

    for path in sys.argv[1:]:
        basename, _ = os.path.splitext(os.path.basename(path))
        export_pdf(os.path.abspath(path))
