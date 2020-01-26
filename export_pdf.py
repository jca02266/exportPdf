import win32com.client as win32
import os
from typing import Optional, Collection
import pathlib


def make_pdf_path(path: str, basedir: Optional[str], outdir: str):
    if basedir is None:
        basedir = os.getcwd()

    relative_path = pathlib.Path(path).relative_to(basedir)
    pdf_path = pathlib.Path(outdir) / relative_path.parent / \
        (relative_path.stem + ".pdf")
    return str(pdf_path)


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


def export_pdf_word(wd, path: str, pdf_path: str = None, title: str = None,
                    visible: bool = False, basedir: str = None, outdir="."):
    if pdf_path is None:
        pdf_path = os.path.abspath(make_pdf_path(path, basedir, outdir))

    print(pdf_path)
    os.makedirs(pathlib.Path(pdf_path).parent, exist_ok=True)
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


def export_pdf_excel(xl, path: str, pdf_path: str = None, title: str = None,
                     visible: bool = False, basedir: str = None, outdir=".",
                     target_sheets: Collection[str] = ()):
    if pdf_path is None:
        pdf_path = os.path.abspath(make_pdf_path(path, basedir, outdir))

    print(pdf_path)
    os.makedirs(pathlib.Path(pdf_path).parent, exist_ok=True)
    xl.DisplayAlerts = False

    wb = xl.Workbooks.Open(path, ReadOnly=True)

    if title:
        # Not Bult*I*n , But Built*i*n
        wb.BuiltinDocumentProperties("Title").Value = title

    if len(target_sheets) == 0:
        target_sheets = list(target_sheets)
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


def export_pdf(wd, xl, path: str, **opt):
    if path.endswith(".docx"):
        export_pdf_word(wd, path, **opt)
    elif path.endswith(".xlsx"):
        export_pdf_excel(xl, path, **opt)


if __name__ == '__main__':
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument('files', metavar='PATH', type=str, nargs='+',
                        help='paths of office document (xlsx or docx)')
    parser.add_argument('--out', dest='outdir', action='store', default='out',
                        help='out directory (default: out)')

    args = parser.parse_args()
    visible = True

    with Word(visible) as wd:
        with Excel(visible) as xl:

            for path in args.files:
                opt = {"outdir": args.outdir}
                export_pdf(wd, xl, os.path.abspath(path), **opt)
