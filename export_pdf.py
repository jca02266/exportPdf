from typing import Optional, Collection
import win32com.client as win32
import pywintypes
import os
import pathlib
import re
import shutil
import subprocess
import sys
import time
import inspect
import datetime
import normalize_excel
import normalize_word
from msoffice import Excel, Word

def info(*args):
    print(*args, file=sys.stderr)

def log(filename, rev='', sheet=''):
    with open('update.txt', 'a') as f:
        caller = inspect.getframeinfo(inspect.stack()[1][0])
        print(time.strftime("%Y/%m/%d %X"), f"{caller.filename}:{caller.lineno}", filename, rev, sheet, file=f)

def make_out_pathobj(path: str, basedir: Optional[str], outdir: str):
    relative_path = pathlib.Path(path)
    if relative_path.is_absolute():
        relative_path = relative_path.relative_to(os.getcwd())
    elif basedir:
        relative_path = relative_path.relative_to(basedir)

    out_path = pathlib.Path(outdir) / relative_path
    return out_path

def make_pdf_path(path: str, basedir: Optional[str], outdir: str):
    out_path = make_out_pathobj(path, basedir, outdir)
    pdf_path = out_path.parent / (out_path.stem + ".pdf")
    return str(pdf_path)

def make_out_path(path: str, basedir: Optional[str], outdir: str):
    out_path = make_out_pathobj(path, basedir, outdir)
    return str(out_path)



def export_pdf_word(wd, path: str, pdf_path: str = None, title: str = None,
                    visible: bool = False, basedir: str = None, outdir=".",
                    is_first = True):
    if pdf_path is None:
        pdf_path = make_pdf_path(path, basedir, outdir)
    info("", "=> ", pdf_path)

    os.makedirs(pathlib.Path(pdf_path).parent, exist_ok=True)
    wd.DisplayAlerts = False

    doc = wd.Documents.Open(os.path.abspath(path), ReadOnly=True)

    if title:
        # Not Bult*i*n , But Built*I*n
        doc.BuiltInDocumentProperties("Title").Value = title

    # update TOC
    for toc in doc.TablesOfContents:
        toc.Update()
    # update fields
    for story in doc.StoryRanges:
         story.Fields.Update()

    # At first, record settings of PDF exporting
    if is_first:
        doc.ExportAsFixedFormat(
            OutputFileName=os.path.abspath(pdf_path),
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
        # remove
        os.remove(pdf_path)

    # export PDF instead of ExportAsFixedFormat to record title property in generated PDF
    doc.SaveAs2(
        FileName=os.path.abspath(pdf_path),
        FileFormat=win32.constants.wdFormatPDF)

    # info("%d Pages" % doc.Content.Information(
    #     win32.constants.wdNumberOfPagesInDocument))

    doc.Saved = True
    doc.Close()

    return pdf_path

def export_pdf_excel(xl, path: str, pdf_path: str = None, title: str = None,
                     visible: bool = False, basedir: str = None, outdir=".",
                     target_sheets: Collection[str] = ()):
    if pdf_path is None:
        pdf_path = make_pdf_path(path, basedir, outdir)
    info("", "=> ", pdf_path)

    os.makedirs(pathlib.Path(pdf_path).parent, exist_ok=True)
    xl.DisplayAlerts = False

    try:
        wb = xl.Workbooks.Open(os.path.abspath(path), ReadOnly=True, UpdateLinks=0)
    except pywintypes.com_error as e:
        log("Error: file open error", path)
        return

    if title:
        # Not Bult*I*n , But Built*i*n
        wb.BuiltinDocumentProperties("Title").Value = title

    if len(target_sheets) == 0:
        target_sheets = list(target_sheets)
        for ws in wb.Worksheets:
            # Visible and Sheet tab color is not Black
            if not (ws.Visible and ws.Tab.ColorIndex != 1):
                info("", f"skip sheet {ws.Name} because color is {ws.Tab.ColorIndex}")
                continue
            target_sheets.append(ws.Name)

    wb.Worksheets(target_sheets).Select()

    wb.ActiveSheet.ExportAsFixedFormat(
        Type=win32.constants.xlTypePDF,
        Filename=os.path.abspath(pdf_path),
        Quality=win32.constants.xlQualityStandard,
        IncludeDocProperties=True,
        IgnorePrintAreas=False,
        OpenAfterPublish=False)

    wb.Saved = True
    wb.Close()

    return pdf_path

def pdf_to_jpeg(pdf_path: str, rev: str, outdir: pathlib.Path):
    info("", "=> ", outdir / rev, "(jpeg)")

    # mkdir tmp/[rev]
    tmpdir = os.path.join("tmp", rev)
    shutil.rmtree(tmpdir, ignore_errors=True)
    os.makedirs(tmpdir, exist_ok=True)

    # pdftoppm
    command = ["pdftoppm", "-jpeg", pdf_path, f"{tmpdir}/file"]
    res = subprocess.run(command, stdout=subprocess.PIPE)
    sys.stdout.buffer.write(res.stdout)

    # rename file-1.jpg to file-001.jpg
    for f in pathlib.Path(tmpdir).glob("file*.jpg"):
        g = re.sub(r'(.*file-)(\d{1,2})\.jpg', lambda m: m.group(1) + "%03d" % int(m.group(2)) + '.jpg', str(f))
        shutil.move(f, g)

    # mkdir pdf_path(exclude .pdf)/[rev] and move file-nnn.jpg to here
    os.makedirs(outdir, exist_ok=True)
    shutil.move(tmpdir, outdir)


is_first = True
def export(wd, xl, path: str, rev: str, **opt):
    global is_first

    info(rev, path)

    if path.endswith(".docx"):
        None
    elif path.endswith(".xlsx"):
        None
    else:
        info("", 'Ignore file')
        return

    # set the filename excluding the suffix into outdir
    outfile = pathlib.Path(make_out_path(path, None, opt['outdir']))
    outdir = outfile.parent / outfile.stem
    if os.path.exists(outdir / rev):
        info("", 'Skip: already output')
        return

    pdf_path = None
    if path.endswith(".docx"):
        normalize_word.normalize_file(wd, path, outfile)

        pdf_path = export_pdf_word(wd, path, is_first=is_first, **opt)
        is_first = False

    elif path.endswith(".xlsx"):
        normalize_excel.normalize_file(xl, path, outfile)

        pdf_path = export_pdf_excel(xl, path, **opt)
    else:
        info("", f"Ignore file: {path}")


    if pdf_path:
        pdf_to_jpeg(pdf_path, rev, outdir)

if __name__ == '__main__':
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument('files', metavar='PATH', type=str, nargs='+',
                        help='paths of office document (xlsx or docx)')
    parser.add_argument('--out', dest='outdir', action='store', default='out',
                        help='out directory (default: out)')
    parser.add_argument('--rev', dest='rev', action='store', default=None,
                        help='revision number to save the image (default: 1)')
    parser.add_argument('--title', dest='title', action='store', default=None,
                        help='set TITLE text property of the MS Word file')

    args = parser.parse_args()
    visible = True

    with Word(visible) as wd, \
        Excel(visible) as xl:

        for path in args.files:
            if os.path.isdir(path):
                continue

            opt = {"outdir": args.outdir,
                    "title": args.title }

            if not args.rev:
                args.rev = datetime.datetime.fromtimestamp(os.path.getmtime(path)).strftime("%Y%m%d-%H%M%S")
            export(wd, xl, path, args.rev, **opt)
