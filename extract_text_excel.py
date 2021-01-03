from typing import Optional, Collection
import win32com.client as win32
import pywintypes
from msoffice import Excel, Word
import os

def info(*args):
    print(*args, file=sys.stderr)

def extract_text(sheet):
  used_range = sheet.UsedRange
  for row in used_range.Rows:
    for col in row.Columns:
      text = col.Value
      if text is None:
        continue
      yield text, 'cell', col.Address

  for shape in sheet.Shapes:
    textframe = shape.TextFrame
    try:
      characters = textframe.Characters
      text = characters().Text
      if not text:
        text = None
    except:
      text = None

    if text is None:
      continue
    yield text, 'shape', shape.TopLeftCell.Address

if __name__ == '__main__':
  with Excel(visible = True) as xl:
    wb = xl.Workbooks.Open(os.path.abspath("test/test1.xlsx"))
    for text,tot,addr in extract_text(wb.Sheets(1)):
      print(tot, addr, repr(text))
