from typing import Optional, Collection
import win32com.client as win32
import pywintypes
import os
import sys
import pathlib
import stat
import shutil
from xlslib import XlsRange, col2int

def info(*args):
    print(*args, file=sys.stderr)

# Excelの整形をする
#  印刷欄外のデータを削除する
def clear_cells_outof_printarea(sheet):
  printArea = XlsRange(sheet.PageSetup.PrintArea)
  if not printArea.isValid:
    # 印刷範囲が設定されていないなら削除するデータなし
    return False

  usedRange = XlsRange(sheet.UsedRange.Address)

  try:
    if printArea.end_col >= usedRange.end_col:
      return False

    r = sheet.Range(sheet.Columns(printArea.end_col+1), sheet.Columns(usedRange.end_col))
    r.Clear()

    info("", "", "clear out of PrintArea: ", sheet.Parent.Name, sheet.Name, printArea.end_col+1, usedRange.end_col)
    return True
  except (AttributeError, win32.pywintypes.com_error) as e:
    info("", "Error: ", sheet.Name, e, sheet.PageSetup.PrintArea, sheet.UsedRange.Address)
    raise RuntimeError("skip")

# Excelの整形をする
#  印刷欄外の図形を削除する
def delete_shape_outof_printarea(sheet):
  printArea = XlsRange(sheet.PageSetup.PrintArea)
  if not printArea.isValid:
    # 印刷範囲が設定されていないなら使用されているデータ範囲を使う
    printArea = XlsRange(sheet.UsedRange.Address)

  shapesCount = sheet.Shapes.Count
  if shapesCount == 0:
    return False

  indexes = []
  for i in range(1, shapesCount+1):
    shape = sheet.Shapes(i)
    if shape.TopLeftCell.Column < printArea.end_col:
      # 図形の左上のセル位置が印刷範囲の右端列より左なら対象外
      continue
    if shape.BottomRightCell.Column <= printArea.end_col:
      # 図形の右下のセル位置が印刷範囲の右端列と同じか左なら対象外
      continue

    info("", "",  "delete Shape: ", sheet.Parent.Name, sheet.Name, shape.TopLeftCell.Address)
    indexes.append(i)

  if len(indexes) == 0:
    return False

  # インデックス番号がずれないよう逆順に削除
  for i in reversed(indexes):
    sheet.Shapes(i).Delete()
  return True

def normalize_each_sheet(wb):
  dirty = False

  wb.Worksheets(1).Select # 複数シート選択状態を解除
  for sheet in wb.Worksheets:
    # タブの色を色なしにする
    if sheet.Tab.Color != False:
      dirty = True
      sheet.Tab.Color = False

    # 印刷範囲外のデータを削除する
    if clear_cells_outof_printarea(sheet):
      dirty = True

    # 印刷範囲外のShapeを削除する
    if delete_shape_outof_printarea(sheet):
      dirty = True

  return dirty

def normalize_file(xl, path: str, outfile: str):

    info("", "=> ", outfile, "(normalize)")
    os.makedirs(pathlib.Path(outfile).parent, exist_ok=True)

    # Excelファイル以外は単純コピー
    if not path.endswith(".xlsx"):
        shutil.copy2(path, outfile)
        os.chmod(outfile, stat.S_IWRITE)
        return

    xl.DisplayAlerts = False

    try:
        wb = xl.Workbooks.Open(os.path.abspath(path), ReadOnly=False, UpdateLinks=0)
    except pywintypes.com_error as e:
        info("", "Error: file open error", path, e)
        return

    dirty = normalize_each_sheet(wb)
    if dirty:
        wb.SaveAs(Filename=os.path.abspath(outfile))
    else:
        shutil.copy2(path, outfile)
        os.chmod(outfile, stat.S_IWRITE)

    wb.Saved = True
    wb.Close()
