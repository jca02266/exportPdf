import re

class XlsRange():
  # $A$1:$B$2 にマッチ
  range_regexp = re.compile(r'\$?(?P<startCol>[A-Z]+)\$?(?P<startRow>[0-9]+):\$?(?P<endCol>[A-Z]+)\$?(?P<endRow>[0-9]+)')
  # $A:$B にマッチ
  range_col_regexp = re.compile(r'\$?(?P<startCol>[A-Z]+):\$?(?P<endCol>[A-Z]+)')
  # $1:$2 にマッチ
  range_row_regexp = re.compile(r'\$?(?P<startRow>[0-9]+):\$?(?P<endRow>[0-9]+)')
  # $A$1 にマッチ
  range_cell_regexp = re.compile(r'\$?(?P<startCol>[A-Z]+)\$?(?P<startRow>[0-9]+)')

  def __init__(self, range):
    self.isValid = True
    self.entire_row = False
    self.entire_column = False

    m = XlsRange.range_cell_regexp.fullmatch(range)
    if m:
      self.start_col, self.start_row = m.group('startCol', 'startRow')
      self.start_col, self.start_row = col2int(self.start_col), int(self.start_row)
      self.end_col, self.end_row = self.start_col, self.start_row
      return

    m = XlsRange.range_row_regexp.fullmatch(range)
    if m:
      self.entire_row = True
      self.start_row, self.end_row = m.group('startRow', 'endRow')
      self.start_col, self.start_row, self.end_col, self.end_row = \
        1, int(self.start_row), 0x4000, int(self.end_row)
      return

    m = XlsRange.range_col_regexp.fullmatch(range)
    if m:
      self.entire_column = True
      self.start_col, self.end_col = m.group('startCol', 'endCol')
      self.start_col, self.start_row, self.end_col, self.end_row = \
        col2int(self.start_col), 1, col2int(self.end_col), 0x100000
      return

    m = XlsRange.range_regexp.fullmatch(range)
    if m:
      self.start_col, self.start_row, self.end_col, self.end_row = m.group('startCol', 'startRow', 'endCol', 'endRow')
      self.start_col, self.start_row, self.end_col, self.end_row = \
        col2int(self.start_col), int(self.start_row), col2int(self.end_col), int(self.end_row)
      return

    self.isValid = False
    return

def col2int(s: str):
  """
  utility function

  return column index consists of str

  A -> 1
  Z -> 26
  AA -> 27
  AB -> 28
  """
  weight = 1
  n = 0
  nums = []
  nums_list = [chr(ord('0') + n) for n in range(10)]
  list_s = list(s)
  while list_s:
    if list_s[-1] in nums_list:
      nums.insert(0, list_s.pop())
      continue
    n += (ord(list_s.pop()) - ord('A')+1) * weight
    weight *= 26
  if nums:
    # (row, column)
    return (int(''.join(nums)), n)
  else:
    return n
