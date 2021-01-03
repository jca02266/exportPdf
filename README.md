export_pdf

# Requirement

Microsoft Office 2013 or later
Poppler (to convert pdf to jpeg)
  https://blog.alivate.com.au/poppler-windows/

# Setup

```
pip install pywin32
```

# Usage

```
python export_pdf test/book1.xlsx
==> out/test/book1.xlsx
```

```
python export_pdf.py *.docx *.xlsx --out out
==> out/test/book1.xlsx
```
