Attribute VB_Name = "Module1"
Option Explicit

' �Q�Ɛݒ�
' Scripting.Dictionary -> Microsoft Scripting Runtime
' Microsoft Word X.X Document Library

Dim wd As New Word.Application

Sub ExportPdfWord(info As PdfInfo, Path As String)
    Dim doc As Word.Document

    wd.Visible = True
    
    Set doc = wd.Documents.Open(Path, ReadOnly:=True)

    wd.DisplayAlerts = wdAlertsNone
    doc.BuiltinDocumentProperties(wdPropertyTitle) = "�Ă���"
    
    doc.ExportAsFixedFormat _
        OutputFileName:=JoinPath(ThisWorkbook.Path, info.PdfFilename), _
        ExportFormat:=wdExportFormatPDF, _
        OpenAFterExport:=False, _
        OptimizeFor:=wdExportOptimizeForPrint, _
        Range:=wdExportAllDocument, _
        IncludeDocProps:=False, _
        KeepIRM:=False, _
        CreateBookmarks:=wdExportCreateNoBookmarks, _
        DocStructureTags:=False, _
        BitmapMissingFonts:=True, _
        UseISO19005_1:=False

    Debug.Print doc.Content.Information(wdNumberOfPagesInDocument) & "Pages"
    
    doc.Saved = True
    doc.Close
End Sub

'
' book�̃V�[�g�̂����z�� sheetNames �ɂȂ��V�[�g�̔z���Ԃ�
'
Function GetRemainingSheetNames(book As Workbook, SheetNames As Variant) As Variant
    Dim ary()
    Dim i As Integer
    i = 0
    
    Dim printSheets As Scripting.Dictionary
    Set printSheets = Utils.aryToDictionary(SheetNames)

    Dim sheet As Worksheet
    For Each sheet In book.Sheets
        Debug.Print sheet.name
        If Not printSheets.Exists(sheet.name) Then
            ReDim Preserve ary(i)
            ary(i) = sheet.name
        End If
    Next sheet
    
    GetRemainingSheetNames = ary
End Function

Sub ExportPdfExcel(info As PdfInfo, Path As String)
    Dim book As Workbook
    
    'Application.DisplayAlerts = False
    Set book = Workbooks.Open(Path, ReadOnly:=True)

    Dim remainingSheetNames()
    
    remainingSheetNames = GetRemainingSheetNames(book, info.SheetNames)
    
    book.Worksheets(info.SheetNames).Select
    ActiveSheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        filename:=JoinPath(ThisWorkbook.Path, info.PdfFilename), _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    '
    Debug.Print "*** �o�͂��Ȃ������V�[�g ***"
    Dim name As Variant
    For Each name In remainingSheetNames
        Debug.Print book.name, name
    Next name
    book.Saved = True
    book.Close
End Sub


Sub EachFile(dict As Scripting.Dictionary, WildCard As String)
    Dim filename As String
    Dim folderName As String
    
    folderName = Utils.Dirname(WildCard)
    
    ' �����J�n
    Debug.Print "*** �����t�@�C�� ***"
    filename = dir(WildCard)
    Do While filename <> ""
        If dict.Exists(filename) Then
            If filename Like "*.xlsx" Then
                Debug.Print filename, "Done"
                ExportPdfExcel dict(filename), Utils.JoinPath(folderName, filename)
            ElseIf filename Like "*.docx" Then
                Debug.Print filename, "Done"
                ExportPdfWord dict(filename), Utils.JoinPath(folderName, filename)
            Else
                Debug.Print filename, "No Action"
            End If
            
            dict.Remove filename
        Else
            Debug.Print filename, "No action"
        End If
Continue:
        filename = dir
    Loop
    
    ' �������t�@�C���̃`�F�b�N
    Debug.Print "*** �������t�@�C�� ***"
    Dim k
    For Each k In dict
        Debug.Print k
    Next k
    
End Sub

'
' 1�s (cols) �� startColumn ���ʒu�ȍ~�̃Z���̒l��z��ɂ��ĕԂ�
' ��̒l�̓X�L�b�v����
'
Function GetTargetSheetNames(cols As Range, startColumn As Integer) As Variant
    ' TODO: �ΏۃV�[�g�I��(for Excel)
    Dim i As Integer
    Dim j As Integer
    Dim ary()
    
    j = 0
    For i = startColumn To cols.Count
        ReDim Preserve ary(j)
        If cols(i) <> "" Then
            ary(j) = cols(i)
            j = j + 1
        End If
    Next i

    GetTargetSheetNames = ary

End Function

' �ꗗ�V�[�g����1�s��pdfInfo�̃f�[�^
'
' srcFilename, pdfFilename, sheetName,...
'
' ���擾���Adict�ɃZ�b�g����
Sub �ꗗ�擾(dict As Scripting.Dictionary, xlsPath As String)
    Dim book As Workbook
    Set book = Workbooks.Open(xlsPath, ReadOnly:=True)
    
    Dim sheet As Worksheet
    Set sheet = book.Sheets("�ꗗ")
    
    Dim r As Range
    
    For Each r In rngExpand(sheet.UsedRange.Rows, xlUp, -1) ' �P�s��(�w�b�_)���΂���2�s�ڂ���
        If r.Columns(1) = "" Then GoTo Continue
        
        Dim target As New PdfInfo
        
        target.SrcFilename = r.Columns(1)
        target.PdfFilename = r.Columns(2)

        target.SheetNames = GetTargetSheetNames(r.Columns, 3)

        If dict.Exists(target.SrcFilename) Then
            Err.Raise Utils.UserErrorNumber, description:="�L�[:[" & target.SrcFilename & "]���d�����Ă��܂�"
        End If
        dict.Add target.SrcFilename, target

Continue:
        Set r = r.Offset(1, 0)
    Next r
    
    book.Saved = True
    book.Close
End Sub

Sub PDF����()
    Dim dict As New Scripting.Dictionary
    
    �ꗗ�擾 dict, Utils.JoinPath(ThisWorkbook.Path, "�ꗗ.xlsx")
    EachFile dict, Utils.JoinPath(ThisWorkbook.Path, "*")
End Sub

Sub �{�^��1_Click()
    PDF����
End Sub
