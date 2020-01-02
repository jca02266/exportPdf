Attribute VB_Name = "Utils"
Option Explicit

Public Const UserErrorNumber As Long = vbObjectError + 513

Sub Assert(expect As Variant, actual As Variant, Optional description As String)
    If Not IsMissing(description) Then
        Debug.Print description,
    End If
    
    If expect = actual Then
        Debug.Print "OK"
    Else
        Debug.Print "NG " & expect & "<>" & actual
    End If
End Sub

Function JoinPath(ParamArray args()) As String
    If IsMissing(args) Then
        JoinPath = ""
        Exit Function
    End If

    JoinPath = args(LBound(args))

    Dim i As Integer

    For i = LBound(args) + 1 To UBound(args)
        Dim arg As String
        arg = args(i)

        If Len(JoinPath) > 0 And Right(JoinPath, 1) = "\" Then
            JoinPath = JoinPath & arg
        Else
            JoinPath = JoinPath & "\" & arg
        End If
    Next i
End Function

Private Sub test_JoinPath()
    Assert "", JoinPath(), "1"
    Assert "foo", JoinPath("foo"), "2"
    Assert "foo\bar", JoinPath("foo", "bar"), "3"
    Assert "foo\bar\baz", JoinPath("foo", "bar", "baz"), "4"
End Sub

Function Basename(ByVal Path As String) As String
    Dim pos As Integer
    
    If Right(Path, 1) = "\" Then
        Path = Mid(Path, 1, Len(Path) - 1)
    End If
    
    pos = InStrRev(Path, "\")
    If pos > 0 Then
        Basename = Mid(Path, pos + 1)
    Else
        Basename = Path
    End If
End Function

Private Sub test_Basename()
    Assert "", Basename(""), "1"
    Assert "foo", Basename("foo"), "2"
    Assert "bar", Basename("foo\bar"), "3"
    Assert "baz", Basename("foo\bar\baz"), "4"
    
    Assert "foo", Basename("foo\"), "5"

    ' 引数を変更しないこと
    Dim foo As String
    foo = "foo\"
    Assert "foo", Basename(foo), 6
    Assert "foo\", foo, 7
End Sub

Function Dirname(ByVal Path As String) As String
    Dim pos As Integer
    
    If Right(Path, 1) = "\" Then
        Path = Mid(Path, 1, Len(Path) - 1)
    End If
    
    pos = InStrRev(Path, "\")
    If pos > 0 Then
        Dirname = Mid(Path, 1, pos - 1)
    Else
        Dirname = "."
    End If
End Function

Private Sub test_Dirname()
    Assert ".", Dirname(""), "1"
    Assert ".", Dirname("foo"), "2"
    Assert "foo", Dirname("foo\bar"), "3"
    Assert "foo\bar", Dirname("foo\bar\baz"), "4"
    
    Assert ".", Dirname("foo\"), "5"

    ' 引数を変更しないこと
    Dim foo As String
    foo = "foo\"
    Assert ".", Dirname(foo), 6
    Assert "foo\", foo, 7
End Sub

Function rngExpand(r As Range, direction As XlDirection, num As Integer) As Range
    Select Case direction
    Case XlDirection.xlUp
        Set rngExpand = r.Offset(-num, 0).Resize(r.Rows.Count + num, r.Columns.Count)
    Case XlDirection.xlToLeft
        Set rngExpand = r.Offset(0, -num).Resize(r.Rows.Count, r.Columns.Count + num)
    Case XlDirection.xlToRight
        Set rngExpand = r.Resize(r.Rows.Count, r.Columns.Count + num)
    Case XlDirection.xlDown
        Set rngExpand = r.Resize(r.Rows.Count + num, r.Columns.Count)
    Case Default
        Err.Raise UserErrorNumber, description:="XlDirection の値を指定してください"
    End Select

End Function

Sub test_rngExpand()
    Dim r As Range
    
    Set r = ThisWorkbook.Sheets(1).Range("B2:C3")
    
    Set r = rngExpand(r, xlUp, 1)
    Assert "B1:C3", r.Address(False, False), "expand top"

    Set r = rngExpand(r, xlUp, -1)
    Assert "B2:C3", r.Address(False, False), "shrink top"

    Set r = rngExpand(r, xlToLeft, 1)
    Assert "A2:C3", r.Address(False, False), "expand left"

    Set r = rngExpand(r, xlToLeft, -1)
    Assert "B2:C3", r.Address(False, False), "shrink left"

    Set r = rngExpand(r, xlToRight, 1)
    Assert "B2:D3", r.Address(False, False), "expand left"

    Set r = rngExpand(r, xlToRight, -1)
    Assert "B2:C3", r.Address(False, False), "shrink left"

    Set r = rngExpand(r, xlDown, 1)
    Assert "B2:C4", r.Address(False, False), "expand bottom"

    Set r = rngExpand(r, xlDown, -1)
    Assert "B2:C3", r.Address(False, False), "shrink bottom"

End Sub

' 配列の値をキーにインデックスを値にした辞書を返す
Function aryToDictionary(ary As Variant) As Scripting.Dictionary
    Dim i As Integer
    Dim dict As New Scripting.Dictionary
    
    For i = LBound(ary) To UBound(ary)
        dict.Add ary(i), i
    Next i

    Set aryToDictionary = dict
End Function
