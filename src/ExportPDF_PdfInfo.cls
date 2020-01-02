VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PdfInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public SrcFilename As String
Public PdfFilename As String
Private mSheetNames()

Sub Class_Initialize()
End Sub

Public Property Let SheetNames(names As Variant)
    mSheetNames = names
End Property

Public Property Get SheetNames() As Variant
    SheetNames = mSheetNames
End Property

