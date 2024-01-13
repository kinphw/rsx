Attribute VB_Name = "bCommon"
'namespace=vba-files\RS\common
Option Explicit

Public Function IsExistSheet(sheetName As String) As Boolean
    Dim ws As Worksheet
    For Each ws In Sheets
        If ws.Name = sheetName Then
            IsExistSheet = True
            Exit Function
        End If
    Next
    IsExistSheet = False
End Function

Public Sub CallSet(control As IRibbonControl)

    '¼³Á¤Æû È£Ãâ
    frSet.Show

End Sub
