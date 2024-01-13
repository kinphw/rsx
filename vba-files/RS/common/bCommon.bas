Attribute VB_Name = "bCommon"
'namespace=vba-files\RS\common
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
