Attribute VB_Name = "bOpen"
'namespace=vba-files\RS\open
Function OpenDoc() '워드파일을 선택한다
    
'Dim objWord As Object '전역변수
'Dim objDoc As Object '전역변수
Dim wdFileName

Set objWord = CreateObject("word.Application")
wdFileName = Application.GetOpenFilename("Word Documents, *.doc*")

'If wdFileName = False Then Exit Function
If wdFileName = False Then
    OpenDoc = False: Exit Function
End If

Set objDoc = GetObject(wdFileName)
Call objWord.Documents.Open(wdFileName, ReadOnly:=True)   '열어야 객체에 접근 가능

OpenDoc = True

End Function

Sub CloseDoc()

Const wdDoNotSaveChanges As Long = 0
objDoc.Close SaveChanges:=wdDoNotSaveChanges
objWord.Quit

End Sub
