Attribute VB_Name = "bOpen"
'namespace=vba-files\RS\open
Function OpenDoc() '���������� �����Ѵ�
    
'Dim objWord As Object '��������
'Dim objDoc As Object '��������
Dim wdFileName

Set objWord = CreateObject("word.Application")
wdFileName = Application.GetOpenFilename("Word Documents, *.doc*")

'If wdFileName = False Then Exit Function
If wdFileName = False Then
    OpenDoc = False: Exit Function
End If

Set objDoc = GetObject(wdFileName)
Call objWord.Documents.Open(wdFileName, ReadOnly:=True)   '����� ��ü�� ���� ����

OpenDoc = True

End Function

Sub CloseDoc()

Const wdDoNotSaveChanges As Long = 0
objDoc.Close SaveChanges:=wdDoNotSaveChanges
objWord.Quit

End Sub
