Attribute VB_Name = "bRS"
'namespace=vba-files\RS\handler
Option Explicit
'���忡�� ������ �ڵ带 ������ ���� ����

'''''''''''''''''''
'Main Procedure
'''''''''''''''''''
Sub RSWrapper() 'Rå��S������ Main Sub

'���尴ü�� �д´� : ���θ�� �������� ���� �ǽ�

'��� ���� �迭�� �д´�
Dim arrString() As String
arrString = bRSOther.CreateArray()

'ArrayList�� ����� ��ó��(�� �� ����)
Dim ArrList As Object
Set ArrList = PrepArray(arrString)

'���۾� ����
'Call bRSHandler.Handler(ArrList)
Call bRSHandler.SetVar(ArrList)
if gEnd Then: Exit Sub
Call bRSHandler.RunLoop(ArrList)
Call bRSHandler.SetHeader()

End Sub
