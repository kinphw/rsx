Attribute VB_Name = "bRS"
'namespace=vba-files\RS\handler
'���忡�� ������ �ڵ带 ������ ���� ����

'''''''''''''''''''
'Main Procedure
'''''''''''''''''''
Sub RSWrapper() 'Rå��S������ Main Sub

'���尴ü�� �д´� : ���θ�� �������� ���� �ǽ�


'��� ���� �迭�� �д´�
Dim arrString() As String
arrString = CreateArray()


'ArrayList�� ����� ��ó��(�� �� ����)
Dim ArrList As Object
Set ArrList = PrepArray(arrString)


'���۾� ����
Call Handler(ArrList)


End Sub
