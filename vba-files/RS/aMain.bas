Attribute VB_Name = "aMain"
'namespace=vba-files\RS
''''''''''''''''''''''''''''''''
' å�������� �������� �ڵ�ȭ
' v0.0.1 DD 240111
' by PHW
''''''''''''''''''''''''''''''''

'���尴ü �������� ����
Public objWord As Object
Public objDoc As Object
Public gstrName As String

Sub RSMain(control As IRibbonControl)

'������ ���� ���尴ü�� ���������� ����. ���� �߿�
If Not OpenDoc() Then: MsgBox ("�������� �ʾҽ��ϴ�"): Exit Sub:

'�׳� �׽�Ʈ
Debug.Print GetHeader

'���� �ڵ鷯 ȣ��
Call RSWrapper

'���� �� ���幮���� �ݴ´�
Call CloseDoc

'��ó��
Call Postp

'���� ����
If IsExistSheet(gstrName) Then: ActiveSheet.Name = gstrName + "_1": Else: ActiveSheet.Name = gstrName:

Dim strFileName As String
strFileName = gstrName + ".xlsx"

ActiveSheet.Copy
With ActiveSheet
    ActiveWorkbook.SaveAs Filename:=strFileName
End With

End Sub
