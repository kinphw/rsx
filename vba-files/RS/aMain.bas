Attribute VB_Name = "aMain"
'namespace=vba-files\RS
Option Explicit
''''''''''''''''''''''''''''''''
' å�������� �������� �ڵ�ȭ
' v0.0.1 DD 240111
' by PHW
''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''

'���尴ü �������� ����
Public objWord As Object
Public objDoc As Object
Public gstrName As String

Public gfUseDept As Boolean '�μ����� �����Ұ�����
Public gfUseName As Boolean '���Ը��� �����Ұ�����
Public gfUseRev As Boolean '���������ڸ� �����Ұ�����
'������ȣ�� �ֵ� ���� �����ǵ��� �����ϸ� ��

Public gEnd As Boolean '��������

''''''''''''''''''''''''''''''''

Sub RSMain(control As IRibbonControl)

'������ ���� ���尴ü�� ���������� ����. ���� �߿�
If Not bOpen.OpenDoc() Then: MsgBox ("�������� �ʾҽ��ϴ�"): Exit Sub:

'���� �ڵ鷯 ȣ��
Call bRS.RSWrapper
If gEnd = True Then: Exit Sub

'���� �� ���幮���� �ݴ´�
Call bOpen.CloseDoc

'��ó��
Call bPostp.Postp

'���� ����
If gstrName = "" Then: gstrName = "0" 'Debug
If bCommon.IsExistSheet(gstrName) Then: ActiveSheet.Name = gstrName + "_1": Else: ActiveSheet.Name = gstrName:

Dim strFileName As String
strFileName = gstrName + ".xlsx"

ActiveSheet.Copy
With ActiveSheet
    ActiveWorkbook.SaveAs Filename:=strFileName
End With

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub Auto_Open()
    '�������� �ʱ�ȭ
    gfUseDept = True
    gfUseName = True
    gfUseRev = True
    gEnd = False
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''    