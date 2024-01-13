Attribute VB_Name = "bRSHandler"
'namespace=vba-files\RS\handler
'Output ó����

Sub Handler(ArrList As Object)

'�� ������ ArrList���� ��� ���� ��ü�� �������
Dim i As Integer
i = 1

'���빮�ڿ� ����

Dim strDept As String '�μ��� : ������� ����
Dim strName As String '������ : �� ó��
Dim strRev As String '��1�� �ٷ� ����

'�μ��� ����
Dim strTmp As String
strTmp = GetHeader()
strTmp = Replace(strTmp, vbCr, "") '�������� Carrige Return�� �־ �׿��� ��
strTmp = Replace(strTmp, " ", "") 'Debug : ����
strDept = Header2Dept(strTmp)

'������ ����
strName = ArrList.Item(0)

'�������� ����
Dim row As Variant
For Each row In ArrList
    If TestRegex(row, "^��\s?[1]\s?��") Then
        strRev = ArrList.Item((ArrList.IndexOf(row, 0)) - 1)
        Exit For 'ã������ ����
    End If
Next row

''''''''''''''''''''''''''''''''''''''''''''''''''''''
'���۾�
''''''''''''''''''''''''''''''''''''''''''''''''''''''
ThisWorkbook.Activate

Dim ws As Worksheet
Set ws = Worksheets.Add '��Ʈ�� �߰��ϰ� �ű⼭ �۾��Ѵ�

ws.Activate

Dim r As Integer '��
Dim c As Integer '��

'1> ��� ���
Cells(1, 1) = "�Ұ��μ�"
Cells(1, 2) = "���Ը�"
Cells(1, 3) = "����������"
Cells(1, 4) = "������ȣ"
Cells(1, 5) = "��������"

'2> ��ȯ���鼭 ���Խİ˻��ϸ� �����Ѵ�. (3,4��)
'Dim row As Variant
Dim FlagStart As Boolean
FlagStart = False

Dim FlagNowJang As Boolean '������ '��'�̾������� ���� ���� �˰����� �޶���
FlagNowJang = False

Dim strNowJang As String

For Each row In ArrList

    Debug.Print row
    
    '�ϴ� �� ó������ ��ȯ ���鼭 ���� ã�´�
    If Not FlagStart Then
        If TestRegex(row, "^��\s?\d*\s?��") Then
            FlagStart = True '�忡 �ɸ���, �÷���ó���ϰ� (2,4)�� ��´�
            Cells(2, 4) = row
            r = 2
            c = 4
            FlagNowJang = True
            strNowJang = row
            
        End If:
    
    '�÷��װ� �ɸ��� ���� ������� �ɸ� ����
    Else
        'Log1 : ��Ģ�̸� �׳� �����Ŵ
        If TestRegex(row, "^��\s*Ģ") Then
            Debug.Print "��Ģ ����"
            Exit For '�����Ŵ. ��Ģ ���� ���� �ʴ´�
        
        'Log2 : ���̸� 1���� ��ġ ; r = r + 0
        ElseIf TestRegex(row, "^��\s?\d*\s?��") Then
            Debug.Print "��. 1���� ��ġ"
            '�ϴ� ���� �� 4���� ��´�
            r = r + 1
            c = 4
            Cells(r, c) = row
            FlagNowJang = True
            strNowJang = row
        
        'Log3 : ���� 2 ���� ��ġ ; r = r + 1
        ElseIf TestRegex(row, "^��\s?\d*\s?��") Then
            
            If FlagNowJang Then
                r = r
            Else
                r = r + 1 '������ ���� �ƴ϶� ���� �ٸ����̾����� ���� ������ �Ѿ�� ��
            End If
            
            Debug.Print "��. 2���� ��ġ"
            c = 5
            Cells(r, c) = row
            Cells(r, 4) = strNowJang
            
            FlagNowJang = False
            
        'Log4 : �̿ܿ��� 2���� �߰� ; r = r + 0
        Else
            Debug.Print "��Ÿ. 2���� �߰���ġ"
            Cells(r, c) = Cells(r, c).Value + vbCrLf + row
            Cells(r, 4) = strNowJang
            FlagNowJang = False
        End If
        
    End If
    
Next row


'3> �ֿ����� ä���
Debug.Print "�߿����� ä���"

Dim countRow As Integer
countRow = ActiveSheet.UsedRange.Rows.Count

Debug.Print strDept
Range("A2", Cells(countRow, 1)) = strDept

Debug.Print strName
Range("B2", Cells(countRow, 2)) = strName
gstrName = strName

Debug.Print strRev
Range("C2", Cells(countRow, 3)) = strRev

End Sub
