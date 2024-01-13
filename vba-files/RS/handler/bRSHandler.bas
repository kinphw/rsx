Attribute VB_Name = "bRSHandler"
'namespace=vba-files\RS\handler
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''
'��⺯��

Dim strLv1 As String '��
Dim strLv2 As String '��

Dim strRegexLv1 As String
Dim strRegexLv2 As String
Dim strRegexRev As String

Dim strDept As String '�μ��� : ������� ����
Dim strName As String '������ : �� ó��
Dim strRev As String '��1�� �ٷ� ����

Dim idStart As Integer
'''''''''''''''''''''''''''''''''''''''''''''
'Output ó����

Sub SetVar(ArrList As Object)

'�� ������ ArrList���� ��� ���� ��ü�� �������
Dim i As Integer
i = 1

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'1. �⺻������

'240113 : Ȥ�� �𸣴� Lv1/Lv2 Constant => ����ȭ
'a. ���Խ� ����

strLv1 = "��"
strLv2 = "��"

strRegexLv1 = "^��\s?\d*\s?" + strLv1 '^��\s?\d*\s?��
strRegexLv2 = "^��\s?\d*\s?" + strLv2 '^��\s?\d*\s?��
strRegexRev = "^\d{1,4}\.\d{1,2}\.\d{1,2}\.?\s[����]��" '2023.01.13 ���� / 2023.01.13 ����

'''''''''''''''''''''''''''''''''''''''''''''''''''''''

'b. ���빮�ڿ� ����

' a> �μ��� ����
If gfUseDept Then
    Dim strTmp As String
    strTmp = GetHeader()
    strTmp = Replace(strTmp, vbCr, "") '�������� Carrige Return�� �־ �׿��� ��
    strTmp = Replace(strTmp, " ", "") 'Debug : ����
    strDept = Header2Dept(strTmp)
Else
    strDept = ""
End If

' b> ������ ���� : ���� ù ��
If gfUseName Then
    strName = ArrList.Item(0)
Else
    strName = ""
End If

'''''''''''''''''''''''''''''''''''''
' c> �������� ����

Dim ArrListRev As Object
Set ArrListRev = CreateObject("System.Collections.ArrayList")

Dim row As Variant '���ν��� ����

If gfUseRev Then
    For Each row In ArrList
        'If TestRegex(row, "^��\s?[1]\s?��") Then
        'If TestRegex(row, strRegexLv1) Then
        If TestRegex(row, strRegexRev) Then
            'strRev = ArrList.Item((ArrList.IndexOf(row, 0)) - 1)
            strRev = row '0.0.2 ���� ���Խ����� �ν�
            Call ArrListRev.Add(row)
            'Exit For 'ã������ ����
        End If
    Next row

    If ArrListRev.Count = 0 Then '���� ã���� ������
        MsgBox "���������ڸ� �ν����� ���߽��ϴ�."
        strRev = ""
    Else
        strRev = ArrListRev.Item(ArrListRev.Count - 1) '������ ������
    End If

Else
    strRev = ""
End If

'''''''''''''''''''''''''''''''''''''
' d> ��ȯ ������ �ν�
Dim bExistLv1Lv2 As Boolean
bExistLv1Lv2 = False

For Each row In ArrList
    If TestRegex(row, strRegexLv1) Then
        idStart = ArrList.IndexOf(row,0)
        bExistLv1Lv2 = True
        'MsgBox "recognize ��"
        Exit For
    ElseIf TestRegex(row, strRegexLv2) Then
        idStart = ArrList.IndexOf(row,0)
        bExistLv1Lv2 = True
        'MsgBox "recognize ��"        
        Exit For
    Else
        Debug.Print "No Lv1 or Lv2"
    End If
Next row

If bExistLv1Lv2 = False Then
    MsgBox "�����ν� ��� Lv1(��), Lv2(��) ��� �νĵ��� �ʾҽ��ϴ�. �˰��� �������� �ν��� �� �����ϴ�."
    gEnd = True
End If

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub RunLoop(ArrList As Object)
''''''''''''''''''''''''''''''''''''''''''''''''''''''
'���۾�
''''''''''''''''''''''''''''''''''''''''''''''''''''''
ThisWorkbook.Activate

Dim ws As Worksheet
Set ws = Worksheets.Add '��Ʈ�� �߰��ϰ� �ű⼭ �۾��Ѵ�

ws.Activate
''''''''''''''''''''''''''''''''''''''''''''''''''''''
'���ڿ� ��ȯ��
''''''''''''''''''''''''''''''''''''''''''''''''''''''

'��ȯ���鼭 ���Խİ˻��ϸ� �����Ѵ�. (3,4��)
Dim row As Variant
'Dim FlagStart As Boolean
'FlagStart = False

'Dim bNowLv1 As Boolean '������ '��'�̾������� ���� �˰����� �޶���
'bNowLv1 = False '���ʿ��� bNowLv1 = False
'Dim strNowLv1 As String

'["First", "Lv1", "Lv2", "Other"] �� �ϳ�. ������ �پ��ϴ� �Ҹ��� �����ϰ� ���ڿ�ó��
Dim flagNow As Integer
Dim strNowLv1 As String 'Lv2���� Lv1 ���� ��������

Const FIRST = 0
Const LV1 = 1
Const LV2 = 2
Const OTHER = 3

flagNow = FIRST

Dim r As Integer '��
Dim c As Integer '��

r = 2
c = 4

For Each row In ArrList

    Debug.Print row

    If ArrList.IndexOf(row,0) < idStart Then
        Debug.Print "������ ����"
    ' Else
    
    '�ϴ� �� ó������ ��ȯ ���鼭 ���� ã�´�
    ' If Not FlagStart Then
    '     If TestRegex(row, strRegexLv1) Then
    '         FlagStart = True '�忡 �ɸ���, �÷���ó���ϰ� (2,4)�� ��´�
    '         Cells(2, 4) = row
    '         r = 2
    '         c = 4
    '         FlagNowJang = True
    '         strNowJang = row
            
    '     End If:
    
    '�÷��װ� �ɸ��� ���� ������� �ɸ� ����
    Else
        'Log1 : ��Ģ�̸� �׳� �����Ŵ
        If TestRegex(row, "^��\s*Ģ") Then
            Debug.Print "��Ģ ����"
            Exit For '�����Ŵ. ��Ģ ���� ���� �ʴ´�
        
        'Log2 : Lv1(��)�� ���
        ElseIf TestRegex(row, strRegexLv1) Then
            'Debug.Print "��. 1���� ��ġ"
            '���ó�� : 
            If flagNow = FIRST Then
                r = r 'ó���̸� �״�� : r = 2
            Else
                r = r + 1 '��/��/��Ÿ�� r +=1
            End If           
            c = 4
            Cells(r, c) = row            
            flagNow = LV1 'FlagNowJang = True
            strNowLv1 = row
        
        'Log3 : Lv2(��)�� ���
        ElseIf TestRegex(row, strRegexLv2) Then
            
            '���ó�� :
            If (flagNow = FIRST) Or (flagNow = LV1) Then 'ó���̰ų� ���̾�����, �ø��� �ʴ´�.
                r = r
            ElseIf (flagNow = LV2) Or (flagNow = OTHER) Then '������ �����ų� ��Ÿ������ r+=1
                r = r + 1 
            End If           
            
            ' If FlagNowJang Then
            '     r = r
            ' Else
            '     r = r + 1 '������ ���� �ƴ϶� ���� �ٸ����̾����� ���� ������ �Ѿ�� ��
            ' End If            
            'Debug.Print "��. 2���� ��ġ"
            c = 5
            Cells(r, c) = row
            Cells(r, 4) = strNowLv1

            flagNow = LV2 'FlagNowJang = False
            
        'Log4 : �̿ܿ��� 2���� �߰� ; r = r + 0
        Else
            '��Ÿ�� ��쿡�� �׳� 2���� �߰��ؼ� ��´�.
            c = 5
            Cells(r, c) = Cells(r, c).Value + vbCrLf + row
            Cells(r, 4) = strNowLv1

            flagNow = OTHER 'FlagNowJang = False
        End If
        
    End If
    
Next row

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''    

Sub SetHeader()
'3> �ֿ����� ä���
'Debug.Print "�߿����� ä���"

'1> ��� ���
Cells(1, 1) = "�Ұ��μ�"
Cells(1, 2) = "���Ը�"
Cells(1, 3) = "����������"
Cells(1, 4) = "������ȣ"
Cells(1, 5) = "��������"

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
