Attribute VB_Name = "bRSOther"
'namespace=vba-files\RS\handler
Function CreateArray() '���ϰ�ü�� ���������κ��� �޾Ƽ� ����
' ��� ���� arrString���� ����

    Dim strAll As String
    Dim arrString() As String
    objWord.Selection.WholeStory '������������ �����´�
    strAll = objWord.Selection.Range.Text
    arrString = Strings.Split(strAll, vbCr)
    
    CreateArray = arrString
    
End Function

Function PrepArray(arrString() As String) '���ϰ�ü�� �޾Ƽ� ����

    'Array to ArrayList
    Dim ArrList As Object
    Set ArrList = CreateObject("System.Collections.ArrayList")
    For Each i In arrString: ArrList.Add (i): Next i 'Array to ArrayList
    
    '��� : FOR TEST
    For Each i In ArrList: Debug.Print (i): Next i
    
    Dim arrList2 As Object
    Set arrList2 = CreateObject("System.Collections.ArrayList")
        
    '������ �ƴ� ��츸 �ٽ� ���ο� ArrayList�� �ִ´�
    For Each i In ArrList:
        If Not i = "" Then: arrList2.Add (i)
    Next i
    
    Set PrepArray = arrList2
    
End Function


' ���Խ� Test (Function arg���� ����)
Function TestRegex(strTest As Variant, strReg As String) ' Str : �׽�Ʈ���, 'Regex : ���ԽĹ��ڿ�

    'Dim Str As String
    Dim Regex As Object
    Set Regex = New RegExp
    
    'Regex.Pattern = "^��\s?\d*\s?��"
    'Str = "�� 1�� �� 1��"
    Regex.Pattern = strReg
    
    TestRegex = Regex.Test(strTest)

End Function


'���尴ü ���� �� ����� ����
Function GetHeader()

'Dim oSection As Section
'Dim oHeader As HeaderFooter
Dim oSection As Object
Dim oHeader As Object

    For Each oSection In objDoc.Sections
        For Each oHeader In oSection.Headers
            If oHeader.Exists Then
                GetHeader = oHeader.Range.Text
            End If
        Next oHeader
    Next oSection
    
End Function


'����� ���Խ����� �ش� �κ� �����ؼ� �μ����� �����
Function Header2Dept(strFull As String)

Dim stringOne As String
Dim regexOne As Object
Dim theMatches As Object
Dim Match As Object
Set regexOne = New RegExp

'regexOne.Pattern = "[:]\s.*[)]$" '�μ��� ���� ���Խ�
regexOne.Pattern = "[:].*[)]$" '�μ��� ���� ���Խ�

regexOne.Global = False
regexOne.IgnoreCase = True

Set theMatches = regexOne.Execute(strFull)

Dim res As String

For Each Match In theMatches
  res = Match.Value
Next

res = Replace(res, ":", "")
res = Replace(res, ")", "")
res = Replace(res, " ", "")

Header2Dept = res

End Function
