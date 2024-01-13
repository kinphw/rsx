Attribute VB_Name = "bRSOther"
'namespace=vba-files\RS\handler
Function CreateArray() '파일객체를 전역변수로부터 받아서 실행
' 모든 행을 arrString으로 추출

    Dim strAll As String
    Dim arrString() As String
    objWord.Selection.WholeStory '전역변수에서 ㅇ릭는다
    strAll = objWord.Selection.Range.Text
    arrString = Strings.Split(strAll, vbCr)
    
    CreateArray = arrString
    
End Function

Function PrepArray(arrString() As String) '파일객체를 받아서 실행

    'Array to ArrayList
    Dim ArrList As Object
    Set ArrList = CreateObject("System.Collections.ArrayList")
    For Each i In arrString: ArrList.Add (i): Next i 'Array to ArrayList
    
    '출력 : FOR TEST
    For Each i In ArrList: Debug.Print (i): Next i
    
    Dim arrList2 As Object
    Set arrList2 = CreateObject("System.Collections.ArrayList")
        
    '공백이 아닌 경우만 다시 새로운 ArrayList에 넣는다
    For Each i In ArrList:
        If Not i = "" Then: arrList2.Add (i)
    Next i
    
    Set PrepArray = arrList2
    
End Function


' 정규식 Test (Function arg으로 구현)
Function TestRegex(strTest As Variant, strReg As String) ' Str : 테스트대상, 'Regex : 정규식문자열

    'Dim Str As String
    Dim Regex As Object
    Set Regex = New RegExp
    
    'Regex.Pattern = "^제\s?\d*\s?조"
    'Str = "제 1장 제 1조"
    Regex.Pattern = strReg
    
    TestRegex = Regex.Test(strTest)

End Function


'워드객체 설정 후 헤더를 추출
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


'헤더를 정규식으로 해당 부분 추출해서 부서명을 만든다
Function Header2Dept(strFull As String)

Dim stringOne As String
Dim regexOne As Object
Dim theMatches As Object
Dim Match As Object
Set regexOne = New RegExp

'regexOne.Pattern = "[:]\s.*[)]$" '부서명 추출 정규식
regexOne.Pattern = "[:].*[)]$" '부서명 추출 정규식

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
