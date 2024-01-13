Attribute VB_Name = "bRSHandler"
'namespace=vba-files\RS\handler
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''
'모듈변수

Dim strLv1 As String '장
Dim strLv2 As String '조

Dim strRegexLv1 As String
Dim strRegexLv2 As String
Dim strRegexRev As String

Dim strDept As String '부서명 : 헤더에서 읽음
Dim strName As String '규정명 : 맨 처음
Dim strRev As String '제1장 바로 직전

Dim idStart As Integer
'''''''''''''''''''''''''''''''''''''''''''''
'Output 처리부

Sub SetVar(ArrList As Object)

'이 시점에 ArrList에는 대상 문자 전체가 담겨있음
Dim i As Integer
i = 1

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'1. 기본설정부

'240113 : 혹시 모르니 Lv1/Lv2 Constant => 변수화
'a. 정규식 설정

strLv1 = "장"
strLv2 = "조"

strRegexLv1 = "^제\s?\d*\s?" + strLv1 '^제\s?\d*\s?장
strRegexLv2 = "^제\s?\d*\s?" + strLv2 '^제\s?\d*\s?조
strRegexRev = "^\d{1,4}\.\d{1,2}\.\d{1,2}\.?\s[제개]정" '2023.01.13 제정 / 2023.01.13 개정

'''''''''''''''''''''''''''''''''''''''''''''''''''''''

'b. 공통문자열 설정

' a> 부서명 설정
If gfUseDept Then
    Dim strTmp As String
    strTmp = GetHeader()
    strTmp = Replace(strTmp, vbCr, "") '마지막에 Carrige Return이 있어서 죽여야 함
    strTmp = Replace(strTmp, " ", "") 'Debug : 공백
    strDept = Header2Dept(strTmp)
Else
    strDept = ""
End If

' b> 규정명 설정 : 가장 첫 행
If gfUseName Then
    strName = ArrList.Item(0)
Else
    strName = ""
End If

'''''''''''''''''''''''''''''''''''''
' c> 최종개정 설정

Dim ArrListRev As Object
Set ArrListRev = CreateObject("System.Collections.ArrayList")

Dim row As Variant '프로시져 변수

If gfUseRev Then
    For Each row In ArrList
        'If TestRegex(row, "^제\s?[1]\s?장") Then
        'If TestRegex(row, strRegexLv1) Then
        If TestRegex(row, strRegexRev) Then
            'strRev = ArrList.Item((ArrList.IndexOf(row, 0)) - 1)
            strRev = row '0.0.2 직접 정규식으로 인식
            Call ArrListRev.Add(row)
            'Exit For '찾았으면 종료
        End If
    Next row

    If ArrListRev.Count = 0 Then '만약 찾은게 없으면
        MsgBox "제개정일자를 인식하지 못했습니다."
        strRev = ""
    Else
        strRev = ArrListRev.Item(ArrListRev.Count - 1) '마지막 아이템
    End If

Else
    strRev = ""
End If

'''''''''''''''''''''''''''''''''''''
' d> 반환 시작점 인식
Dim bExistLv1Lv2 As Boolean
bExistLv1Lv2 = False

For Each row In ArrList
    If TestRegex(row, strRegexLv1) Then
        idStart = ArrList.IndexOf(row,0)
        bExistLv1Lv2 = True
        'MsgBox "recognize 장"
        Exit For
    ElseIf TestRegex(row, strRegexLv2) Then
        idStart = ArrList.IndexOf(row,0)
        bExistLv1Lv2 = True
        'MsgBox "recognize 조"        
        Exit For
    Else
        Debug.Print "No Lv1 or Lv2"
    End If
Next row

If bExistLv1Lv2 = False Then
    MsgBox "문자인식 결과 Lv1(장), Lv2(조) 모두 인식되지 않았습니다. 알고리즘 시작점을 인식할 수 없습니다."
    gEnd = True
End If

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub RunLoop(ArrList As Object)
''''''''''''''''''''''''''''''''''''''''''''''''''''''
'본작업
''''''''''''''''''''''''''''''''''''''''''''''''''''''
ThisWorkbook.Activate

Dim ws As Worksheet
Set ws = Worksheets.Add '시트를 추가하고 거기서 작업한다

ws.Activate
''''''''''''''''''''''''''''''''''''''''''''''''''''''
'문자열 순환부
''''''''''''''''''''''''''''''''''''''''''''''''''''''

'순환돌면서 정규식검사하며 진행한다. (3,4열)
Dim row As Variant
'Dim FlagStart As Boolean
'FlagStart = False

'Dim bNowLv1 As Boolean '직전에 '장'이었는지에 따라 알고리즘이 달라짐
'bNowLv1 = False '최초에는 bNowLv1 = False
'Dim strNowLv1 As String

'["First", "Lv1", "Lv2", "Other"] 중 하나. 변수가 다양하니 불린은 포기하고 문자열처리
Dim flagNow As Integer
Dim strNowLv1 As String 'Lv2에서 Lv1 같이 찍으려고

Const FIRST = 0
Const LV1 = 1
Const LV2 = 2
Const OTHER = 3

flagNow = FIRST

Dim r As Integer '행
Dim c As Integer '열

r = 2
c = 4

For Each row In ArrList

    Debug.Print row

    If ArrList.IndexOf(row,0) < idStart Then
        Debug.Print "시작점 이전"
    ' Else
    
    '일단 맨 처음에는 순환 돌면서 장을 찾는다
    ' If Not FlagStart Then
    '     If TestRegex(row, strRegexLv1) Then
    '         FlagStart = True '장에 걸리면, 플래그처리하고 (2,4)에 찍는다
    '         Cells(2, 4) = row
    '         r = 2
    '         c = 4
    '         FlagNowJang = True
    '         strNowJang = row
            
    '     End If:
    
    '플래그가 걸리면 이제 여기부터 걸릴 것임
    Else
        'Log1 : 부칙이면 그냥 종료시킴
        If TestRegex(row, "^부\s*칙") Then
            Debug.Print "부칙 등장"
            Exit For '종료시킴. 부칙 밑은 보지 않는다
        
        'Log2 : Lv1(장)인 경우
        ElseIf TestRegex(row, strRegexLv1) Then
            'Debug.Print "장. 1열에 배치"
            '행렬처리 : 
            If flagNow = FIRST Then
                r = r '처음이면 그대로 : r = 2
            Else
                r = r + 1 '장/조/기타면 r +=1
            End If           
            c = 4
            Cells(r, c) = row            
            flagNow = LV1 'FlagNowJang = True
            strNowLv1 = row
        
        'Log3 : Lv2(조)인 경우
        ElseIf TestRegex(row, strRegexLv2) Then
            
            '행렬처리 :
            If (flagNow = FIRST) Or (flagNow = LV1) Then '처음이거나 장이었으면, 늘리지 않는다.
                r = r
            ElseIf (flagNow = LV2) Or (flagNow = OTHER) Then '직전에 조였거나 기타였으면 r+=1
                r = r + 1 
            End If           
            
            ' If FlagNowJang Then
            '     r = r
            ' Else
            '     r = r + 1 '직전에 장이 아니라 조나 다른것이었으면 다음 행으로 넘어가야 함
            ' End If            
            'Debug.Print "조. 2열에 배치"
            c = 5
            Cells(r, c) = row
            Cells(r, 4) = strNowLv1

            flagNow = LV2 'FlagNowJang = False
            
        'Log4 : 이외에는 2열에 추가 ; r = r + 0
        Else
            '기타인 경우에는 그냥 2열에 추가해서 찍는다.
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
'3> 주요정보 채우기
'Debug.Print "중요정보 채우기"

'1> 헤더 출력
Cells(1, 1) = "소관부서"
Cells(1, 2) = "내규명"
Cells(1, 3) = "제개정일자"
Cells(1, 4) = "조문번호"
Cells(1, 5) = "조문내용"

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
