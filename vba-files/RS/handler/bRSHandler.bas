Attribute VB_Name = "bRSHandler"
'namespace=vba-files\RS\handler
'Output 처리부

Sub Handler(ArrList As Object)

'이 시점에 ArrList에는 대상 문자 전체가 담겨있음
Dim i As Integer
i = 1

'공통문자열 설정

Dim strDept As String '부서명 : 헤더에서 읽음
Dim strName As String '규정명 : 맨 처음
Dim strRev As String '제1장 바로 직전

'부서명 설정
Dim strTmp As String
strTmp = GetHeader()
strTmp = Replace(strTmp, vbCr, "") '마지막에 Carrige Return이 있어서 죽여야 함
strTmp = Replace(strTmp, " ", "") 'Debug : 공백
strDept = Header2Dept(strTmp)

'규정명 설정
strName = ArrList.Item(0)

'최종개정 설정
Dim row As Variant
For Each row In ArrList
    If TestRegex(row, "^제\s?[1]\s?장") Then
        strRev = ArrList.Item((ArrList.IndexOf(row, 0)) - 1)
        Exit For '찾았으면 종료
    End If
Next row

''''''''''''''''''''''''''''''''''''''''''''''''''''''
'본작업
''''''''''''''''''''''''''''''''''''''''''''''''''''''
ThisWorkbook.Activate

Dim ws As Worksheet
Set ws = Worksheets.Add '시트를 추가하고 거기서 작업한다

ws.Activate

Dim r As Integer '행
Dim c As Integer '열

'1> 헤더 출력
Cells(1, 1) = "소관부서"
Cells(1, 2) = "내규명"
Cells(1, 3) = "제개정일자"
Cells(1, 4) = "조문번호"
Cells(1, 5) = "조문내용"

'2> 순환돌면서 정규식검사하며 진행한다. (3,4열)
'Dim row As Variant
Dim FlagStart As Boolean
FlagStart = False

Dim FlagNowJang As Boolean '직전에 '장'이었는지에 따라 조의 알고리즘이 달라짐
FlagNowJang = False

Dim strNowJang As String

For Each row In ArrList

    Debug.Print row
    
    '일단 맨 처음에는 순환 돌면서 장을 찾는다
    If Not FlagStart Then
        If TestRegex(row, "^제\s?\d*\s?장") Then
            FlagStart = True '장에 걸리면, 플래그처리하고 (2,4)에 찍는다
            Cells(2, 4) = row
            r = 2
            c = 4
            FlagNowJang = True
            strNowJang = row
            
        End If:
    
    '플래그가 걸리면 이제 여기부터 걸릴 것임
    Else
        'Log1 : 부칙이면 그냥 종료시킴
        If TestRegex(row, "^부\s*칙") Then
            Debug.Print "부칙 등장"
            Exit For '종료시킴. 부칙 밑은 보지 않는다
        
        'Log2 : 장이면 1열에 배치 ; r = r + 0
        ElseIf TestRegex(row, "^제\s?\d*\s?장") Then
            Debug.Print "장. 1열에 배치"
            '일단 다음 행 4열에 찍는다
            r = r + 1
            c = 4
            Cells(r, c) = row
            FlagNowJang = True
            strNowJang = row
        
        'Log3 : 조면 2 열에 배치 ; r = r + 1
        ElseIf TestRegex(row, "^제\s?\d*\s?조") Then
            
            If FlagNowJang Then
                r = r
            Else
                r = r + 1 '직전에 장이 아니라 조나 다른것이었으면 다음 행으로 넘어가야 함
            End If
            
            Debug.Print "조. 2열에 배치"
            c = 5
            Cells(r, c) = row
            Cells(r, 4) = strNowJang
            
            FlagNowJang = False
            
        'Log4 : 이외에는 2열에 추가 ; r = r + 0
        Else
            Debug.Print "기타. 2열에 추가배치"
            Cells(r, c) = Cells(r, c).Value + vbCrLf + row
            Cells(r, 4) = strNowJang
            FlagNowJang = False
        End If
        
    End If
    
Next row


'3> 주요정보 채우기
Debug.Print "중요정보 채우기"

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
