Attribute VB_Name = "bRS"
'namespace=vba-files\RS\handler
Option Explicit
'워드에서 설계한 코드를 엑셀에 붙인 것임

'''''''''''''''''''
'Main Procedure
'''''''''''''''''''
Sub RSWrapper() 'R책무S구조도 Main Sub

'워드객체를 읽는다 : 메인모듈 단위에서 별도 실시

'모든 행을 배열로 읽는다
Dim arrString() As String
arrString = bRSOther.CreateArray()

'ArrayList로 만들고 전처리(빈 행 삭제)
Dim ArrList As Object
Set ArrList = PrepArray(arrString)

'본작업 개시
'Call bRSHandler.Handler(ArrList)
Call bRSHandler.SetVar(ArrList)
if gEnd Then: Exit Sub
Call bRSHandler.RunLoop(ArrList)
Call bRSHandler.SetHeader()

End Sub
