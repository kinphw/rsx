Attribute VB_Name = "aMain"
'namespace=vba-files\RS
''''''''''''''''''''''''''''''''
' 책무구조도 규정추출 자동화
' v0.0.1 DD 240111
' by PHW
''''''''''''''''''''''''''''''''

'워드객체 전역변수 관리
Public objWord As Object
Public objDoc As Object
Public gstrName As String

Sub RSMain(control As IRibbonControl)

'파일을 열고 워드객체를 전역변수에 설정. 가장 중요
If Not OpenDoc() Then: MsgBox ("선택하지 않았습니다"): Exit Sub:

'그냥 테스트
Debug.Print GetHeader

'메인 핸들러 호출
Call RSWrapper

'종료 후 워드문서를 닫는다
Call CloseDoc

'후처리
Call Postp

'새로 저장
If IsExistSheet(gstrName) Then: ActiveSheet.Name = gstrName + "_1": Else: ActiveSheet.Name = gstrName:

Dim strFileName As String
strFileName = gstrName + ".xlsx"

ActiveSheet.Copy
With ActiveSheet
    ActiveWorkbook.SaveAs Filename:=strFileName
End With

End Sub
