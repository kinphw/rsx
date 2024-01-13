Attribute VB_Name = "aMain"
'namespace=vba-files\RS
Option Explicit
''''''''''''''''''''''''''''''''
' 책무구조도 규정추출 자동화
' v0.0.1 DD 240111
' by PHW
''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''

'워드객체 전역변수 관리
Public objWord As Object
Public objDoc As Object
Public gstrName As String

Public gfUseDept As Boolean '부서명을 적용할것인지
Public gfUseName As Boolean '내규명을 적용할것인지
Public gfUseRev As Boolean '제개정일자를 적용할것인지
'조문번호는 있든 없든 구현되도록 구현하면 됨

Public gEnd As Boolean '강제종료

''''''''''''''''''''''''''''''''

Sub RSMain(control As IRibbonControl)

'파일을 열고 워드객체를 전역변수에 설정. 가장 중요
If Not bOpen.OpenDoc() Then: MsgBox ("선택하지 않았습니다"): Exit Sub:

'메인 핸들러 호출
Call bRS.RSWrapper
If gEnd = True Then: Exit Sub

'종료 후 워드문서를 닫는다
Call bOpen.CloseDoc

'후처리
Call bPostp.Postp

'새로 저장
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
    '전역변수 초기화
    gfUseDept = True
    gfUseName = True
    gfUseRev = True
    gEnd = False
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''    