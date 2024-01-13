Attribute VB_Name = "aHelp"
'namespace=vba-files\RS

Sub RSHelp(control As IRibbonControl)

     'Debug.Print "1"
     
     Dim msg As String
     
     msg = "책무구조도 규정인식 자동화 v0.0.2 by PHW" + vbCrLf + vbCrLf
     msg = msg + "대상 : K사 Word 서식 내규"
     
     Call MsgBox(msg, , "책무구조도")

End Sub
