Attribute VB_Name = "aHelp"
'namespace=vba-files\RS

Sub RSHelp(control As IRibbonControl)

     'Debug.Print "1"
     
     Dim msg As String
     
     msg = "å�������� �����ν� �ڵ�ȭ v0.0.2 by PHW" + vbCrLf + vbCrLf
     msg = msg + "��� : K�� Word ���� ����"
     
     Call MsgBox(msg, , "å��������")

End Sub
