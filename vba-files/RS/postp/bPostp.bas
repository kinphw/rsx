Attribute VB_Name = "bPostp"
'namespace=vba-files\RS\postp

Sub Postp()

'ÈÄÃ³¸®

Columns("A:D").ColumnWidth = 20
Columns("E").ColumnWidth = 100

ActiveWindow.Zoom = 85
ActiveSheet.UsedRange.WrapText = True
Call drawLine

Range("A1:E1").HorizontalAlignment = xlCenter

End Sub

Sub drawLine()

Set Rng = ActiveSheet.UsedRange

Rng.Borders.LineStyle = 1
Rng.Borders.Weight = xlThin
Rng.Borders.ColorIndex = 1

End Sub
