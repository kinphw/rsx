Attribute VB_Name = "cTest"
'namespace=vba-files\

Sub Test()

    Dim ArrList As Object
    Set ArrList = CreateObject("System.Collections.ArrayList")

    ArrList.Add "1"
    ArrList.Add "2"

    Debug.Print ArrList.Item(ArrList.Count - 1)

End Sub
