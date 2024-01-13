VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frSet 
   Caption         =   "기본설정"
   ClientHeight    =   2565
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4485
   OleObjectBlob   =   "frSet.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub UserForm_Initialize()

If gfUseDept = True Then: cb1.Value = True: Else: cb1.Value = False
If gfUseName = True Then: cb2.Value = True: Else: cb2.Value = False
If gfUseRev = True Then: cb3.Value = True: Else: cb3.Value = False

End Sub

Private Sub cb1_Change()

If cb1.Value = True Then: gfUseDept = True: Else: gfUseDept = False
    
End Sub


Private Sub cb2_Change()

If cb2.Value = True Then: gfUseName = True: Else: gfUseName = False

End Sub


Private Sub cb3_Change()

If cb3.Value = True Then: gfUseRev = True: Else: gfUseRev = False
    
End Sub
