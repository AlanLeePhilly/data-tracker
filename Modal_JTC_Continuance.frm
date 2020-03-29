VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Modal_JTC_Continuance 
   Caption         =   "JTC Continuance"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6015
   OleObjectBlob   =   "Modal_JTC_Continuance.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Modal_JTC_Continuance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Continuance_Type_Change()
    If Not InStr(Continuance_Type.value, "Commonwealth") = 0 Then
        Reasons_Label.Enabled = True
        Reason1.Enabled = True
        Reason2.Enabled = True
        Reason3.Enabled = True
    Else
        Reasons_Label.Enabled = False
        Reason1.Enabled = False
        Reason1.value = "N/A"
        Reason2.Enabled = False
        Reason2.value = "N/A"
        Reason3.Enabled = False
        Reason3.value = "N/A"
    End If
End Sub

Private Sub Continue_Click()
    ClientUpdateForm.JTC_Continuance_Update.BackColor = selectedColor
    ClientUpdateForm.JTC_Continuance_Remain.BackColor = unselectedColor
    ClientUpdateForm.JTC_Return_Continuance.Caption = "Yes"

    Modal_JTC_Continuance.Hide
End Sub

Private Sub Cancel_Click()
    Unload Modal_JTC_Continuance
End Sub

