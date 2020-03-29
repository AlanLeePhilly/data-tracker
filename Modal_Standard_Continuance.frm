VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Modal_Standard_Continuance 
   Caption         =   "UserForm1"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5520
   OleObjectBlob   =   "Modal_Standard_Continuance.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Modal_Standard_Continuance"
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
    ClientUpdateForm.Standard_Continuance_Update.BackColor = selectedColor
    ClientUpdateForm.Standard_Continuance_Remain.BackColor = unselectedColor
    ClientUpdateForm.Standard_Return_Continuance.Caption = "Yes"

    Modal_Standard_Continuance.Hide
End Sub

Private Sub Cancel_Click()
    Unload Modal_Standard_Continuance
End Sub

