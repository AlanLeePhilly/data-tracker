VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Modal_Standard_Court_Transfer 
   Caption         =   "Standard Court Discharge"
   ClientHeight    =   3765
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   8088
   OleObjectBlob   =   "Modal_Standard_Court_Transfer.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Modal_Standard_Court_Transfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancel_Click()
    ClientUpdateForm.Standard_Court_Transfer.BackColor = unselectedColor
    Modal_Standard_Court_Transfer.Hide
End Sub

Private Sub Continue_Click()
    If Detailed_Outcome.value = "N/A" Then
        MsgBox "Detailed Outcome Required"
        Exit Sub
    End If

    If Courtroom.value = "N/A" Then
        MsgBox "Courtroom Required"
        Exit Sub
    End If

    ClientUpdateForm.Standard_Court_Transfer.BackColor = selectedColor
    Modal_Standard_Court_Transfer.Hide
End Sub

