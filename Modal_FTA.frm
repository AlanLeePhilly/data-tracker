VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Modal_FTA 
   ClientHeight    =   2550
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4485
   OleObjectBlob   =   "Modal_FTA.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Modal_FTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Submit_Click()
    If BW.value = "" Then
        MsgBox "Bench warrant action required"
        Exit Sub
    End If
    
    
    'note: this modal serves JTC and standard updaters, so it colors buttons on both forms no matter who calls them
    ClientUpdateForm.JTC_FTA_No.BackColor = unselectedColor
    ClientUpdateForm.JTC_FTA_Yes.BackColor = selectedColor
    ClientUpdateForm.Standard_FTA_No.BackColor = unselectedColor
    ClientUpdateForm.Standard_FTA_Yes.BackColor = selectedColor
    
    Me.Hide
End Sub
