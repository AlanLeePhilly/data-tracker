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


    'note: this modal serves JTC, PJJSC, and standard updaters, so it colors buttons on all forms no matter who calls them
    ClientUpdateForm.JTC_FTA_No.BackColor = unselectedColor
    ClientUpdateForm.JTC_FTA_Yes.BackColor = selectedColor
    ClientUpdateForm.Standard_FTA_No.BackColor = unselectedColor
    ClientUpdateForm.Standard_FTA_Yes.BackColor = selectedColor
    
    If BW.value = "Yes" Then
        ClientUpdateForm.PJJSC_DA_Action.value = "N/A"
        ClientUpdateForm.PJJSC_DA_Action.Enabled = False
        
        ClientUpdateForm.PJJSC_ActionAccepted.value = "N/A"
        ClientUpdateForm.PJJSC_ActionAccepted.Enabled = False
        
        ClientUpdateForm.PJJSC_Facility.value = "N/A"
        ClientUpdateForm.PJJSC_Facility.Enabled = False
        
        ClientUpdateForm.PJJSC_NextLocation.value = "N/A"
        ClientUpdateForm.PJJSC_NextLocation.Enabled = False
    Else
        ClientUpdateForm.PJJSC_DA_Action.Enabled = True
        ClientUpdateForm.PJJSC_ActionAccepted.Enabled = True
        ClientUpdateForm.PJJSC_Facility.Enabled = True
        ClientUpdateForm.PJJSC_NextLocation.Enabled = True
    End If

    Me.Hide
End Sub
