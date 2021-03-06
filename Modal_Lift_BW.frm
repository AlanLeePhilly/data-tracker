VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Modal_Lift_BW 
   Caption         =   "UserForm1"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5745
   OleObjectBlob   =   "Modal_Lift_BW.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Modal_Lift_BW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub DateBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Set ctl = Me.DateBox
    'send to date validation
    Call DateValidation(ctl, Cancel)
End Sub

Private Sub DateBox_Enter()
    DateBox.value = CalendarForm.GetDate(RangeOfYears:=5)
End Sub

Private Sub InsertDoH_Click()
    DateBox.value = ClientUpdateForm.DateOfHearing.value
End Sub

Private Sub Submit_Click()
    If DateBox.value = "" Then
        MsgBox "Date required"
        Exit Sub
    End If

    ClientUpdateForm.JTC_Lift_BW.BackColor = selectedColor
    ClientUpdateForm.Standard_Lift_BW.BackColor = selectedColor
    ClientUpdateForm.PJJSC_Lift_BW.BackColor = selectedColor
    
    With ClientUpdateForm
        Call enableFrame(.PJJSC_Reasons_Frame)
        Call enableFrame(.PJJSC_Outcome_Frame)
        Call enableFrame(.PJJSC_Sup_Frame)
        Call enableFrame(.PJJSC_Cond_Frame)
        .PJJSC_Cancel.Enabled = True
        .PJJSC_Submit.Enabled = True
    End With
    
    Me.Hide
End Sub
Private Sub Cancel_Click()
    ClientUpdateForm.JTC_Lift_BW.BackColor = unselectedColor
    ClientUpdateForm.Standard_Lift_BW.BackColor = unselectedColor
    ClientUpdateForm.PJJSC_Lift_BW.BackColor = unselectedColor
    
    Me.Hide
End Sub

