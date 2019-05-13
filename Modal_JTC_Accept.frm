VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Modal_JTC_Accept 
   Caption         =   "JTC - Accept Client"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6000
   OleObjectBlob   =   "Modal_JTC_Accept.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Modal_JTC_Accept"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub DateOfAcceptance_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Set ctl = Modal_JTC_Accept.DateOfAcceptance
    'send to date validation
    Call DateValidation(ctl, Cancel)
End Sub

Private Sub StepupDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Set ctl = Me.StepupDate
    'send to date validation
    Call DateValidation(ctl, Cancel)
End Sub
Private Sub StepupDate_Enter()

    StepupDate.value = CalendarForm.GetDate(RangeOfYears:=5)
    
End Sub

Private Sub Cancel_Click()
    Unload Me
End Sub

Private Sub InsertDoH_Click()
    DateOfAcceptance.value = ClientUpdateForm.DateOfHearing.value
End Sub

Private Sub Continue_Click()

    If DateOfAcceptance.value = "" Then
        MsgBox "Date of Acceptance Required"
        Exit Sub
    End If
    
    If StepupDate.value = "" Then
        MsgBox "Step-Up Date Required"
        Exit Sub
    End If
    
    
    ClientUpdateForm.JTC_Reject.BackColor = &H8000000F
    ClientUpdateForm.JTC_Accept.BackColor = &H8000000A
    
    ClientUpdateForm.JTC_Return_Phase.Caption = 1
    
    ClientUpdateForm.JTC_Accept_Reject_Date.Caption = DateOfAcceptance.value
    ClientUpdateForm.JTC_Accept_Reject_Date_Label.Visible = True
    ClientUpdateForm.JTC_Accept_Reject_Date_Label.Caption = "Date Accepted:"
    
    ClientUpdateForm.JTC_Return_Phase_Label.Enabled = True
    ClientUpdateForm.JTC_Return_Stepup_Date_Label.Visible = True
    ClientUpdateForm.JTC_Return_Stepup_Date.Caption = StepupDate.value
    
    ClientUpdateForm.JTC_Referred_To_Label.Visible = False
    ClientUpdateForm.JTC_Referred_To.Caption = ""
    
    ClientUpdateForm.JTC_Pushback_Reason_Label.Enabled = True
    ClientUpdateForm.JTC_Treatment_Stepdown.Enabled = True
    ClientUpdateForm.JTC_Treatment_Provider_Update.Enabled = True
    ClientUpdateForm.JTC_Treatment_Provider_Remain.Enabled = True
    ClientUpdateForm.JTC_Service_Add.Enabled = True
    ClientUpdateForm.JTC_Service_Discharge.Enabled = True
    ClientUpdateForm.JTC_Return_Treatment_Provider_Label.Enabled = True
    ClientUpdateForm.JTC_Stepdown_Label.Enabled = True
    ClientUpdateForm.JTC_Return_Service_Box_Label.Enabled = True

    Me.Hide
End Sub




