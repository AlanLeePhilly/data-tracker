VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Modal_JTC_Reject 
   Caption         =   "JTC - Reject Client"
   ClientHeight    =   3765
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   6252
   OleObjectBlob   =   "Modal_JTC_Reject.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Modal_JTC_Reject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Continue_Click()
    ClientUpdateForm.JTC_Reject.BackColor = &H8000000A
    ClientUpdateForm.JTC_Accept.BackColor = &H8000000F

    ClientUpdateForm.JTC_Return_Phase = "Rejected"

    ClientUpdateForm.JTC_Accept_Reject_Date.Caption = DateOfRejection.value
    ClientUpdateForm.JTC_Accept_Reject_Date_Label.Visible = True
    ClientUpdateForm.JTC_Accept_Reject_Date_Label.Caption = "Date Rejected:"

    ClientUpdateForm.JTC_Referred_To_Label.Visible = True
    ClientUpdateForm.JTC_Referred_To.Caption = ReferredTo

    ClientUpdateForm.JTC_Return_Stepup_Date_Label.Visible = False
    ClientUpdateForm.JTC_Return_Stepup_Date.Caption = ""

    ClientUpdateForm.JTC_Pushback_Reason_Label.Enabled = False
    ClientUpdateForm.JTC_Pushback_Reason1.Caption = ""
    ClientUpdateForm.JTC_Pushback_Reason2.Caption = ""
    ClientUpdateForm.JTC_Pushback_Reason3.Caption = ""

    ClientUpdateForm.JTC_Treatment_Stepdown.Enabled = False
    ClientUpdateForm.JTC_Treatment_Stepdown.BackColor = &H8000000F
    ClientUpdateForm.JTC_Treatment_Provider_Update.Enabled = False
    ClientUpdateForm.JTC_Treatment_Provider_Update.BackColor = &H8000000F
    ClientUpdateForm.JTC_Treatment_Provider_Remain.Enabled = False
    ClientUpdateForm.JTC_Treatment_Provider_Remain.BackColor = &H8000000F

    ClientUpdateForm.JTC_Return_Treatment_Provider_Label.Enabled = False
    ClientUpdateForm.JTC_Return_Treatment_Provider.Caption = ""
    ClientUpdateForm.JTC_Stepdown_Label.Enabled = False
    ClientUpdateForm.JTC_Return_Stepdown_Date.Caption = ""

    ClientUpdateForm.JTC_Service_Add.Enabled = False
    ClientUpdateForm.JTC_Service_Discharge.Enabled = False


    ClientUpdateForm.JTC_Return_Service_Box_Label.Enabled = False

    Me.Hide


End Sub

Private Sub Cancel_Click()
    Me.Hide
End Sub


Private Sub DateOfRejection_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Set ctl = Modal_JTC_Reject.DateOfRejection
    'send to date validation
    Call DateValidation(ctl, Cancel)
End Sub

Private Sub InsertDoH_Click()
    DateOfRejection.value = ClientUpdateForm.DateOfHearing.value
End Sub

