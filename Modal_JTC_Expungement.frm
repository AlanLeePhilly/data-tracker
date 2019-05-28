VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Modal_JTC_Expungement 
   Caption         =   "JTC Expungement"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6150
   OleObjectBlob   =   "Modal_JTC_Expungement.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Modal_JTC_Expungement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub ExpungementDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Set ctl = Modal_JTC_Expungement.ExpungementDate
    'send to date validation
    Call DateValidation(ctl, Cancel)
End Sub

Private Sub Cancel_Click()
    ClientUpdateForm.JTC_Expungement.BackColor = &H8000000F
    ClientUpdateForm.JTC_Return_Phase.Caption = ""
    ClientUpdateForm.JTC_Accept_Reject_Date.Caption = ""
    Unload Me
End Sub

Private Sub InsertDoH_Click()
    ExpungementDate.value = ClientUpdateForm.DateOfHearing.value
End Sub

Private Sub Continue_Click()

    If ExpungementDate.value = "" Then
        MsgBox "Date of Expungement Required"
        Exit Sub
    End If

    ClientUpdateForm.JTC_Expungement.BackColor = &H8000000A

    ClientUpdateForm.JTC_Return_Phase.Caption = "Graduated, Record Expunged"

    ClientUpdateForm.JTC_Accept_Reject_Date.Caption = ExpungementDate.value
    ClientUpdateForm.JTC_Accept_Reject_Date_Label.Visible = True
    ClientUpdateForm.JTC_Accept_Reject_Date_Label.Caption = "Date Expunged:"

    ClientUpdateForm.JTC_Return_Phase_Label.Enabled = True

    ClientUpdateForm.JTC_Return_Stepup_Date_Label.Visible = False
    ClientUpdateForm.JTC_Referred_To_Label.Visible = False
    ClientUpdateForm.JTC_Pushback_Reason_Label.Visible = False
    ClientUpdateForm.JTC_Treatment_Provider_Update.Visible = False
    ClientUpdateForm.JTC_Stepdown_Label.Visible = False

    Me.Hide
End Sub

