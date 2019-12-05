VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Modal_JTC_Stepup 
   Caption         =   "JTC - Phase Step-Up"
   ClientHeight    =   3195
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   5688
   OleObjectBlob   =   "Modal_JTC_Stepup.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Modal_JTC_Stepup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Continue_Click()
    If New_Stepup_Date.value = "" Then
        MsgBox "New Date Required"
        Exit Sub
    End If

    'advance phase by one
    Select Case ClientUpdateForm.JTC_Fetch_Phase
        Case "Assessment"
            ClientUpdateForm.JTC_Return_Phase.Caption = 1
        Case 1
            ClientUpdateForm.JTC_Return_Phase.Caption = 2
        Case 2
            ClientUpdateForm.JTC_Return_Phase.Caption = 3
    End Select

    'set new step-up date
    ClientUpdateForm.JTC_Return_Stepup_Date.Caption = New_Stepup_Date.value

    'hide pushback display
    ClientUpdateForm.JTC_Pushback_Reason_Label.Visible = False
    ClientUpdateForm.JTC_Pushback_Reason1.Caption = ""
    ClientUpdateForm.JTC_Pushback_Reason2.Caption = ""
    ClientUpdateForm.JTC_Pushback_Reason3.Caption = ""

    'color buttons
    ClientUpdateForm.JTC_Phase_Remain.BackColor = &H8000000F
    ClientUpdateForm.JTC_Phase_Stepup.BackColor = &H8000000A
    ClientUpdateForm.JTC_Phase_Pushback.BackColor = &H8000000F
    ClientUpdateForm.JTC_Discharge.BackColor = &H8000000F

    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    Current_Phase.Caption = ClientUpdateForm.JTC_Fetch_Phase.Caption
End Sub

Private Sub Cancel_Click()

    New_Stepup_Date = ""
    Modal_JTC_Stepup.Hide
End Sub

Private Sub New_Stepup_Date_Enter()

    New_Stepup_Date.value = CalendarForm.GetDate(RangeOfYears:=5)

End Sub



Private Sub New_Stepup_Date_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Set ctl = Me.New_Stepup_Date

    'send to date validation
    Call DateValidation(ctl, Cancel)
End Sub
