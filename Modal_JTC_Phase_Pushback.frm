VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Modal_JTC_Phase_Pushback 
   Caption         =   "JTC - Phase Pushback"
   ClientHeight    =   5685
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   5748
   OleObjectBlob   =   "Modal_JTC_Phase_Pushback.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Modal_JTC_Phase_Pushback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cancel_Click()
    New_Stepup_Date = ""
    reason = ""
    Modal_JTC_Phase_Pushback.Hide
End Sub

Private Sub Continue_Click()
    'sub for submission of JTC Pushback Modal

    'validate presence of date
    If Not HasContent(New_Stepup_Date) Then
        MsgBox "New Step-Up Date Required"
        Exit Sub
    End If

    'validate presence of reason
    If Reason1.value = "" Then
        MsgBox "Pushback Reason Required"
        Exit Sub
    End If

    'maintain current phase
    ClientUpdateForm.JTC_Return_Phase.Caption = ClientUpdateForm.JTC_Fetch_Phase.Caption

    'Print new date and reason for pushback
    ClientUpdateForm.JTC_Return_Stepup_Date.Caption = New_Stepup_Date
    ClientUpdateForm.JTC_Pushback_Reason_Label.Visible = True
    ClientUpdateForm.JTC_Pushback_Reason1.Caption = Reason1
    If Not Reason1 = "N/A" Then
        If Not Reason2 = "N/A" Then
            ClientUpdateForm.JTC_Pushback_Reason2.Caption = Reason2
        End If
        If Not Reason3 = "N/A" Then
            ClientUpdateForm.JTC_Pushback_Reason3.Caption = Reason3
        End If
    End If

    'update button color
    ClientUpdateForm.JTC_Phase_Remain.BackColor = &H8000000F
    ClientUpdateForm.JTC_Phase_Stepup.BackColor = &H8000000F
    ClientUpdateForm.JTC_Phase_Pushback.BackColor = &H8000000A
    ClientUpdateForm.JTC_Discharge.BackColor = &H8000000F


    Modal_JTC_Phase_Pushback.Hide
End Sub
Private Sub New_Stepup_Date_Enter()

    New_Stepup_Date.value = CalendarForm.GetDate(RangeOfYears:=5)

End Sub


Private Sub New_Stepup_Date_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Set ctl = Me.New_Stepup_Date

    'send to date validation
    Call DateValidation(ctl, Cancel)
End Sub


Private Sub Reason1_Change()
    If Reason1 = "N/A" Then
        Reason2.Enabled = False
        Reason2_Label.Enabled = False
        Reason2 = "N/A"
    Else
        Reason2.Enabled = True
        Reason2_Label.Enabled = True
    End If
End Sub
Private Sub Reason2_Change()
    If Reason2 = "N/A" Then
        Reason3.Enabled = False
        Reason3_Label.Enabled = False
        Reason3 = "N/A"
    Else
        Reason3.Enabled = True
        Reason3_Label.Enabled = True
    End If
End Sub

Private Sub UserForm_Initialize()
    Current_Phase.Caption = ClientUpdateForm.JTC_Fetch_Phase.Caption
    Current_Stepup_Date.Caption = ClientUpdateForm.JTC_Fetch_Stepup_Date.Caption
End Sub
