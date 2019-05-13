VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Modal_JTC_Discharge 
   Caption         =   "JTC - Discharge Client"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   -75
   ClientWidth     =   6690
   OleObjectBlob   =   "Modal_JTC_Discharge.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Modal_JTC_Discharge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Continue_Click()
    'VALIDATION: cannot submit without choosing a detailed outcome
    If DetailedOutcome.value = "N/A" Then
        MsgBox "Detailed Outcome Required"
        Exit Sub
    End If
    
    ''''''''''''
    'VALIDATION NEEDED: only some detailed outcomes can match with certain Natures of Discharge
    ''''''''''''

    Modal_JTC_Discharge.Hide
    
    Select Case DetailedOutcome.value
        Case "Rearrested & Held (adult)", "Aged out", "Show Cause"
            ClientUpdateForm.JTC_Return_Phase.Caption = "Negative Discharge"
        Case "Acceptance Not Granted", "Transfer to Dependent", "Transfer to Other County", "Hosp. (Mental Health)", "Hosp. (Physical Health)", "Other", "Unknown"
            ClientUpdateForm.JTC_Return_Phase.Caption = "Neutral Discharge"
        Case "Positive Completion"
            ClientUpdateForm.JTC_Return_Phase.Caption = "Positive Discharge"
        Case Else
    End Select
    
    ClientUpdateForm.JTC_Return_Stepup_Date.Caption = "N/A"
    
    'hide reason for pushback display on the main update form
    ClientUpdateForm.JTC_Pushback_Reason_Label.Visible = False
    ClientUpdateForm.JTC_Pushback_Reason1.Caption = ""
    ClientUpdateForm.JTC_Pushback_Reason2.Caption = ""
    ClientUpdateForm.JTC_Pushback_Reason3.Caption = ""
    
    'color buttons on the main update form
    ClientUpdateForm.JTC_Phase_Remain.BackColor = &H8000000F
    ClientUpdateForm.JTC_Phase_Stepup.BackColor = &H8000000F
    ClientUpdateForm.JTC_Phase_Pushback.BackColor = &H8000000F
    ClientUpdateForm.JTC_Discharge.BackColor = &H8000000A
End Sub

Private Sub Cancel_Click()
    'clear form and hide it
    Call Clear_Click
    Modal_JTC_Discharge.Hide
End Sub

Private Sub DetailedOutcome_Change()
    Select Case DetailedOutcome.value
        Case "Rearrested & Held (adult)", "Positive Completion", "Aged Out", "Transfer to Dependent", "Transfer to Other County"
            New_CR_Label.Enabled = False
            New_CR.Enabled = False
            New_CR.value = "N/A"
        'Case "Acceptance Not Granted", "Show Cause", "Hosp. (Mental Health)", "Hosp. (Physical Health)", "Other", "Unknown"
        Case Else
            New_CR_Label.Enabled = True
            New_CR.Enabled = True
            
    End Select
    
    Select Case DetailedOutcome.value
        Case "Rearrested & Held (adult)", "Aged Out", "Show Cause"
            ReasonForDischarge1.Enabled = True
            LabelReason1.Enabled = True
            ReasonForDischarge2.Enabled = True
            LabelReason2.Enabled = True
            ReasonForDischarge3.Enabled = True
            LabelReason3.Enabled = True
        Case Else
            ReasonForDischarge1.Enabled = False
            ReasonForDischarge1.value = "N/A"
            LabelReason1.Enabled = False
            ReasonForDischarge2.Enabled = False
            ReasonForDischarge2.value = "N/A"
            LabelReason2.Enabled = False
            ReasonForDischarge3.Enabled = False
            ReasonForDischarge3.value = "N/A"
            LabelReason3.Enabled = False
    End Select
End Sub



Private Sub Clear_Click()

    DetailedOutcome.value = "N/A"
    New_CR.value = "N/A"
    ReasonForDischarge1.value = "N/A"
    ReasonForDischarge2.value = "N/A"
    ReasonForDischarge3.value = "N/A"
End Sub

Private Sub UserForm_Initialize()
    New_CR_Label.Enabled = False
    New_CR.Enabled = False
End Sub
