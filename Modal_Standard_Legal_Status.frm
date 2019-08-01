VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Modal_Standard_Legal_Status 
   Caption         =   "Update Legal Status"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14655
   OleObjectBlob   =   "Modal_Standard_Legal_Status.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Modal_Standard_Legal_Status"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False










Private Sub Current_Detailed_Outcome_Change()
    If isTerminal("Legal Status", Current_Detailed_Outcome.value) Then
        New_Legal_Status.Enabled = False
        New_Legal_Status.value = ""
        New_Start_Date.Enabled = False
        New_Start_Date.value = ""
        New_Notes.Enabled = False
        New_Notes.value = ""
        New_Legal_Status_Label.Enabled = False
        New_Start_Date_Label.Enabled = False
        New_Notes_Label.Enabled = False
    Else
        New_Legal_Status.Enabled = True
        New_Start_Date.Enabled = True
        New_Notes.Enabled = True
        New_Legal_Status_Label.Enabled = True
        New_Start_Date_Label.Enabled = True
        New_Notes_Label.Enabled = True

    End If
End Sub

''''''''''''''''
'INITIALIZATION'
''''''''''''''''
Private Sub UserForm_Initialize()
    Current_Legal_Status = ClientUpdateForm.Standard_Fetch_Legal_Status
End Sub

'''''''''''''
'VALIDATIONS'
'''''''''''''
Private Sub Current_Discharge_Date_Enter()
    Current_Discharge_Date.value = CalendarForm.GetDate(RangeOfYears:=5)
End Sub
Private Sub Current_Discharge_Date_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Set ctl = Modal_Standard_Legal_Status.Current_Discharge_Date

    Call DateValidation(ctl, Cancel)
End Sub
Private Sub New_Start_Date_Enter()
    New_Start_Date.value = CalendarForm.GetDate(RangeOfYears:=5)
End Sub
Private Sub New_Start_Date_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Set ctl = Modal_Standard_Legal_Status.New_Start_Date

    Call DateValidation(ctl, Cancel)
End Sub

''''''''''''''''''
'''''BUTTONS''''''
''''''''''''''''''

Private Sub InsertDoH1_Click()
    Current_Discharge_Date = ClientUpdateForm.DateOfHearing
End Sub
Private Sub InsertDoH2_Click()
    New_Start_Date = ClientUpdateForm.DateOfHearing
End Sub
Private Sub Cancel_Click()
    Unload Modal_Standard_Legal_Status
End Sub


'''''''''''''''''''''
'''''FORM LOGIC''''''
'''''''''''''''''''''

Private Sub Current_Discharge_Nature_Change()
    If Current_Discharge_Nature = "Negative" Then
        Reasons_Label.Enabled = True
        Reason1.Enabled = True
        Reason2.Enabled = True
        Reason3.Enabled = True
        Reason4.Enabled = True
        Reason5.Enabled = True
    Else
        Reasons_Label.Enabled = False
        Reason1.Enabled = False
        Reason1.value = "N/A"
        Reason2.Enabled = False
        Reason2.value = "N/A"
        Reason3.Enabled = False
        Reason3.value = "N/A"
        Reason4.Enabled = False
        Reason4.value = "N/A"
        Reason5.Enabled = False
        Reason5.value = "N/A"
    End If
End Sub

'''''''''''''''''''''''
'''''SUBMIT LOGIC''''''
'''''''''''''''''''''''

Private Sub Continue_Click()
    'VALIDATIONS
    If Not HasContent(Current_Discharge_Date) Then
        MsgBox "Date of Discharge Required"
        Exit Sub
    End If

    If Not isTerminal("Legal Status", Current_Detailed_Outcome.value) Then
        If Not HasContent(New_Start_Date) Then
            MsgBox "Start Date Required"
            Exit Sub
        End If
        If New_Legal_Status.value = "N/A" Then
            MsgBox "New Legal Status Required"
            Exit Sub
        End If
    End If

    ClientUpdateForm.Standard_Legal_Status_Update.BackColor = selectedColor
    ClientUpdateForm.Standard_Legal_Status_Remain.BackColor = unselectedColor
    ClientUpdateForm.Standard_Return_Legal_Status.Caption = New_Legal_Status
    Modal_Standard_Legal_Status.Hide

End Sub
