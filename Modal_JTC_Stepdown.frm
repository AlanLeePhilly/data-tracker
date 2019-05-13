VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Modal_JTC_Stepdown 
   Caption         =   "JTC - Treatment Stepdown"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5490
   OleObjectBlob   =   "Modal_JTC_Stepdown.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Modal_JTC_Stepdown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cancel_Click()
    Me.Hide
End Sub

Private Sub Continue_Click()
    If Stepdown_Date = "" Then
        MsgBox "Date required"
        Exit Sub
    End If

    ClientUpdateForm.JTC_Stepdown_Label.Visible = True
    ClientUpdateForm.JTC_Return_Stepdown_Date.Caption = Stepdown_Date.value
    ClientUpdateForm.JTC_Return_Treatment_Provider.Caption = ClientUpdateForm.JTC_Fetch_Treatment_Provider.Caption

    'color buttons
    ClientUpdateForm.JTC_Treatment_Provider_Remain.BackColor = &H8000000F
    ClientUpdateForm.JTC_Treatment_Stepdown.BackColor = &H8000000A
    ClientUpdateForm.JTC_Treatment_Provider_Update.BackColor = &H8000000F
    ClientUpdateForm.JTC_Treatment_Discharge.BackColor = &H8000000F
    
    Me.Hide
End Sub

Private Sub InsertDoH_Click()
    Stepdown_Date.value = ClientUpdateForm.DateOfHearing.value
End Sub

Private Sub UserForm_Initialize()
    Provider.Caption = ClientUpdateForm.JTC_Fetch_Treatment_Provider
End Sub
Private Sub Stepdown_Date_Enter()

    Stepdown_Date.value = CalendarForm.GetDate(RangeOfYears:=5)
    
End Sub
Private Sub Stepdown_Date_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Set ctl = Me.Stepdown_Date

    Call DateValidation(ctl, Cancel)
End Sub
