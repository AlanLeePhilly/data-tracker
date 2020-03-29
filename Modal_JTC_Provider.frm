VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Modal_JTC_Provider 
   Caption         =   "JTC - Change Provider"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6780
   OleObjectBlob   =   "Modal_JTC_Provider.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Modal_JTC_Provider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cancel_Click()
    New_Treatment_Provider.value = ""
    Me.Hide
End Sub

Private Sub Continue_Click()
    If New_Treatment_Provider.value = "" Then
        MsgBox "New Provider Required"
        Exit Sub
    End If

    If Referral_Date.value = "" Then
        MsgBox "Referral Date Required"
        Exit Sub
    End If

    ClientUpdateForm.JTC_Return_Treatment_Provider.Caption = New_Treatment_Provider.value
    ClientUpdateForm.JTC_Return_Stepdown_Date.Caption = ""
    ClientUpdateForm.JTC_Stepdown_Label.Visible = False
    'color buttons
    ClientUpdateForm.JTC_Treatment_Provider_Remain.BackColor = &H8000000F
    ClientUpdateForm.JTC_Treatment_Stepdown.BackColor = &H8000000F
    ClientUpdateForm.JTC_Treatment_Provider_Update.BackColor = &H8000000A
    ClientUpdateForm.JTC_Treatment_Discharge.BackColor = &H8000000F
    Me.Hide
End Sub

Private Sub InsertDoH_Click()
    Referral_Date.value = ClientUpdateForm.DateOfHearing.value
End Sub

Private Sub Referral_Date_Enter()

    Referral_Date.value = CalendarForm.GetDate(RangeOfYears:=5)

End Sub




