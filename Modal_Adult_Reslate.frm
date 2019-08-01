VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Modal_Adult_Reslate 
   Caption         =   "UserForm1"
   ClientHeight    =   4995
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8790.001
   OleObjectBlob   =   "Modal_Adult_Reslate.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Modal_Adult_Reslate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Continue_Click()
    Select Case Hearing_Outcome.value
        Case "Granted"

            Adult_Reslate_Juvenile_Petition.Show
        Case Else
            ClientUpdateForm.Adult_Return_Reslate.Caption _
                = Hearing_Outcome.value
    End Select

    Modal_Adult_Reslate.Hide
End Sub

Private Sub InsertDoH_Click()
    Reslate_Date = ClientUpdateForm.DateOfHearing.value
End Sub

Private Sub Reslate_Date_Enter()
    Reslate_Date.value = CalendarForm.GetDate(RangeOfYears:=5)
End Sub

Private Sub Cancel_Click()
    Unload Me
End Sub
