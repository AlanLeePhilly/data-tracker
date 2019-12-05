VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Modal_Adult_Decertification 
   Caption         =   "UserForm1"
   ClientHeight    =   4755
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   8448.001
   OleObjectBlob   =   "Modal_Adult_Decertification.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Modal_Adult_Decertification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub UserForm_Initialize()
    Prior_Status.Caption = ClientUpdateForm.Adult_Fetch_Decertification
    If ClientUpdateForm.Adult_Fetch_Decertification = "Filed" Then
        MultiPage1.value = 1
    Else
        MultiPage1.value = 0
    End If
End Sub


'''''''''''''
'VALIDATIONS'
'''''''''''''
Private Sub Motion_Date_Enter()
    Motion_Date.value = CalendarForm.GetDate(RangeOfYears:=5)
End Sub
Private Sub Motion_Date_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Set ctl = Modal_Adult_Decertification.Motion_Date

    Call DateValidation(ctl, Cancel)
End Sub

''''''''''''''''''
'''''BUTTONS''''''
''''''''''''''''''

Private Sub InsertDoH_Click()
    Motion_Date = ClientUpdateForm.DateOfHearing
End Sub

Private Sub Cancel_Click()
    Unload Modal_Adult_Decertification
End Sub

'''''''''''''''''''''''
'''''SUBMIT LOGIC''''''
'''''''''''''''''''''''

Private Sub Continue_Click()
    'VALIDATIONS
    If Prior_Status.Caption = "Filed" Then
        If Motion_Result.value = "N/A" Then
            MsgBox "Result of Motion Required"
            Exit Sub
        End If
    Else
        If Not HasContent(Motion_Date) Then
            MsgBox "Date of Motion Required"
            Exit Sub
        End If
    End If



    ClientUpdateForm.Adult_Decertification_Update.BackColor = selectedColor
    ClientUpdateForm.Adult_Decertification_Remain.BackColor = unselectedColor


    ClientUpdateForm.Adult_Return_Decertification.Caption = Hearing_Outcome.value

    Modal_Adult_Decertification.Hide
End Sub



