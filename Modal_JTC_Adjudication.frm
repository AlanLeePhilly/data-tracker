VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Modal_JTC_Adjudication 
   Caption         =   "JTC Adjudication"
   ClientHeight    =   9975.001
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   6468
   OleObjectBlob   =   "Modal_JTC_Adjudication.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Modal_JTC_Adjudication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    Call addPetitionsToBox(PetitionBox)
End Sub

'''''''''''''
'VALIDATIONS'
'''''''''''''
Private Sub Adjudication_Date_Enter()
    Adjudication_Date.value = CalendarForm.GetDate(RangeOfYears:=5)
End Sub
Private Sub Adjudication_Date_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Set ctl = Modal_JTC_Adjudication.Adjudication_Date

    Call DateValidation(ctl, Cancel)
End Sub


''''''''''''''''''
'''''BUTTONS''''''
''''''''''''''''''

Private Sub InsertDoH_Click()
    Adjudication_Date = ClientUpdateForm.DateOfHearing
End Sub
Private Sub Cancel_Click()
    Modal_JTC_Adjudication.Hide
End Sub


'''''''''''''''''''''
'''''FORM LOGIC''''''
'''''''''''''''''''''

Private Sub Type_of_Change()
    If Type_of = "Technical Violations" Then
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
    If Not HasContent(Adjudication_Date) Then
        MsgBox "Date of Adjudication Required"
        Exit Sub
    End If
    If Not HasContent(DA) Then
        MsgBox "DA Name Required"
        Exit Sub
    End If
    If Type_of.value = "N/A" Then
        MsgBox "Type of Adjudication Required"
        Exit Sub
    End If
    If PetitionBox.value = Null Then
        MsgBox "Please select a petition"
        Exit Sub
    End If

    ClientUpdateForm.JTC_Adjudication_Update.BackColor = selectedColor
    ClientUpdateForm.JTC_Adjudication_Remain.BackColor = unselectedColor
    ClientUpdateForm.JTC_Return_Adjudication.Caption = "Yes"

    Modal_JTC_Adjudication.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Modal_JTC_Adjudication.Hide
        Cancel = True
    End If
End Sub
