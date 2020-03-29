VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Modal_Standard_Admission 
   Caption         =   "Admission"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8805.001
   OleObjectBlob   =   "Modal_Standard_Admission.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Modal_Standard_Admission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    Call addChargesToBox(PetitionBox)
End Sub
'''''''''''''
'VALIDATIONS'
'''''''''''''
Private Sub Admission_Date_Enter()
    Admission_Date.value = CalendarForm.GetDate(RangeOfYears:=5)
End Sub
Private Sub Admission_Date_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Set ctl = Modal_Standard_Admission.Admission_Date

    Call DateValidation(ctl, Cancel)
End Sub

''''''''''''''''''
'''''BUTTONS''''''
''''''''''''''''''

Private Sub InsertDoH_Click()
    Admission_Date = ClientUpdateForm.DateOfHearing
End Sub

Private Sub Cancel_Click()
    Modal_Standard_Admission.Hide
End Sub

'''''''''''''''''''''''
'''''SUBMIT LOGIC''''''
'''''''''''''''''''''''

Private Sub Continue_Click()
    'VALIDATIONS
    If Not HasContent(Admission_Date) Then
        MsgBox "Date of Admission Required"
        Exit Sub
    End If
    If Result.value = "N/A" Then
        MsgBox "Result of Admission Required"
        Exit Sub
    End If
    If PetitionBox.value = Null Then
        MsgBox "Please select a petition"
        Exit Sub
    End If

    ClientUpdateForm.Standard_Admission_Update.BackColor = selectedColor
    ClientUpdateForm.Standard_Admission_Remain.BackColor = unselectedColor
    ClientUpdateForm.Standard_Return_Admission.Caption = "Yes"


    Modal_Standard_Admission.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Modal_Standard_Admission.Hide
        Cancel = True
    End If
End Sub
