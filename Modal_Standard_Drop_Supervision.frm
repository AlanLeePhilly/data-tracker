VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Modal_Standard_Drop_Supervision 
   Caption         =   "Drop Supervision"
   ClientHeight    =   9645.001
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6765
   OleObjectBlob   =   "Modal_Standard_Drop_Supervision.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Modal_Standard_Drop_Supervision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''
'VALIDATIONS'
'''''''''''''
Private Sub Discharge_Date_Enter()
    Discharge_Date.value = CalendarForm.GetDate(RangeOfYears:=5)
End Sub
Private Sub Discharge_Date_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Set ctl = Modal_Standard_Drop_Supervision.Discharge_Date

    Call DateValidation(ctl, Cancel)
End Sub

''''''''''''''''''
'''''BUTTONS''''''
''''''''''''''''''

Private Sub InsertDoH_Click()
    Discharge_Date = ClientUpdateForm.DateOfHearing
End Sub

Private Sub Cancel_Click()
    
    Discharge_Date.value = ""
    Nature.value = "N/A"
    Reason1.value = "N/A"
    Reason2.value = "N/A"
    Reason3.value = "N/A"
    Reason4.value = "N/A"
    Reason5.value = "N/A"
    Notes.value = ""
    Modal_Standard_Drop_Supervision.Hide
End Sub

Private Sub Continue_Click()
    Dim i As Integer
    If Not HasContent(Discharge_Date) Then
        MsgBox "Discharge Date Required"
        Exit Sub
    End If
    Dim returnBox As Object
    Set returnBox = ClientUpdateForm.Standard_Return_Supervision_Box

    With Supervision_Box
        For i = 0 To .ListCount - 1
            If .Selected(i) = True Then
                If returnBox.List(i, 2) = Discharge_Date Then
                    returnBox.RemoveItem (i)
                    .RemoveItem (i)
                Else
                    returnBox.List(i, 3) = Discharge_Date
                    returnBox.List(i, 5) = Nature
                    returnBox.List(i, 6) = encodeReasons(Reason1, Reason2, Reason3, Reason4, Reason5)
                    returnBox.List(i, 9) = Notes
                End If
            End If
        Next
    End With
    Call Cancel_Click
End Sub


Private Sub Nature_Change()
    If Nature = "Negative" Then
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

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Modal_Standard_Drop_Supervision.Hide
        Cancel = True
    End If
End Sub

