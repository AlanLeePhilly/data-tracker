VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Modal_JTC_Drop_Condition 
   Caption         =   "JTC Drop Condition"
   ClientHeight    =   8220.001
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   6444
   OleObjectBlob   =   "Modal_JTC_Drop_Condition.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Modal_JTC_Drop_Condition"
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
    Set ctl = Modal_JTC_Drop_Condition.Discharge_Date

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
    Notes.value = ""
    Modal_JTC_Drop_Condition.Hide
End Sub

Private Sub Continue_Click()
    Dim i As Integer
    If Not HasContent(Discharge_Date) Then
        MsgBox "Discharge Date Required"
        Exit Sub
    End If
    Dim returnBox As Object
    Set returnBox = ClientUpdateForm.JTC_Return_Condition_Box

    With Condition_Box
        For i = 0 To .ListCount - 1
            If .Selected(i) = True Then
            
                Select Case .List(i, 0)
                    Case "Restitution"
                        ClientUpdateForm.JTC_Restitution.Visible = False
                        ClientUpdateForm.JTC_Restitution_Label.Visible = False
                        ClientUpdateForm.JTC_Restitution.Caption = ""
                    Case "Comm. Serv"
                        ClientUpdateForm.JTC_Comm_Service.Visible = False
                        ClientUpdateForm.JTC_Comm_Service_Label.Visible = False
                        ClientUpdateForm.JTC_Comm_Service.Caption = ""
                End Select
            
                If returnBox.List(i, 2) = Discharge_Date Then
                    returnBox.RemoveItem (i)
                    .RemoveItem (i)
                Else
                    returnBox.List(i, 3) = Discharge_Date
                    'location
                    returnBox.List(i, 5) = Nature
                    returnBox.List(i, 6) = Reason1
                    returnBox.List(i, 7) = Reason2
                    returnBox.List(i, 8) = Reason3
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
    Else
        Reasons_Label.Enabled = False
        Reason1.Enabled = False
        Reason1.value = "N/A"
        Reason2.Enabled = False
        Reason2.value = "N/A"
        Reason3.Enabled = False
        Reason3.value = "N/A"
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Modal_JTC_Drop_Condition.Hide
        Cancel = True
    End If
End Sub
