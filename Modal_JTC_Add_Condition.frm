VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Modal_JTC_Add_Condition 
   Caption         =   "JTC Add Condition"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9480.001
   OleObjectBlob   =   "Modal_JTC_Add_Condition.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Modal_JTC_Add_Condition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ConditionType_Change()
    Select Case ConditionType.value
        Case "Restitution"
            Restitution_Label.Visible = True
            Restitution.Visible = True
            Restitution.value = ""
            
            Comm_Service_Label.Visible = False
            Comm_Service.Visible = False
            Comm_Service.value = ""
        Case "Comm. Serv"
            Restitution_Label.Visible = False
            Restitution.Visible = False
            Restitution.value = ""
            
            Comm_Service_Label.Visible = True
            Comm_Service.Visible = True
            Comm_Service.value = ""
        Case Else
            Restitution_Label.Visible = False
            Restitution.Visible = False
            Restitution.value = ""
            
            Comm_Service_Label.Visible = False
            Comm_Service.Visible = False
            Comm_Service.value = ""
    End Select
End Sub

'''''''''''''
'VALIDATIONS'
'''''''''''''
Private Sub Start_Date_Enter()
    Start_Date.value = CalendarForm.GetDate(RangeOfYears:=5)
End Sub
Private Sub Start_Date_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Set ctl = Modal_JTC_Add_Condition.Start_Date

    Call DateValidation(ctl, Cancel)
End Sub

''''''''''''''''''
'''''BUTTONS''''''
''''''''''''''''''

Private Sub InsertDoH_Click()
    Start_Date = ClientUpdateForm.DateOfHearing
End Sub

Private Sub Cancel_Click()
    Unload Modal_JTC_Add_Condition
End Sub

'''''''''''''''''''''''
'''''SUBMIT LOGIC''''''
'''''''''''''''''''''''

Private Sub Continue_Click()
    'VALIDATIONS
    If ConditionType.value = "None" Then
        MsgBox "'Condition Ordered' Required"
        Exit Sub
    End If

    If Not HasContent(Start_Date) Then
        MsgBox "Start Date Required"
        Exit Sub
    End If

    With ClientUpdateForm.JTC_Return_Condition_Box
        .ColumnCount = 10
        .ColumnWidths = "90;75;75;75;0;0;0;0;0;0;"
        .AddItem
        .List(ClientUpdateForm.JTC_Return_Condition_Box.ListCount - 1, 0) = ConditionType
        .List(ClientUpdateForm.JTC_Return_Condition_Box.ListCount - 1, 1) = Provider
        .List(ClientUpdateForm.JTC_Return_Condition_Box.ListCount - 1, 2) = Start_Date
        'end date
        .List(ClientUpdateForm.JTC_Return_Condition_Box.ListCount - 1, 4) = "New"
        'nature
        .List(ClientUpdateForm.JTC_Return_Condition_Box.ListCount - 1, 6) = Reason1
        .List(ClientUpdateForm.JTC_Return_Condition_Box.ListCount - 1, 7) = Reason2
        .List(ClientUpdateForm.JTC_Return_Condition_Box.ListCount - 1, 8) = Reason3
        .List(ClientUpdateForm.JTC_Return_Condition_Box.ListCount - 1, 9) = Notes
    End With

    With Modal_JTC_Drop_Condition.Condition_Box
        .ColumnCount = 10
        .ColumnWidths = "90;75;75;75;0;0;0;0;0;0;"
        .AddItem
        .List(Modal_JTC_Drop_Condition.Condition_Box.ListCount - 1, 0) = ConditionType
        .List(Modal_JTC_Drop_Condition.Condition_Box.ListCount - 1, 1) = Provider
        .List(Modal_JTC_Drop_Condition.Condition_Box.ListCount - 1, 2) = Start_Date

        .List(Modal_JTC_Drop_Condition.Condition_Box.ListCount - 1, 4) = "New"
    End With
    
    Select Case ConditionType.value
        Case "Restitution"
            ClientUpdateForm.JTC_Restitution.Visible = True
            ClientUpdateForm.JTC_Restitution_Label.Visible = True
            ClientUpdateForm.JTC_Restitution.Caption = Restitution.value
        Case "Comm. Serv"
            ClientUpdateForm.JTC_Comm_Service.Visible = True
            ClientUpdateForm.JTC_Comm_Service_Label.Visible = True
            ClientUpdateForm.JTC_Comm_Service.Caption = Comm_Service.value
    End Select
    
    Unload Modal_JTC_Add_Condition
End Sub


