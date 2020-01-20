VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Modal_JTC_Add_Condition 
   Caption         =   "JTC Add Condition"
   ClientHeight    =   8772.001
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   10872
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
            Label.Visible = True
            Label.Caption = "Restitution:  $"
            Amount.Visible = True
            Amount.value = ""
        Case "Comm. Serv"
            Label.Visible = True
            Label.Caption = "Comm. Service Hrs: "
            Amount.Visible = True
            Amount.value = ""
        Case "Court Costs"
            Label.Visible = True
            Label.Caption = "Court Costs:  $"
            Amount.Visible = True
            Amount.value = ""
        Case Else
            Label.Visible = False
            Amount.Visible = False
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
        .List(ClientUpdateForm.JTC_Return_Condition_Box.ListCount - 1, 6) = encodeReasons(Reason1, Reason2, Reason3, Reason4, Reason5)
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
            ClientUpdateForm.JTC_Restitution.Caption = Amount.value
        Case "Comm. Serv"
            ClientUpdateForm.JTC_Comm_Service.Visible = True
            ClientUpdateForm.JTC_Comm_Service_Label.Visible = True
            ClientUpdateForm.JTC_Comm_Service.Caption = Amount.value
        Case "Court Costs"
            ClientUpdateForm.JTC_Court_Costs.Visible = True
            ClientUpdateForm.JTC_Court_Costs_Label.Visible = True
            ClientUpdateForm.JTC_Court_Costs.Caption = Amount.value
    End Select
    
    Unload Modal_JTC_Add_Condition
End Sub


