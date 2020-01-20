VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Modal_Standard_Add_Condition 
   Caption         =   "Add Condition"
   ClientHeight    =   7608
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   9348.001
   OleObjectBlob   =   "Modal_Standard_Add_Condition.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Modal_Standard_Add_Condition"
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
    Set ctl = Modal_Standard_Add_Condition.Start_Date

    Call DateValidation(ctl, Cancel)
End Sub

''''''''''''''''''
'''''BUTTONS''''''
''''''''''''''''''

Private Sub InsertDoH_Click()
    Start_Date = ClientUpdateForm.DateOfHearing
End Sub

Private Sub Cancel_Click()
    Unload Modal_Standard_Add_Condition
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

    With ClientUpdateForm.Standard_Return_Condition_Box
        .ColumnCount = 10
        .ColumnWidths = "50;50;50;50;0;0;0;0;0;0;"
        .AddItem
        .List(ClientUpdateForm.Standard_Return_Condition_Box.ListCount - 1, 0) = ConditionType
        .List(ClientUpdateForm.Standard_Return_Condition_Box.ListCount - 1, 1) = Provider
        .List(ClientUpdateForm.Standard_Return_Condition_Box.ListCount - 1, 2) = Start_Date
        'end date
        .List(ClientUpdateForm.Standard_Return_Condition_Box.ListCount - 1, 4) = "New"
        'nature
        .List(ClientUpdateForm.Standard_Return_Condition_Box.ListCount - 1, 6) = encodeReasons(Reason1, Reason2, Reason3, Reason4, Reason5)
        .List(ClientUpdateForm.Standard_Return_Condition_Box.ListCount - 1, 9) = Notes
    End With

    With Modal_Standard_Drop_Condition.Condition_Box
        .ColumnCount = 10
        .ColumnWidths = "50;50;50;50;0;0;0;0;0;0;"
        .AddItem
        .List(Modal_Standard_Drop_Condition.Condition_Box.ListCount - 1, 0) = ConditionType
        .List(Modal_Standard_Drop_Condition.Condition_Box.ListCount - 1, 1) = Provider
        .List(Modal_Standard_Drop_Condition.Condition_Box.ListCount - 1, 2) = Start_Date
        'end date
        .List(Modal_Standard_Drop_Condition.Condition_Box.ListCount - 1, 4) = "New"
    End With
    
    Select Case ConditionType.value
        Case "Restitution"
            ClientUpdateForm.Standard_Restitution.Visible = True
            ClientUpdateForm.Standard_Restitution_Label.Visible = True
            ClientUpdateForm.Standard_Restitution.Caption = Restitution.value
        Case "Comm. Serv"
            ClientUpdateForm.Standard_Comm_Service.Visible = True
            ClientUpdateForm.Standard_Comm_Service_Label.Visible = True
            ClientUpdateForm.Standard_Comm_Service.Caption = Comm_Service.value
    End Select

    Unload Modal_Standard_Add_Condition
End Sub

