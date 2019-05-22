VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Modal_JTC_Add_Condition 
   Caption         =   "JTC Add Condition"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6360
   OleObjectBlob   =   "Modal_JTC_Add_Condition.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Modal_JTC_Add_Condition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
                .List(ClientUpdateForm.JTC_Return_Condition_Box.ListCount - 1, 9) = ""
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
    
    Unload Modal_JTC_Add_Condition
End Sub


