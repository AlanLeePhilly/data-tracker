VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Modal_JTC_Add_Service 
   Caption         =   "JTC - Add Service"
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   -75
   ClientWidth     =   6330
   OleObjectBlob   =   "Modal_JTC_Add_Service.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Modal_JTC_Add_Service"
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
    Set ctl = Modal_Standard_Add_Supervision.Start_Date

    Call DateValidation(ctl, Cancel)
End Sub

''''''''''''''''''
'''''BUTTONS''''''
''''''''''''''''''

Private Sub InsertDoH_Click()
    Start_Date = ClientUpdateForm.DateOfHearing
End Sub

Private Sub Cancel_Click()
    ServiceOrdered.value = ""
    ServiceProvider.value = "None"
    Start_Date.value = ""
    Reason1.value = "N/A"
    Reason2.value = "N/A"
    Reason3.value = "N/A"
    Reason4.value = "N/A"
    Reason5.value = "N/A"
    Notes.value = ""
    Me.Hide
End Sub

Private Sub Continue_Click()

    If ServiceOrdered.value = "" Then
        MsgBox "'Service Ordered' is required"
        Exit Sub
    End If

    If ServiceProvider.value = "" Then
        MsgBox "'Provider' is required"
        Exit Sub
    End If


    'This is to add the new service to the table representing the post-hearing status.
    'It will be used to push the data to the database in the submit logic, so we need to include all relevant data
    With ClientUpdateForm.JTC_Return_Service_Box
        .ColumnCount = 10
        .ColumnWidths = "90;75;75;75;0;0;0;0;0;0;"
        .AddItem
        .List(ClientUpdateForm.JTC_Return_Service_Box.ListCount - 1, 0) = ServiceOrdered.value
        .List(ClientUpdateForm.JTC_Return_Service_Box.ListCount - 1, 1) = ServiceProvider.value
        .List(ClientUpdateForm.JTC_Return_Service_Box.ListCount - 1, 2) = Start_Date.value
        'end date
        .List(ClientUpdateForm.JTC_Return_Service_Box.ListCount - 1, 4) = "New"
        'nature
        .List(ClientUpdateForm.JTC_Return_Service_Box.ListCount - 1, 6) = encodeReasons(Reason1, Reason2, Reason3, Reason4, Reason5)
        .List(ClientUpdateForm.JTC_Return_Service_Box.ListCount - 1, 9) = Notes
        'notes
    End With

    'This is to add the new service to the table in the modal which allows discharging from a service
    'It needs to be done in case someone adds a service incorrectly, so that it can be removed from the list before submission
    With Modal_JTC_Drop_Service.Service_Box
        .ColumnCount = 10
        .ColumnWidths = "90;75;75;75;0;0;0;0;0;0;"
        .AddItem
        .List(Modal_JTC_Drop_Service.Service_Box.ListCount - 1, 0) = ServiceOrdered.value
        .List(Modal_JTC_Drop_Service.Service_Box.ListCount - 1, 1) = ServiceProvider.value
        .List(Modal_JTC_Drop_Service.Service_Box.ListCount - 1, 2) = ClientUpdateForm.DateOfHearing.value
        'end date
        .List(Modal_JTC_Drop_Service.Service_Box.ListCount - 1, 4) = "New"
    End With
    Call Cancel_Click
End Sub


Private Sub ServiceOrdered_Change()
    ServiceProvider.value = "None"
    If isResidential(ServiceOrdered) Then
        ServiceProvider.RowSource = "Residential_Supervision_Provider"
    Else
        ServiceProvider.RowSource = "Community_Based_Supervision_Provider"
    End If
End Sub

