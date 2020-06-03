VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Modal_Apprehension 
   Caption         =   "Apprehension"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8985.001
   OleObjectBlob   =   "Modal_Apprehension.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Modal_Apprehension"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub IntakeDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Set ctl = Me.IntakeDate
    'send to date validation
    Call DateValidation(ctl, Cancel)
End Sub
Private Sub IntakeDate_Enter()
    IntakeDate.value = CalendarForm.GetDate(RangeOfYears:=5)
End Sub





Private Sub enable_supervisions()
    Dim ctl As MSForms.Control
    
    For Each ctl In Me.SupervisionsFrame.Controls
        ctl.Enabled = True
    Next ctl
End Sub

Private Sub disable_supervisions()
    Dim ctl As MSForms.Control
    On Error Resume Next
    For Each ctl In Me.SupervisionsFrame.Controls
        ctl.Enabled = False
    Next ctl
    
    Supv1.value = "None"
    Supv1Pro.value = "None"
    Supv1Re1.value = "N/A"
    Supv1Re2.value = "N/A"
    Supv1Re3.value = "N/A"
    Supv2.value = "None"
    Supv2Pro.value = "None"
    Supv2Re1.value = "N/A"
    Supv2Re2.value = "N/A"
    Supv2Re3.value = "N/A"
End Sub



Private Sub IntakeOutcomeRelease_Click()
    Call enable_supervisions
End Sub
Private Sub IntakeOutcomeRoll_Click()
    Call disable_supervisions
End Sub

Private Sub Cancel_Click()
    Unload Me
End Sub


Private Sub Submit_Click()
    If IntakeDate.value = "" Then
        MsgBox "Intake Date required"
        Exit Sub
    End If
    
    If IntakeOutcomeRelease = False And IntakeOutcomeRoll = False Then
        MsgBox "Intake Conference Outcome required"
        Exit Sub
    End If
    

    If IntakeOutcomeRoll = True Then
        MsgBox "Someday, this will be functional :D"
        Unload Me
        Exit Sub
    End If
    
    Call formSubmitStart(updateRow)
    
    Range(hFind("Active B/W?") & updateRow).value = Lookup("Generic_YNOU_Name")("No")

    For i = 15 To 1 Step -1
        If isNotEmptyOrZero(Range(hFind("FTA #" & i & " Date", "AGGREGATES") & updateRow)) _
                And Range(hFind("B/W Action", "FTA #" & i & " Date", "AGGREGATES") & updateRow).value _
                    = Lookup("BW_Action_Name")("Begin B/W") Then

            bucketHead = hFind("FTA #" & i & " Date", "AGGREGATES")
            Exit For
        End If
        
        If i = 1 Then
            MsgBox "Error: Can not find FTA instance to close"
            Exit Sub
        End If
    Next i
    
    Range(headerFind("Intake Conference Date", bucketHead) & updateRow) = IntakeDate
    Range(headerFind("Intake Conference Notes", bucketHead) & updateRow) = Notes
    
    
    If IntakeOutcomeRelease = True Then
        Range(headerFind("B/W Lifted Date", bucketHead) & updateRow).value = IntakeDate
        Range(headerFind("LOS B/W", bucketHead) & updateRow) _
            = calcLOS(Range(bucketHead & updateRow).value, IntakeDate.value)
            
        Range(headerFind("Intake Conference Notes", bucketHead) & updateRow).value _
            = Lookup("Intake_Conference_Outcome_Name")("Release for Court")
       
            
        If Not Supv1 = "None" Then
            Call addSupervision( _
                clientRow:=updateRow, _
                serviceType:=Supv1.value, _
                legalStatus:=Lookup("Legal_Status_Num")(Range(hFind("Legal Status") & updateRow).value), _
                Courtroom:=Lookup("Courtroom_Num")(Range(hFind("Active Courtroom") & updateRow).value), _
                CourtroomOfOrder:="Intake Conf. BW", _
                DA:="", _
                agency:=Supv1Pro.value, _
                startDate:=IntakeDate.value, _
                re1:=Supv1Re1.value, _
                re2:=Supv1Re2.value, _
                re3:=Supv1Re3.value, _
                Notes:="Referred at Bench Warrant Intake Conference ")
        End If
        
         If Not Supv2 = "None" Then
            Call addSupervision( _
                clientRow:=updateRow, _
                serviceType:=Supv2.value, _
                legalStatus:=Lookup("Legal_Status_Num")(Range(hFind("Legal Status") & updateRow).value), _
                Courtroom:=Lookup("Courtroom_Num")(Range(hFind("Active Courtroom") & updateRow).value), _
                CourtroomOfOrder:="Intake Conf. BW", _
                DA:="", _
                agency:=Supv2Pro.value, _
                startDate:=IntakeDate.value, _
                re1:=Supv2Re1.value, _
                re2:=Supv2Re2.value, _
                re3:=Supv2Re3.value, _
                Notes:="Referred at Bench Warrant Intake Conference ")
        End If
    End If

    Unload Me
End Sub

Private Sub UserForm_Click()

End Sub
