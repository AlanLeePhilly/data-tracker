Attribute VB_Name = "Helpers_JTC"
Sub JTC_Fetch()
    Dim head As Long
    Dim Num As Long
    Dim k As Variant
    
    'For Each k In Lookup.Keys
        'Debug.Print "Table: " & k
        'For Each J In Lookup(K).Keys
            'Debug.Print "Key: " & J & " Val: " & Lookup(K)(J)
        'Next J
    'Next k
    
    'clear all listboxes on form and modals
    With ClientUpdateForm
        courtHead = hFind("JTC")
        aggHead = hFind("AGGREGATES")
    
        Modal_JTC_Drop_Service.Service_Box.Clear
        Modal_JTC_Drop_Condition.Condition_Box.Clear
        .JTC_Fetch_Condition_Box.Clear
        .JTC_Return_Condition_Box.Clear
        .JTC_Fetch_Service_Box.Clear
        .JTC_Return_Service_Box.Clear
    
        'fetch first and last name from top of record
        .JTC_Fetch_First_Name.Caption = Range(headerFind("First Name") & updateRow).value
        .JTC_Fetch_Last_Name.Caption = Range(headerFind("Last Name") & updateRow).value
        
        'fetch phase number from column "Phase" after column "JTC"
        .JTC_Fetch_Phase.Caption = Lookup("JTC_Phase_Num")(Range(headerFind("Phase", headerFind("JTC")) & updateRow).value)
        
        'Cert Status
            If Range(headerFind("Was Notice of Certification Given?", aggHead) & updateRow).value = 2 Then '2 = "No"
                .JTC_Fetch_Certification = "None"
                 ClientUpdateForm.JTC_Certification_Update.Enabled = True
            Else
                .JTC_Fetch_Certification = _
                    Lookup("Result_of_Certification_Notice_Num") _
                    (Range(headerFind("Result of Certification Motion", aggHead) & updateRow).value)
                Call ClientUpdateForm.JTC_Certification_Remain_Click
                ClientUpdateForm.JTC_Certification_Update.Enabled = False
            End If
        
        .JTC_Fetch_Admission = Lookup("Generic_YNOU_Num")(Range(headerFind("Did Youth Enter an Admission?", aggHead) & updateRow).value)
        If .JTC_Fetch_Admission.Caption = "Yes" Then
            Call ClientUpdateForm.JTC_Admission_Remain_Click
            ClientUpdateForm.JTC_Admission_Update.Enabled = False
        Else
            ClientUpdateForm.JTC_Admission_Update.Enabled = True
        End If
        
        
        .JTC_Fetch_Adjudication = Lookup("Generic_YNOU_Num")(Range(headerFind("Adjudicated Delinquent?", aggHead) & updateRow).value)
        If .JTC_Fetch_Adjudication.Caption = "Yes" Then
            Call ClientUpdateForm.JTC_Adjudication_Remain_Click
            ClientUpdateForm.JTC_Adjudication_Update.Enabled = False
        Else
            ClientUpdateForm.JTC_Adjudication_Update.Enabled = True
        End If
        
        'set `phaseHead` string as column letters representing the banner column of current phase
        Select Case Lookup("JTC_Phase_Num")(Range(headerFind("Phase", headerFind("JTC")) & updateRow).value)
            Case "Referred"
                .JTC_Reject.Visible = True
                .JTC_Accept.Visible = True
                .JTC_Phase_Stepup.Visible = False
                .JTC_Phase_Pushback.Visible = False
                '.JTC_Phase_Remain.Visible = False
                .JTC_Discharge.Visible = False
                .JTC_Return_Stepup_Date_Label.Visible = False
                .JTC_Accept_Reject_Date_Label.Visible = False
                .JTC_Fetch_Stepup_Date_Label.Visible = False
                phaseHead = headerFind("PHASE 1", headerFind("JTC"))
            Case 1
                .JTC_Reject.Visible = False
                .JTC_Accept.Visible = False
                .JTC_Phase_Stepup.Visible = True
                .JTC_Phase_Pushback.Visible = True
                .JTC_Phase_Remain.Visible = True
                .JTC_Discharge.Visible = True
                phaseHead = headerFind("PHASE 1", headerFind("JTC"))
            Case 2
                .JTC_Reject.Visible = False
                .JTC_Accept.Visible = False
                .JTC_Phase_Stepup.Visible = True
                .JTC_Phase_Pushback.Visible = True
                .JTC_Phase_Remain.Visible = True
                .JTC_Discharge.Visible = True
                phaseHead = headerFind("PHASE 2", headerFind("JTC"))
            Case 3
                .JTC_Reject.Visible = False
                .JTC_Accept.Visible = False
                .JTC_Phase_Stepup.Enabled = False
                .JTC_Phase_Stepup.Visible = True
                .JTC_Phase_Pushback.Visible = True
                .JTC_Phase_Remain.Visible = True
                .JTC_Discharge.Visible = True
                phaseHead = headerFind("PHASE 3", headerFind("JTC"))
            Case "Graduated, Awaiting Expungment"
                .JTC_Reject.Visible = False
                .JTC_Accept.Visible = False
                .JTC_Fetch_Stepup_Date_Label.Visible = False
                .JTC_Fetch_Stepup_Date.Visible = False
                .JTC_Return_Stepup_Date.Visible = False
                .JTC_Phase_Stepup.Enabled = False
                .JTC_Phase_Pushback.Enabled = False
                .JTC_Phase_Remain.Enabled = False
                .JTC_Discharge.Enabled = False
                .JTC_Treatment_Stepdown.Enabled = False
                .JTC_Treatment_Provider_Update.Enabled = False
                .JTC_Treatment_Provider_Remain.Enabled = False
                .JTC_Treatment_Discharge.Enabled = False
                .JTC_Service_Add.Enabled = False
                .JTC_Service_Discharge.Enabled = False
                .JTC_Expungement.Visible = True
                phaseHead = headerFind("PHASE 3", headerFind("JTC"))
            Case "Graduated, Record Expunged"
                .JTC_Reject.Visible = False
                .JTC_Accept.Visible = False
                .JTC_Fetch_Stepup_Date_Label.Visible = False
                .JTC_Fetch_Stepup_Date.Visible = False
                .JTC_Return_Stepup_Date.Visible = False
                .JTC_Phase_Stepup.Enabled = False
                .JTC_Phase_Pushback.Enabled = False
                .JTC_Phase_Remain.Enabled = False
                .JTC_Discharge.Enabled = False
                .JTC_Treatment_Stepdown.Enabled = False
                .JTC_Treatment_Provider_Update.Enabled = False
                .JTC_Treatment_Provider_Remain.Enabled = False
                .JTC_Treatment_Discharge.Enabled = False
                .JTC_Service_Add.Enabled = False
                .JTC_Service_Discharge.Enabled = False
                .JTC_Expungement.Visible = True
                phaseHead = headerFind("PHASE 3", headerFind("JTC"))
            Case Else
                phaseHead = headerFind("PHASE 3", headerFind("JTC"))
        End Select
        
        'fetch current step-up date by checking the current phase
        'for the most recently posted push-back date or original scheduled step-up date
        With .JTC_Fetch_Stepup_Date
            Select Case False
                Case isEmptyOrZero(Range(headerFind("Push-Back Date #3", phaseHead) & updateRow))
                    .Caption = Range(headerFind("Push-Back Date #3", phaseHead) & updateRow)
                Case isEmptyOrZero(Range(headerFind("Push-Back Date #2", phaseHead) & updateRow))
                    .Caption = Range(headerFind("Push-Back Date #2", phaseHead) & updateRow)
                Case isEmptyOrZero(Range(headerFind("Push-Back Date #1", phaseHead) & updateRow))
                    .Caption = Range(headerFind("Push-Back Date #1", phaseHead) & updateRow)
                Case isEmptyOrZero(Range(headerFind("Scheduled Step-Up Date", phaseHead) & updateRow))
                    .Caption = Range(headerFind("Scheduled Step-Up Date", phaseHead) & updateRow)
            End Select
        End With
        
        If Range(hFind("Active B/W?") & updateRow).value = Lookup("Generic_YNOU_Name")("Yes") Then
            .JTC_Lift_BW.Enabled = True
        Else
            .JTC_Lift_BW.Enabled = False
        End If
        
        
        'fetch latest IOP provider from top of JTC section
        With .JTC_Fetch_Treatment_Provider
            'select the first case which resolves to false
            Select Case False
                'if that cell is empty, this expression is TRUE so we won't execute code listed after
                Case isEmptyOrZero(Range(hFind("IOP Provider #3", "JTC") & updateRow))
                    If isEmptyOrZero(Range(hFind("Discharge Date", "IOP Provider #3", "JTC") & updateRow)) Then
                        .Caption = Lookup("IOP_Provider_Num")(Range(hFind("IOP Provider #3", "JTC") & updateRow).value)
                    Else
                        .Caption = "Not currently assigned"
                    End If
                Case isEmptyOrZero(Range(hFind("IOP Provider #2", "JTC") & updateRow))
                    If isEmptyOrZero(Range(hFind("Discharge Date", "IOP Provider #2", "JTC") & updateRow)) Then
                        .Caption = Lookup("IOP_Provider_Num")(Range(hFind("IOP Provider #2", "JTC") & updateRow).value)
                    Else
                        .Caption = "Not currently assigned"
                    End If
                Case isEmptyOrZero(Range(hFind("IOP Provider #1", "JTC") & updateRow))
                    If isEmptyOrZero(Range(hFind("Discharge Date", "IOP Provider #1", "JTC") & updateRow)) Then
                        .Caption = Lookup("IOP_Provider_Num")(Range(hFind("IOP Provider #1", "JTC") & updateRow).value)
                    Else
                        .Caption = "Not currently assigned"
                    End If
                Case Else
                    .Caption = "Not currently assigned"
            End Select
        End With
        
        'GRAB AGG-ONLY Buckets
        For Num = 1 To 30
            bucketHead = hFind("Supervision Ordered #" & Num, "AGGREGATES")
            
            If Lookup("Courtroom_Num")(Range(headerFind("Courtroom of Order", bucketHead) & updateRow).value) = "Intake Conf." _
            Or Lookup("Courtroom_Num")(Range(headerFind("Courtroom of Order", bucketHead) & updateRow).value) = "PJJSC" Then
                If isEmptyOrZero(Range(headerFind("End Date", bucketHead) & updateRow)) Then
                    Call JTC_Service_Box_Add(ClientUpdateForm.JTC_Fetch_Service_Box, bucketHead)
                    Call JTC_Service_Box_Add(Modal_JTC_Drop_Service.Service_Box, bucketHead)
                    Call JTC_Service_Box_Add(ClientUpdateForm.JTC_Return_Service_Box, bucketHead)
                End If
            End If
            
            If Num <= 20 Then
                bucketHead = hFind("Condition Ordered #" & Num, "AGGREGATES")
                If Lookup("Courtroom_Num")(Range(headerFind("Courtroom of Order", bucketHead) & updateRow).value) = "Intake Conf." _
                Or Lookup("Courtroom_Num")(Range(headerFind("Courtroom of Order", bucketHead) & updateRow).value) = "PJJSC" Then
                    If isEmptyOrZero(Range(headerFind("End Date", bucketHead) & updateRow)) Then
                        Call JTC_Condition_Box_Add(ClientUpdateForm.JTC_Fetch_Condition_Box, bucketHead)
                        Call JTC_Condition_Box_Add(ClientUpdateForm.JTC_Return_Condition_Box, bucketHead)
                        Call JTC_Condition_Box_Add(Modal_JTC_Drop_Condition.Condition_Box, bucketHead)
                    End If
                End If
            End If
        Next Num
        
        'GRAB COURTROOM BUCKETS
        For Num = 1 To 30
            'if the Supervision Ordered bucket #num is not blank
            If isNotEmptyOrZero(Range(hFind("Supervision Ordered #" & Num, "Supervision Programs", "JTC") & updateRow)) Then
                If isEmptyOrZero(Range(hFind("End Date", "Supervision Ordered #" & Num, "Supervision Programs", "JTC") & updateRow)) Then
                    bucketHead = hFind("Supervision Ordered #" & Num, "Supervision Programs", "JTC")
                    Call JTC_Service_Box_Add(ClientUpdateForm.JTC_Fetch_Service_Box, bucketHead)
                    Call JTC_Service_Box_Add(Modal_JTC_Drop_Service.Service_Box, bucketHead)
                    Call JTC_Service_Box_Add(ClientUpdateForm.JTC_Return_Service_Box, bucketHead)
                End If
            End If
        Next Num
        
        For Num = 1 To 15
            If isNotEmptyOrZero(Range(hFind("Condition Ordered #" & Num, "Conditions", "JTC") & updateRow)) Then
                If isEmptyOrZero(Range(hFind("End Date", "Condition Ordered #" & Num, "Conditions", "JTC") & updateRow)) Then
                    bucketHead = hFind("Condition Ordered #" & Num, "Conditions", "JTC")
                    Call JTC_Condition_Box_Add(ClientUpdateForm.JTC_Fetch_Condition_Box, bucketHead)
                    Call JTC_Condition_Box_Add(ClientUpdateForm.JTC_Return_Condition_Box, bucketHead)
                    Call JTC_Condition_Box_Add(Modal_JTC_Drop_Condition.Condition_Box, bucketHead)
                End If
            End If
        Next Num
    End With
End Sub

Sub JTC_Service_Box_Add(ByRef MyBox As Object, ByVal bucketHead As String)
    Dim newIndex As Integer
    With MyBox
        .ColumnCount = 10
        .ColumnWidths = "90;75;75;75;0;0;0;0;0;0;"
            ' 0 Program                  6 Re1
            ' 1 Provider                 7 Re2
            ' 2 Start Date               8 Re3
            ' 3 End Date                 9 Notes
            ' 4 bucketHead or "New"
            ' 5 Nature

          .AddItem Lookup("JTC_Supervision_Status_Num")(Range(bucketHead & updateRow).value)
          newIndex = MyBox.ListCount - 1
        
          If isNotEmptyOrZero(Range(headerFind("Community-Based Agency", bucketHead) & updateRow)) Then
              .List(newIndex, 1) = _
                  Lookup("Community_Based_Supervision_Provider_Num")(Range(headerFind("Community-Based Agency", bucketHead) & updateRow).value)
          End If
          If isNotEmptyOrZero(Range(headerFind("Residential Agency", bucketHead) & updateRow)) Then
              .List(newIndex, 1) = _
                  Lookup("Residential_Supervision_Provider_Num")(Range(headerFind("Residential Agency", bucketHead) & updateRow).value) 'Res Agency
          End If
          
          .List(newIndex, 2) = Range(headerFind("Start Date", bucketHead) & updateRow).value
          '.List(newIndex, 3) = Range(headerFind("End Date", bucketHead) & updateRow).value
          .List(newIndex, 4) = bucketHead
    End With
End Sub

Sub JTC_Condition_Box_Add(ByRef MyBox As Object, ByVal bucketHead As String)
    Dim newIndex As Long
    With MyBox
        .ColumnCount = 10
        .ColumnWidths = "90;75;75;75;0;0;0;0;0;0;"
            ' 0 Program                  6 Re1
            ' 1 Provider                 7 Re2
            ' 2 Start Date               8 Re3
            ' 3 End Date                 9 Notes
            ' 4 bucketHead or "New"
            ' 5 Nature
            
        .AddItem Lookup("Condition_Num")(Range(bucketHead & updateRow).value)
        newIndex = MyBox.ListCount - 1
        .List(newIndex, 1) = _
            Lookup("Condition_Provider_Num")(Range(headerFind("Condition Agency", bucketHead) & updateRow).value)
        .List(newIndex, 2) = Range(headerFind("Start Date", bucketHead) & updateRow).value
        .List(newIndex, 3) = ""
        .List(newIndex, 4) = bucketHead
        
    End With
End Sub

