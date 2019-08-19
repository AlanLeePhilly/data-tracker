Attribute VB_Name = "Helpers_Standard"
Sub Standard_Fetch()
    With ClientUpdateForm
        If .Courtroom.value = "5E" Then
            courtHead = headerFind("Crossover")
        Else
            courtHead = headerFind(.Courtroom)
        End If

        aggHead = hFind("AGGREGATES")



        Modal_Standard_Drop_Condition.Condition_Box.Clear
        Modal_Standard_Drop_Supervision.Supervision_Box.Clear
        .Standard_Fetch_Condition_Box.Clear
        .Standard_Fetch_Supervision_Box.Clear
        .Standard_Return_Condition_Box.Clear
        .Standard_Return_Supervision_Box.Clear

        .Standard_Title = .Courtroom
        .Standard_Fetch_First_Name = Range(headerFind("First Name") & updateRow)
        .Standard_Fetch_Last_Name = Range(headerFind("Last Name") & updateRow)
        .Standard_Fetch_Legal_Status = Lookup("Legal_Status_Num")(Range(headerFind("Legal Status") & updateRow).value)




        'Cert Status
        If Range(headerFind("Was Notice of Certification Given?", aggHead) & updateRow).value = 2 Then '2 = "No"
            .Standard_Fetch_Certification = "None"
        Else
            .Standard_Fetch_Certification = _
                    Lookup("Result_of_Certification_Notice_Num") _
                    (Range(headerFind("Result of Certification Motion", aggHead) & updateRow).value)
            Call ClientUpdateForm.Standard_Certification_Remain_Click
            ClientUpdateForm.Standard_Certification_Update.Enabled = False
        End If


        .Standard_Fetch_Admission = Lookup("Generic_YNOU_Num")(Range(headerFind("Did Youth Enter an Admission?", aggHead) & updateRow).value)
        If .Standard_Fetch_Admission.Caption = "Yes" Then
            Call ClientUpdateForm.Standard_Admission_Remain_Click
            ClientUpdateForm.Standard_Admission_Update.Enabled = False
        End If

        .Standard_Fetch_Adjudication = Lookup("Generic_YNOU_Num")(Range(headerFind("Adjudicated Delinquent?", aggHead) & updateRow).value)
        If .Standard_Fetch_Adjudication.Caption = "Yes" Then
            Call ClientUpdateForm.Standard_Adjudication_Remain_Click
            ClientUpdateForm.Standard_Adjudication_Update.Enabled = False
        End If

        If Range(hFind("Active B/W?") & updateRow).value = Lookup("Generic_YNOU_Name")("Yes") Then
            .Standard_Lift_BW.Enabled = True
        Else
            .Standard_Lift_BW.Enabled = False
        End If

        'BOX LOADERS
        Dim Num As Long
        Dim bucketHead As String

        'GRAB AGG-ONLY Buckets
        For Num = 1 To 30
            bucketHead = hFind("Supervision Ordered #" & Num, "AGGREGATES")

            If Lookup("Courtroom_Num")(Range(headerFind("Courtroom of Order", bucketHead) & updateRow).value) = "Intake Conf." _
            Or Lookup("Courtroom_Num")(Range(headerFind("Courtroom of Order", bucketHead) & updateRow).value) = "PJJSC" Then
                If isEmptyOrZero(Range(headerFind("End Date", bucketHead) & updateRow)) Then
                    Call Standard_Supervision_Box_Add(ClientUpdateForm.Standard_Fetch_Supervision_Box, bucketHead)
                    Call Standard_Supervision_Box_Add(ClientUpdateForm.Standard_Return_Supervision_Box, bucketHead)
                    Call Standard_Supervision_Box_Add(Modal_Standard_Drop_Supervision.Supervision_Box, bucketHead)
                End If
            End If

            If Num <= 20 Then
                bucketHead = hFind("Condition Ordered #" & Num, "AGGREGATES")
                If Lookup("Courtroom_Num")(Range(headerFind("Courtroom of Order", bucketHead) & updateRow).value) = "Intake Conf." _
                Or Lookup("Courtroom_Num")(Range(headerFind("Courtroom of Order", bucketHead) & updateRow).value) = "PJJSC" Then
                    If isEmptyOrZero(Range(headerFind("End Date", bucketHead) & updateRow)) Then
                        Call Standard_Condition_Box_Add(ClientUpdateForm.Standard_Fetch_Condition_Box, bucketHead)
                        Call Standard_Condition_Box_Add(ClientUpdateForm.Standard_Return_Condition_Box, bucketHead)
                        Call Standard_Condition_Box_Add(Modal_Standard_Drop_Condition.Condition_Box, bucketHead)
                    End If
                End If
            End If
        Next Num

        'GRAB COURTROOM BUCKETS
        For Num = 1 To 15
            If isNotEmptyOrZero(Range(headerFind("Supervision Ordered #" & Num, courtHead) & updateRow)) And _
                isEmptyOrZero(Range(headerFind("End Date", headerFind("Supervision Ordered #" & Num, courtHead)) & updateRow)) Then
                'push the number value to the three functions which will add the Service bucket info to the three relevant tables
                'Fetch is list of services before hearing, Drop is the table in the "Drop Services" modal, and return is the final data about services after the hearing

                bucketHead = headerFind("Supervision Ordered #" & Num, courtHead)

                Call Standard_Supervision_Box_Add(ClientUpdateForm.Standard_Fetch_Supervision_Box, bucketHead)
                Call Standard_Supervision_Box_Add(ClientUpdateForm.Standard_Return_Supervision_Box, bucketHead)
                Call Standard_Supervision_Box_Add(Modal_Standard_Drop_Supervision.Supervision_Box, bucketHead)
            End If
            If isNotEmptyOrZero(Range(headerFind("Condition Ordered #" & Num, courtHead) & updateRow)) And _
                isEmptyOrZero(Range(headerFind("End Date", headerFind("Condition Ordered #" & Num, courtHead)) & updateRow)) Then
                'push the number value to the three functions which will add the Service bucket info to the three relevant tables
                'Fetch is list of services before hearing, Drop is the table in the "Drop Services" modal, and return is the final data about services after the hearing
                bucketHead = headerFind("Condition Ordered #" & Num, courtHead)
                Call Standard_Condition_Box_Add(ClientUpdateForm.Standard_Fetch_Condition_Box, bucketHead)
                Call Standard_Condition_Box_Add(ClientUpdateForm.Standard_Return_Condition_Box, bucketHead)
                Call Standard_Condition_Box_Add(Modal_Standard_Drop_Condition.Condition_Box, bucketHead)
            End If
        Next Num
    End With

    Call addPetitionsToBox(Modal_Standard_Adjudication.PetitionBox)
    Call addPetitionsToBox(Modal_Standard_Admission.PetitionBox)
End Sub

Sub Standard_Supervision_Box_Add(ByRef MyBox As Object, ByVal bucketHead As String)

    Dim newIndex As Integer

    With MyBox
        .ColumnCount = 10
        .ColumnWidths = "50;50;50;50;0;0;0;0;0;0;"
        ' 0 Program                  6 Re1
        ' 1 Provider                 7 Re2
        ' 2 Start Date               8 Re3
        ' 3 End Date                 9 Notes
        ' 4 bucketHead or "New"
        ' 5 Nature
        '

        ' set tempHead to first column in this 'service bucket'

        .AddItem Lookup("Supervision_Program_Num")(Range(bucketHead & updateRow).value)
        newIndex = MyBox.ListCount - 1

        If isResidential(Lookup("Supervision_Program_Num")(Range(bucketHead & updateRow).value)) Then
            .List(newIndex, 1) = _
                Lookup("Residential_Supervision_Provider_Num")(Range(headerFind("Residential Agency", bucketHead) & updateRow).value) 'Res Agency
        Else
            .List(newIndex, 1) = _
                Lookup("Community_Based_Supervision_Provider_Num")(Range(headerFind("Community-Based Agency", bucketHead) & updateRow).value)
        End If
        
        Dim dateCell As Range
        Set dateCell = Range(headerFind("Start Date", bucketHead) & updateRow)
        
        If Not IsDate(dateCell.value) And isNotEmptyOrZero(dateCell) Then
            dateCell.value = CDate(dateCell.value)
        End If
        
        .List(newIndex, 2) = dateCell.value
        '.List(newIndex, 3) = Range(headerFind("End Date", bucketHead) & updateRow).value
        .List(newIndex, 4) = bucketHead
    End With
End Sub

Sub Standard_Condition_Box_Add(ByRef MyBox As Object, ByVal bucketHead As String)

    Dim newIndex As Long

    With MyBox
        .ColumnCount = 10
        .ColumnWidths = "50;50;50;50;0;0;0;0;0;0;"
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

        Dim dateCell As Range
        Set dateCell = Range(headerFind("Start Date", bucketHead) & updateRow)
        
        If Not IsDate(dateCell.value) And isNotEmptyOrZero(dateCell) Then
            dateCell.value = CDate(dateCell.value)
        End If
        
        .List(newIndex, 2) = dateCell.value
        '.List(newIndex, 3) = Range(headerFind("End Date", bucketHead) & updateRow).value
        .List(newIndex, 4) = bucketHead
    End With
End Sub
