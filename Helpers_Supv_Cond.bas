Attribute VB_Name = "Helpers_Supv_Cond"

Sub addSupervision( _
    ByVal clientRow As Long, _
    ByVal serviceType As String, _
    ByVal legalStatus As String, _
    ByVal Courtroom As String, _
    ByVal DA As String, _
    ByVal agency As String, _
    ByVal startDate As String, _
    Optional ByVal re1 As String = "N/A", _
    Optional ByVal re2 As String = "N/A", _
    Optional ByVal re3 As String = "N/A", _
    Optional ByVal re4 As String = "N/A", _
    Optional ByVal re5 As String = "N/A", _
    Optional Notes As String = "", _
    Optional phase As String = "1", _
    Optional NextCourtDate As String, _
    Optional endDate As String = "", _
    Optional CourtroomOfOrder As String = "")
    'WORKS FOR STANDARD (+Cross & WRAP) AND JTC (+ AGG), ADD PHASE FOR LEGAL FOR JTC
    Worksheets("Entry").Activate

    Dim Num As Integer
    Dim bucketHead As String
    Dim section As String
    Dim i As Integer
    
    Call aggFlagSupervision( _
        userRow:=clientRow, _
        Courtroom:=Courtroom, _
        supervisionName:=serviceType)

    If serviceType = "Placement" Then
        Call startPlacement( _
            clientRow:=clientRow, _
            DA:=DA, _
            Courtroom:=Courtroom, _
            legalStatus:=legalStatus, _
            NextCourtDate:=NextCourtDate, _
            agency:=agency, _
            startDate:=startDate, _
            Notes:=Notes, _
            re1:=re1, _
            re2:=re2, _
            re3:=re3 _
        )
    End If

    For i = 1 To 2
        If Courtroom = "Intake Conf." _
        Or Courtroom = "Call-In" _
        Or Courtroom = "PJJSC" Then
            i = 2
            'if intake conf or PJJSC, only aggregate is needed
        End If

        Select Case i
            Case 1
                section = Courtroom
            Case 2
                section = "AGGREGATES"
        End Select

        Select Case section
            Case "4G", "4E", "6F", "6H", "3E", "5E", "WRAP"
                For Num = 1 To 15
                    If isEmptyOrZero(Range(hFind("Supervision Ordered #" & Num, section) & clientRow)) Then
                        bucketHead = hFind("Supervision Ordered #" & Num, section)
                        Num = 15
                    End If
                Next Num
            Case "JTC", "AGGREGATES"
                For Num = 1 To 30
                    If isEmptyOrZero(Range(hFind("Supervision Ordered #" & Num, section) & clientRow)) Then
                        bucketHead = hFind("Supervision Ordered #" & Num, section)
                        Num = 30
                    End If
                Next Num
            Case "Adult"
                For Num = 1 To 5
                    If isEmptyOrZero(Range(hFind("Supervision Ordered #" & Num, section) & clientRow)) Then
                        bucketHead = hFind("Supervision Ordered #" & Num, section)
                        Num = 5
                    End If
                Next Num
        End Select

        If Courtroom = "JTC" Then
            Range(bucketHead & clientRow).value = Lookup("Supervision_Program_Name")(serviceType)
        Else
            Range(bucketHead & clientRow).value = Lookup("Supervision_Program_Name")(serviceType)
        End If

        If section = "JTC" Then
            Select Case phase
                Case "1"
                    Range(headerFind("Phase of Order", bucketHead) & clientRow).value = Lookup("JTC_Phase_Name")(1)
                Case "2"
                    Range(headerFind("Phase of Order", bucketHead) & clientRow).value = Lookup("JTC_Phase_Name")(2)
                Case "3"
                    Range(headerFind("Phase of Order", bucketHead) & clientRow).value = Lookup("JTC_Phase_Name")(3)
                Case Else
                    Range(headerFind("Phase of Order", bucketHead) & clientRow).value = Lookup("JTC_Phase_Name")(phase)
            End Select
        Else
            Range(headerFind("Legal Status of Order", bucketHead) & clientRow).value = Lookup("Legal_Status_Name")(legalStatus)
        End If

        If CourtroomOfOrder = "" Then
            Range(headerFind("Courtroom of Order", bucketHead) & clientRow).value = Lookup("Courtroom_Name")(Courtroom)
        Else
            Range(headerFind("Courtroom of Order", bucketHead) & clientRow).value = Lookup("Courtroom_Name")(CourtroomOfOrder)
        End If
        Range(headerFind("DA", bucketHead) & clientRow).value = Lookup("DA_Last_Name_Name")(DA)
        If isResidential(serviceType) Then
            Range(headerFind("Residential Agency", bucketHead) & clientRow).value = Lookup("Residential_Supervision_Provider_Name")(agency)
        Else
            Range(headerFind("Community-Based Agency", bucketHead) & clientRow).value = Lookup("Community_Based_Supervision_Provider_Name")(agency)
        End If
        Range(headerFind("Start Date", bucketHead) & clientRow).value = startDate
        Range(headerFind("Reason #1 for Referral", bucketHead) & clientRow).value = Lookup("Supervision_Referral_Reason_Name")(re1)
        Range(headerFind("Reason #2 for Referral", bucketHead) & clientRow).value = Lookup("Supervision_Referral_Reason_Name")(re2)
        Range(headerFind("Reason #3 for Referral", bucketHead) & clientRow).value = Lookup("Supervision_Referral_Reason_Name")(re3)
        Range(headerFind("Reason #4 for Referral", bucketHead) & clientRow).value = Lookup("Supervision_Referral_Reason_Name")(re4)
        Range(headerFind("Reason #5 for Referral", bucketHead) & clientRow).value = Lookup("Supervision_Referral_Reason_Name")(re5)
        Range(headerFind("Supervision Description", bucketHead) & clientRow).value = Notes

        If Not endDate = "" Then
            Range(headerFind("End Date", bucketHead) & clientRow).value = endDate
            Range(headerFind("LOS", bucketHead) & clientRow).value = calcLOS(startDate, endDate)

        End If
    Next i

    'front active
    Range(headerFind("Active Supervision") & clientRow).value = Lookup("Supervision_Program_Name")(serviceType)
End Sub


Sub dropSupervision( _
    ByVal clientRow As Long, _
    ByVal serviceType As String, _
    ByVal Courtroom As String, _
    ByVal startDate As String, _
    ByVal endDate As String, _
    ByVal Nature As String, _
    Optional ByVal re1 As String = "N/A", _
    Optional ByVal re2 As String = "N/A", _
    Optional ByVal re3 As String = "N/A", _
    Optional ByVal re4 As String = "N/A", _
    Optional ByVal re5 As String = "N/A", _
    Optional Notes As String = "")

    Worksheets("Entry").Activate
    'WORKS FOR STANDARD (+Cross & WRAP) AND JTC (+ AGG)

    Dim bucketHead
    Dim section As String
    Dim i As Integer

    If serviceType = "Placement" Then
        Call endPlacement( _
            clientRow:=clientRow, _
            serviceType:=serviceType, _
            startDate:=startDate, _
            endDate:=endDate, _
            Nature:=Nature, _
            Notes:=Notes, _
            re1:=re1, _
            re2:=re2, _
            re3:=re3 _
        )
    End If

    For i = 1 To 2
        If Courtroom = "Intake Conf." _
        Or Courtroom = "Call-In" _
        Or Courtroom = "PJJSC" Then
            i = 2
            'if intake conf or PJJSC, only aggregate is needed
        End If

        Select Case i
            Case 1
                section = Courtroom
            Case 2
                section = "AGGREGATES"
        End Select
        
        Select Case section
            Case "4G", "4E", "6F", "6H", "3E", "5E", "WRAP"
                For Num = 1 To 15
                    bucketHead = hFind("Supervision Ordered #" & Num, section)
                    
                    If Range(bucketHead & clientRow) = Lookup("Supervision_Program_Name")(serviceType) _
                    And Range(headerFind("Start Date", bucketHead) & clientRow) = startDate Then
                        Num = 15
                    End If
                Next Num
            Case "JTC", "AGGREGATES"
                For Num = 1 To 30
                    bucketHead = hFind("Supervision Ordered #" & Num, section)
                    
                    If Range(bucketHead & clientRow) = Lookup("Supervision_Program_Name")(serviceType) _
                    And Range(headerFind("Start Date", bucketHead) & clientRow) = startDate Then
                        Num = 30
                    End If
                Next Num
            Case "Adult"
                For Num = 1 To 5
                    bucketHead = hFind("Supervision Ordered #" & Num, section)
                    
                    If Range(bucketHead & clientRow) = Lookup("Supervision_Program_Name")(serviceType) _
                    And Range(headerFind("Start Date", bucketHead) & clientRow) = startDate Then
                        Num = 5
                    End If
                Next Num
        End Select

        Range(headerFind("End Date", bucketHead) & clientRow).value = endDate
        Range(headerFind("Nature of Discharge", bucketHead) & clientRow) = Lookup("Nature_of_Discharge_Name")(Nature)

        If Not Nature = "Positive" Then
            Range(headerFind("Reason #1 for Negative D/C", bucketHead) & clientRow) = Lookup("Negative_Discharge_Reason_Name")(re1)
            Range(headerFind("Reason #2 for Negative D/C", bucketHead) & clientRow) = Lookup("Negative_Discharge_Reason_Name")(re2)
            Range(headerFind("Reason #3 for Negative D/C", bucketHead) & clientRow) = Lookup("Negative_Discharge_Reason_Name")(re3)
            Range(headerFind("Reason #4 for Negative D/C", bucketHead) & clientRow) = Lookup("Negative_Discharge_Reason_Name")(re4)
            Range(headerFind("Reason #5 for Negative D/C", bucketHead) & clientRow) = Lookup("Negative_Discharge_Reason_Name")(re5)
        End If

        Call append(Range(headerFind("Discharge Description", bucketHead) & clientRow), Notes)
        Range(headerFind("LOS", bucketHead) & clientRow) = calcLOS(startDate, endDate)

    Next i
End Sub

Sub addCondition( _
    ByVal clientRow As Long, _
    ByVal condition As String, _
    ByVal legalStatus As String, _
    ByVal Courtroom As String, _
    ByVal DA As String, _
    ByVal agency As String, _
    ByVal startDate As String, _
    Optional ByVal re1 As String = "N/A", _
    Optional ByVal re2 As String = "N/A", _
    Optional ByVal re3 As String = "N/A", _
    Optional ByVal re4 As String = "N/A", _
    Optional ByVal re5 As String = "N/A", _
    Optional Notes As String = "", _
    Optional phase As String = "1", _
    Optional CourtroomOfOrder As String = "" _
    )
    'WORKS FOR STANDARD (+Cross & WRAP) AND JTC (+ AGG), ADD PHASE FOR LEGAL FOR JTC
    Worksheets("Entry").Activate
    
    Dim Num As Integer
    Dim bucketHead As String
    Dim section As String
    Dim i As Integer
    
    Call aggFlagCondition( _
        userRow:=clientRow, _
        Courtroom:=Courtroom, _
        conditionName:=condition)
    
    For i = 1 To 2
        If Courtroom = "Intake Conf." _
        Or Courtroom = "Call-In" _
        Or Courtroom = "PJJSC" Then
            i = 2
            'if intake conf or PJJSC, only aggregate is needed
        End If

        Select Case i
            Case 1
                section = Courtroom
            Case 2
                section = "AGGREGATES"
        End Select


        Select Case section
            Case "4G", "4E", "6F", "6H", "3E", "JTC", "5E", "WRAP"
                For Num = 1 To 15
                    If isEmptyOrZero(Range(hFind("Condition Ordered #" & Num, section) & clientRow)) Then
                        bucketHead = hFind("Condition Ordered #" & Num, section)
                        Num = 15
                    End If
                Next Num
            Case "Adult"
                For Num = 1 To 5
                    If isEmptyOrZero(Range(hFind("Condition Ordered #" & Num, section) & clientRow)) Then
                        bucketHead = hFind("Condition Ordered #" & Num, section)
                        Num = 5
                    End If
                Next Num
            Case "AGGREGATES"
                For Num = 1 To 20
                    If isEmptyOrZero(Range(hFind("Condition Ordered #" & Num, section) & clientRow)) Then
                        bucketHead = hFind("Condition Ordered #" & Num, section)
                        Num = 20
                    End If
                Next Num
        End Select

        Range(bucketHead & clientRow).value = Lookup("Condition_Name")(condition)
        If section = "JTC" Then
            Range(headerFind("Phase of Order", bucketHead) & clientRow).value = Lookup("JTC_Phase_Name")(phase)
        Else
            Range(headerFind("Legal Status of Order", bucketHead) & clientRow).value = Lookup("Legal_Status_Name")(legalStatus)
        End If

        If CourtroomOfOrder = "" Then
            Range(headerFind("Courtroom of Order", bucketHead) & clientRow).value = Lookup("Courtroom_Name")(Courtroom)
        Else
            Range(headerFind("Courtroom of Order", bucketHead) & clientRow).value = Lookup("Courtroom_Name")(CourtroomOfOrder)
        End If

        Range(headerFind("DA", bucketHead) & clientRow).value = Lookup("DA_Last_Name_Name")(DA)

        Range(headerFind("Condition Agency", bucketHead) & clientRow).value = Lookup("Condition_Provider_Name")(agency)
        Range(headerFind("Start Date", bucketHead) & clientRow).value = startDate
        Range(headerFind("Reason #1 for Referral", bucketHead) & clientRow).value = Lookup("Supervision_Referral_Reason_Name")(re1)
        Range(headerFind("Reason #2 for Referral", bucketHead) & clientRow).value = Lookup("Supervision_Referral_Reason_Name")(re2)
        Range(headerFind("Reason #3 for Referral", bucketHead) & clientRow).value = Lookup("Supervision_Referral_Reason_Name")(re3)
        Range(headerFind("Reason #4 for Referral", bucketHead) & clientRow).value = Lookup("Supervision_Referral_Reason_Name")(re4)
        Range(headerFind("Reason #5 for Referral", bucketHead) & clientRow).value = Lookup("Supervision_Referral_Reason_Name")(re5)
        
        Range(headerFind("Condition Description", bucketHead) & clientRow).value = Notes
    Next i
End Sub

Sub dropCondition( _
    ByVal clientRow As Long, _
    ByVal Courtroom As String, _
    ByVal condition As String, _
    ByVal startDate As String, _
    ByVal endDate As String, _
    ByVal Nature As String, _
    Optional ByVal re1 As String = "N/A", _
    Optional ByVal re2 As String = "N/A", _
    Optional ByVal re3 As String = "N/A", _
    Optional ByVal re4 As String = "N/A", _
    Optional ByVal re5 As String = "N/A", _
    Optional Notes As String = "")

    Worksheets("Entry").Activate

    'WORKS FOR STANDARD (+Cross & WRAP) AND JTC (+ AGG)

    Dim bucketHead
    Dim section As String
    Dim i As Integer

    For i = 1 To 2
    
        If Courtroom = "Intake Conf." _
        Or Courtroom = "Call-In" _
        Or Courtroom = "PJJSC" Then
            i = 2
            'if intake conf or PJJSC, only aggregate is needed
        End If
            
        Select Case i
            Case 1
                section = Courtroom
            Case 2
                section = "AGGREGATES"
        End Select
        
        Select Case section
            Case "4G", "4E", "6F", "6H", "3E", "JTC", "5E", "WRAP"
                For Num = 1 To 15
                    bucketHead = hFind("Condition Ordered #" & Num, section)
                    
                    If Range(bucketHead & clientRow) = Lookup("Condition_Name")(condition) _
                    And Range(headerFind("Start Date", bucketHead) & clientRow) = startDate Then
                        Num = 15
                    End If
                Next Num
            Case "Adult"
                For Num = 1 To 5
                    bucketHead = hFind("Condition Ordered #" & Num, section)
                    
                    If Range(bucketHead & clientRow) = Lookup("Condition_Name")(condition) _
                    And Range(headerFind("Start Date", bucketHead) & clientRow) = startDate Then
                        Num = 5
                    End If
                Next Num
            Case "AGGREGATES"
                For Num = 1 To 20
                    bucketHead = hFind("Condition Ordered #" & Num, section)
                    
                    If Range(bucketHead & clientRow) = Lookup("Condition_Name")(condition) _
                    And Range(headerFind("Start Date", bucketHead) & clientRow) = startDate Then
                        Num = 20
                    End If
                Next Num
        End Select

        Range(headerFind("End Date", bucketHead) & clientRow).value = endDate
        Range(headerFind("Nature of Discharge", bucketHead) & clientRow) = Lookup("Nature_of_Discharge_Name")(Nature)

        If Not Nature = "Positive" Then
            Range(headerFind("Reason #1 for Negative D/C", bucketHead) & clientRow) = Lookup("Negative_Discharge_Reason_Name")(re1)
            Range(headerFind("Reason #2 for Negative D/C", bucketHead) & clientRow) = Lookup("Negative_Discharge_Reason_Name")(re2)
            Range(headerFind("Reason #3 for Negative D/C", bucketHead) & clientRow) = Lookup("Negative_Discharge_Reason_Name")(re3)
            Range(headerFind("Reason #4 for Negative D/C", bucketHead) & clientRow) = Lookup("Negative_Discharge_Reason_Name")(re4)
            Range(headerFind("Reason #5 for Negative D/C", bucketHead) & clientRow) = Lookup("Negative_Discharge_Reason_Name")(re5)
        End If

        Call append(Range(headerFind("Discharge Description", bucketHead) & clientRow), Notes)
        Range(headerFind("LOS", bucketHead) & clientRow) = calcLOS(startDate, endDate)
    Next i
End Sub


Sub closeIntakeConference(ByVal DateOf As String, ByVal userRow As Long)
    tempHead = hFind("INTAKE CONFERENCE")
    Dim i As Integer

    If Lookup("Generic_NYNOU_Num")(Range(headerFind("Did Youth Have an Intake Conference?", tempHead) & userRow).value) = "Yes" Then

        For i = 1 To 2
            tempHead = hFind("Supervision Ordered #" & i, "INTAKE CONFERENCE")
            If isNotEmptyOrZero(Range(tempHead & userRow)) Then
                If isEmptyOrZero(Range(headerFind("End Date", tempHead) & userRow)) Then
                    Range(headerFind("End Date", tempHead) & userRow).value = DateOf
                End If
                If isEmptyOrZero(Range(headerFind("LOS", tempHead) & userRow)) _
                And isNotEmptyOrZero(Range(hFind("Date of Intake Conference", "INTAKE CONFERENCE") & userRow)) Then
                    Range(headerFind("LOS", tempHead) & userRow).value _
                        = calcLOS(Range(hFind("Date of Intake Conference", "INTAKE CONFERENCE") & userRow).value, DateOf)
                End If
            End If
        Next i

        For i = 1 To 3
            tempHead = hFind("Other Condition #" & i, "INTAKE CONFERENCE")
            If isNotEmptyOrZero(Range(tempHead & userRow)) Then
                If isEmptyOrZero(Range(headerFind("End Date", tempHead) & userRow)) Then
                    Range(headerFind("End Date", tempHead) & userRow).value = DateOf
                End If
                If isEmptyOrZero(Range(headerFind("LOS", tempHead) & userRow)) _
                And isNotEmptyOrZero(Range(hFind("Date of Intake Conference", "INTAKE CONFERENCE") & userRow)) Then
                    Range(headerFind("LOS", tempHead) & userRow).value _
                        = calcLOS(Range(hFind("Date of Intake Conference", "INTAKE CONFERENCE") & userRow).value, DateOf)
                End If
            End If
        Next i

        If Lookup("Intake_Conference_Outcome_Num")(Range(hFind("Intake Conference Outcome", "INTAKE CONFERENCE") & userRow).value) = "Hold for Detention" Then
            If IsEmpty(Range(headerFind("LOS in Detention", tempHead) & userRow)) Then
                Range(headerFind("LOS in Detention", tempHead) & userRow).value _
                    = calcLOS(Range(hFind("Date of Intake Conference", "INTAKE CONFERENCE") & userRow).value, DateOf)
            End If

        Else
            If Lookup("Intake_Conference_Outcome_Num")(Range(hFind("Intake Conference Outcome", "INTAKE CONFERENCE") & userRow).value) = "Roll to Detention" Then
                If Lookup("DRAI_Action_Num")(Range(hFind("DRAI Action", "CALL-IN") & userRow).value) = "Follow - Hold" _
                Or Lookup("DRAI_Action_Num")(Range(hFind("DRAI Action", "CALL-IN") & userRow).value) = "Override - Hold" Then
                    Range(headerFind("LOS in Detention", tempHead) & userRow).value _
                        = calcLOS(Range(hFind("Date of Intake Conference", "INTAKE CONFERENCE") & userRow).value, DateOf)
                End If
            End If
        End If




        If IsEmpty(Range(headerFind("LOS Until Next Hearing", tempHead) & userRow)) Then
            Range(headerFind("LOS Until Next Hearing", tempHead) & userRow).value _
                = calcLOS(Range(hFind("Date of Intake Conference", "INTAKE CONFERENCE") & userRow).value, DateOf)
        End If
    End If
End Sub

Sub closeIntakeDetentions(ByVal DateOf As String, ByVal userRow As Long)
    Dim i As Integer
    Dim sectionHead As String, bucketHead As String

    sectionHead = hFind("Supervision Programs", "AGGREGATES")
    For i = 1 To 30
        bucketHead = headerFind("Supervision Ordered #" & i, sectionHead)
        If isNotEmptyOrZero(Range(bucketHead & userRow)) Then
            If IsEmpty(Range(headerFind("End Date", bucketHead) & userRow)) Then
                If Lookup("Supervision_Program_Num")(Range(bucketHead & userRow).value) = "Detention (not respite)" Then
                    If Lookup("Courtroom_Num")(Range(headerFind("Courtroom of Order", bucketHead) & userRow).value) = "Intake Conf." _
                    Or Lookup("Courtroom_Num")(Range(headerFind("Courtroom of Order", bucketHead) & userRow).value) = "Call-In." Then
                        Call dropSupervision( _
                            clientRow:=userRow, _
                            Courtroom:="PJJSC", _
                            serviceType:=Lookup("Supervision_Program_Num")(Range(bucketHead & userRow).value), _
                            startDate:=Range(headerFind("Start Date", bucketHead) & userRow).value, _
                            endDate:=DateOf, _
                            Nature:="Neutral", _
                            re1:="N/A", _
                            re2:="N/A", _
                            re3:="N/A", _
                            Notes:="Continued from Intake Conference.")
                    End If
                End If
            End If
        End If
    Next i
End Sub
Sub AggAggSupervisionsAndConditions(ByVal userRow As Long)
    Call AggAggSupervisions(userRow)
    Call AggAggConditions(userRow)
End Sub

Sub AggAggSupervisions(ByVal userRow As Long)
    Dim i As Integer
    Dim k As Integer
    Dim emptyBucketNum As Integer
    Dim closingBucketNum As Integer
    
    Dim aggHead As String
    Dim aggAggHead As String
    Dim aggBucketHead As String
    Dim aggBucketHead2 As String
    Dim aggAggBucketHead As String
    Dim closingBucketHead As String
    
    Dim isFirstInstance As Boolean
   
    Worksheets("Entry").Activate
    aggHead = hFind("Supervision Programs", "AGGREGATES")
    aggAggHead = hFind("Aggregate Supervision Programs", "AGGREGATES")
    Range(aggAggHead & userRow & ":" & headerFind("LOS", headerFind("Supervision Ordered #" & NUM_AGG_SUPERVISION_BUCKETS, aggAggHead)) & userRow).Clear
    
    
    'Crawl through Agg buckets
    For i = 1 To NUM_AGG_SUPERVISION_BUCKETS
        
        'Set head of bucket
        aggBucketHead = headerFind("Supervision Ordered #" & i, aggHead)
                'If the bucket has content
        If isNotEmptyOrZero(Range(aggBucketHead & userRow)) Then
            
            isFirstInstance = True
            
            'Crawl through Agg buckets until current
            For k = 1 To i
            
                'If reached current bucket
                If k = i Then
                    Exit For
                End If
                
                'Set head of bucket
                aggBucketHead2 = headerFind("Supervision Ordered #" & k, aggHead)
                
                'If matches w/ startDate/endDate and type, isFirst = false
                If Range(headerFind("Start Date", aggBucketHead) & userRow).value = Range(headerFind("End Date", aggBucketHead2) & userRow).value And _
                 Range(aggBucketHead & userRow).value = Range(aggBucketHead2 & userRow).value Then
                    isFirstInstance = False
                    Exit For
                End If
            Next k
            
            
            If isFirstInstance Then
                
                'Crawl AggAgg buckets
                For k = 1 To NUM_AGG_AGG_SUPERVISION_BUCKETS
                    
                    'Set head
                    aggAggBucketHead = headerFind("Supervision Ordered #" & k, aggAggHead)
                    
                    'IF Empty
                    If isEmptyOrZero(Range(aggAggBucketHead & userRow)) Then
                        emptyBucketNum = k
                        'Fill in opening of new bucket from agg bucket we're in
                        Range(aggAggBucketHead & userRow).value = Range(aggBucketHead & userRow).value
                        Range(headerFind("Legal Status of Order", aggAggBucketHead) & userRow).value = Range(headerFind("Legal Status of Order", aggBucketHead) & userRow).value
                        Range(headerFind("Courtroom of Order", aggAggBucketHead) & userRow).value = Range(headerFind("Courtroom of Order", aggBucketHead) & userRow).value
                        Range(headerFind("DA", aggAggBucketHead) & userRow).value = Range(headerFind("DA", aggBucketHead) & userRow).value
                        Range(headerFind("Community-Based Agency", aggAggBucketHead) & userRow).value = Range(headerFind("Community-Based Agency", aggBucketHead) & userRow).value
                        Range(headerFind("Residential Agency", aggAggBucketHead) & userRow).value = Range(headerFind("Residential Agency", aggBucketHead) & userRow).value
                        Range(headerFind("Start Date", aggAggBucketHead) & userRow).value = Range(headerFind("Start Date", aggBucketHead) & userRow).value
                        Range(headerFind("Reason #1 for Referral", aggAggBucketHead) & userRow).value = Range(headerFind("Reason #1 for Referral", aggBucketHead) & userRow).value
                        Range(headerFind("Reason #2 for Referral", aggAggBucketHead) & userRow).value = Range(headerFind("Reason #2 for Referral", aggBucketHead) & userRow).value
                        Range(headerFind("Reason #3 for Referral", aggAggBucketHead) & userRow).value = Range(headerFind("Reason #3 for Referral", aggBucketHead) & userRow).value
                        Range(headerFind("Reason #4 for Referral", aggAggBucketHead) & userRow).value = Range(headerFind("Reason #4 for Referral", aggBucketHead) & userRow).value
                        Range(headerFind("Reason #5 for Referral", aggAggBucketHead) & userRow).value = Range(headerFind("Reason #5 for Referral", aggBucketHead) & userRow).value
                        Range(headerFind("Supervision Description", aggAggBucketHead) & userRow).value = Range(headerFind("Supervision Description", aggBucketHead) & userRow).value
                        
                        Exit For
                    End If
                Next k
                
                closingBucketHead = headerFind("Supervision Ordered #" & i, aggHead)
                
                For k = i To NUM_AGG_SUPERVISION_BUCKETS
                
                    'Set head of bucket
                    aggBucketHead2 = headerFind("Supervision Ordered #" & k, aggHead)
                    
                    'each time that next bucket has matching start/endDate and type, attach new current bucket
                    If Range(headerFind("End Date", closingBucketHead) & userRow).value = Range(headerFind("Start Date", aggBucketHead2) & userRow).value And _
                     Range(closingBucketHead & userRow).value = Range(aggBucketHead2 & userRow).value Then
                        closingBucketHead = headerFind("Supervision Ordered #" & k, aggHead)
                        
                    End If
                Next k
                
                'Fill in closing bucket details with
                aggBucketHead = closingBucketHead
                
                Range(headerFind("End Date", aggAggBucketHead) & userRow).value = Range(headerFind("End Date", aggBucketHead) & userRow).value
                Range(headerFind("Nature of Discharge", aggAggBucketHead) & userRow).value = Range(headerFind("Nature of Discharge", aggBucketHead) & userRow).value
                Range(headerFind("Reason #1 for Negative D/C", aggAggBucketHead) & userRow).value = Range(headerFind("Reason #1 for Negative D/C", aggBucketHead) & userRow).value
                Range(headerFind("Reason #2 for Negative D/C", aggAggBucketHead) & userRow).value = Range(headerFind("Reason #2 for Negative D/C", aggBucketHead) & userRow).value
                Range(headerFind("Reason #3 for Negative D/C", aggAggBucketHead) & userRow).value = Range(headerFind("Reason #3 for Negative D/C", aggBucketHead) & userRow).value
                Range(headerFind("Reason #4 for Negative D/C", aggAggBucketHead) & userRow).value = Range(headerFind("Reason #4 for Negative D/C", aggBucketHead) & userRow).value
                Range(headerFind("Reason #5 for Negative D/C", aggAggBucketHead) & userRow).value = Range(headerFind("Reason #5 for Negative D/C", aggBucketHead) & userRow).value
                Range(headerFind("Discharge Description", aggAggBucketHead) & userRow).value = Range(headerFind("Discharge Description", aggBucketHead) & userRow).value
                Range(headerFind("LOS", aggAggBucketHead) & userRow).value = calcLOS(Range(headerFind("Start Date", aggAggBucketHead) & userRow).value, Range(headerFind("End Date", aggAggBucketHead) & userRow).value)
                
            End If
        End If
    Next i
End Sub

Sub AggAggConditions(ByVal userRow As Long)
    Dim i As Integer
    Dim k As Integer
    Dim emptyBucketNum As Integer
    Dim closingBucketNum As Integer
    
    Dim aggHead As String
    Dim aggAggHead As String
    Dim aggBucketHead As String
    Dim aggBucketHead2 As String
    Dim aggAggBucketHead As String
    Dim closingBucketHead As String
    
    Dim isFirstInstance As Boolean
   
    Worksheets("Entry").Activate
    aggHead = hFind("Conditions", "AGGREGATES")
    aggAggHead = hFind("Aggregate Conditions", "AGGREGATES")
    Range(aggAggHead & userRow & ":" & headerFind("LOS", headerFind("Condition Ordered #" & NUM_AGG_CONDITION_BUCKETS, aggAggHead)) & userRow).Clear
    
    
    'Crawl through Agg buckets
    For i = 1 To NUM_AGG_CONDITION_BUCKETS
        
        'Set head of bucket
        aggBucketHead = headerFind("Condition Ordered #" & i, aggHead)
                'If the bucket has content
        If isNotEmptyOrZero(Range(aggBucketHead & userRow)) Then
            
            isFirstInstance = True
            
            'Crawl through Agg buckets until current
            For k = 1 To i
            
                'If reached current bucket
                If k = i Then
                    Exit For
                End If
                
                'Set head of bucket
                aggBucketHead2 = headerFind("Condition Ordered #" & k, aggHead)
                
                'If matches w/ startDate/endDate and type, isFirst = false
                If Range(headerFind("Start Date", aggBucketHead) & userRow).value = Range(headerFind("End Date", aggBucketHead2) & userRow).value And _
                 Range(aggBucketHead & userRow).value = Range(aggBucketHead2 & userRow).value Then
                    isFirstInstance = False
                    Exit For
                End If
            Next k
            
            
            If isFirstInstance Then
                
                'Crawl AggAgg buckets
                For k = 1 To NUM_AGG_AGG_CONDITION_BUCKETS
                    
                    'Set head
                    aggAggBucketHead = headerFind("Condition Ordered #" & k, aggAggHead)
                    
                    'IF Empty
                    If isEmptyOrZero(Range(aggAggBucketHead & userRow)) Then
                        emptyBucketNum = k
                        'Fill in opening of new bucket from agg bucket we're in
                        Range(aggAggBucketHead & userRow).value = Range(aggBucketHead & userRow).value
                        Range(headerFind("Legal Status of Order", aggAggBucketHead) & userRow).value = Range(headerFind("Legal Status of Order", aggBucketHead) & userRow).value
                        Range(headerFind("Courtroom of Order", aggAggBucketHead) & userRow).value = Range(headerFind("Courtroom of Order", aggBucketHead) & userRow).value
                        Range(headerFind("DA", aggAggBucketHead) & userRow).value = Range(headerFind("DA", aggBucketHead) & userRow).value
                        Range(headerFind("Condition Agency", aggAggBucketHead) & userRow).value = Range(headerFind("Condition Agency", aggBucketHead) & userRow).value
                        Range(headerFind("Start Date", aggAggBucketHead) & userRow).value = Range(headerFind("Start Date", aggBucketHead) & userRow).value
                        Range(headerFind("Reason #1 for Referral", aggAggBucketHead) & userRow).value = Range(headerFind("Reason #1 for Referral", aggBucketHead) & userRow).value
                        Range(headerFind("Reason #2 for Referral", aggAggBucketHead) & userRow).value = Range(headerFind("Reason #2 for Referral", aggBucketHead) & userRow).value
                        Range(headerFind("Reason #3 for Referral", aggAggBucketHead) & userRow).value = Range(headerFind("Reason #3 for Referral", aggBucketHead) & userRow).value
                        Range(headerFind("Reason #4 for Referral", aggAggBucketHead) & userRow).value = Range(headerFind("Reason #4 for Referral", aggBucketHead) & userRow).value
                        Range(headerFind("Reason #5 for Referral", aggAggBucketHead) & userRow).value = Range(headerFind("Reason #5 for Referral", aggBucketHead) & userRow).value
                        Range(headerFind("Condition Description", aggAggBucketHead) & userRow).value = Range(headerFind("Condition Description", aggBucketHead) & userRow).value
                        
                        Exit For
                    End If
                Next k
                
                closingBucketHead = headerFind("Condition Ordered #" & i, aggHead)
                
                For k = i To NUM_AGG_CONDITION_BUCKETS
                
                    'Set head of bucket
                    aggBucketHead2 = headerFind("Condition Ordered #" & k, aggHead)
                    
                    'each time that next bucket has matching start/endDate and type, attach new current bucket
                    If Range(headerFind("End Date", closingBucketHead) & userRow).value = Range(headerFind("Start Date", aggBucketHead2) & userRow).value And _
                     Range(closingBucketHead & userRow).value = Range(aggBucketHead2 & userRow).value Then
                        closingBucketHead = headerFind("Condition Ordered #" & k, aggHead)
                        
                    End If
                Next k
                
                'Fill in closing bucket details with
                aggBucketHead = closingBucketHead
                
                Range(headerFind("End Date", aggAggBucketHead) & userRow).value = Range(headerFind("End Date", aggBucketHead) & userRow).value
                Range(headerFind("Nature of Discharge", aggAggBucketHead) & userRow).value = Range(headerFind("Nature of Discharge", aggBucketHead) & userRow).value
                Range(headerFind("Reason #1 for Negative D/C", aggAggBucketHead) & userRow).value = Range(headerFind("Reason #1 for Negative D/C", aggBucketHead) & userRow).value
                Range(headerFind("Reason #2 for Negative D/C", aggAggBucketHead) & userRow).value = Range(headerFind("Reason #2 for Negative D/C", aggBucketHead) & userRow).value
                Range(headerFind("Reason #3 for Negative D/C", aggAggBucketHead) & userRow).value = Range(headerFind("Reason #3 for Negative D/C", aggBucketHead) & userRow).value
                Range(headerFind("Reason #4 for Negative D/C", aggAggBucketHead) & userRow).value = Range(headerFind("Reason #4 for Negative D/C", aggBucketHead) & userRow).value
                Range(headerFind("Reason #5 for Negative D/C", aggAggBucketHead) & userRow).value = Range(headerFind("Reason #5 for Negative D/C", aggBucketHead) & userRow).value
                Range(headerFind("Discharge Description", aggAggBucketHead) & userRow).value = Range(headerFind("Discharge Description", aggBucketHead) & userRow).value
                Range(headerFind("LOS", aggAggBucketHead) & userRow).value = calcLOS(Range(headerFind("Start Date", aggAggBucketHead) & userRow).value, Range(headerFind("End Date", aggAggBucketHead) & userRow).value)
                
            End If
        End If
    Next i
End Sub



Sub aggFlagsConditionsSetNo(ByVal userRow As Long, bucketHead As String)
    Call flagNo(Range(headerFind("Was Youth Ordered Anger Mgt.?", bucketHead) & userRow))
    Call flagNo(Range(headerFind("Was Youth Ordered Alternative School?", bucketHead) & userRow))
    Call flagNo(Range(headerFind("Was Youth Ordered a BHE?", bucketHead) & userRow))
    Call flagNo(Range(headerFind("Was Youth Ordered Community Service?", bucketHead) & userRow))
    Call flagNo(Range(headerFind("Was Youth Ordered a Curfew?", bucketHead) & userRow))
    Call flagNo(Range(headerFind("Was Youth Ordered Daily School?", bucketHead) & userRow))
    Call flagNo(Range(headerFind("Was Youth Ordered Drug Screens?", bucketHead) & userRow))
    Call flagNo(Range(headerFind("Was Youth Ordered Essay?", bucketHead) & userRow))
    Call flagNo(Range(headerFind("Was Youth Ordered Family Therapy?", bucketHead) & userRow))
    Call flagNo(Range(headerFind("Was Youth Ordered GED?", bucketHead) & userRow))
    Call flagNo(Range(headerFind("Was Youth Ordered Grief Counseling?", bucketHead) & userRow))
    Call flagNo(Range(headerFind("Was Youth Ordered Mental Health?", bucketHead) & userRow))
    Call flagNo(Range(headerFind("Was Youth Ordered Inpatient D&A?", bucketHead) & userRow))
    Call flagNo(Range(headerFind("Was Youth Ordered IOP?", bucketHead) & userRow))
    Call flagNo(Range(headerFind("Was Youth Ordered Restitution?", bucketHead) & userRow))
    Call flagNo(Range(headerFind("Was Youth Ordered Sexual Counseling?", bucketHead) & userRow))
    Call flagNo(Range(headerFind("Was Youth Ordered Victim Conference?", bucketHead) & userRow))
    Call flagNo(Range(headerFind("Was Youth Ordered 1st Violation Hold?", bucketHead) & userRow))
    Call flagNo(Range(headerFind("Was Youth Ordered Other Condition?", bucketHead) & userRow))
End Sub

Sub aggFlagsSupervisionsSetNo(ByVal userRow As Long, bucketHead As String)
    Call flagNo(Range(headerFind("Did Youth Have IPS?", bucketHead) & userRow))
    Call flagNo(Range(headerFind("Did Youth Have Pre-ERC?", bucketHead) & userRow))
    Call flagNo(Range(headerFind("Did Youth Have IHD?", bucketHead) & userRow))
    Call flagNo(Range(headerFind("Did Youth Have ISP?", bucketHead) & userRow))
    Call flagNo(Range(headerFind("Did Youth Have GPS?", bucketHead) & userRow))
    Call flagNo(Range(headerFind("Did Youth Have Post-ERC?", bucketHead) & userRow))
    Call flagNo(Range(headerFind("Did Youth Have Detention?", bucketHead) & userRow))
    Call flagNo(Range(headerFind("Did Youth Have Reintegration?", bucketHead) & userRow))
    Call flagNo(Range(headerFind("Did Youth Have CUA?", bucketHead) & userRow))
    Call flagNo(Range(headerFind("Did Youth Have RTF?", bucketHead) & userRow))
    Call flagNo(Range(headerFind("Did Youth Have Inpatient D&A?", bucketHead) & userRow))
    Call flagNo(Range(headerFind("Did Youth Have Other Supervision?", bucketHead) & userRow))
End Sub

Sub aggFlagCondition( _
    ByVal userRow As Long, _
    ByVal Courtroom As String, _
    ByVal conditionName As String)
    
    Dim columnName As String
    
    Select Case conditionName
        Case "Anger Mgt."
            columnName = "Was Youth Ordered Anger Mgt.?"
        Case "Alt. School"
            columnName = "Was Youth Ordered Alternative School?"
        Case "BHE"
            columnName = "Was Youth Ordered a BHE?"
        Case "Comm. Serv"
            columnName = "Was Youth Ordered Community Service?"
        Case "Curfew"
            columnName = "Was Youth Ordered a Curfew?"
        Case "Daily School"
            columnName = "Was Youth Ordered Daily School?"
        Case "Drug Screens", "Drug Screen - FW"
            columnName = "Was Youth Ordered Drug Screens?"
        Case "Essay"
            columnName = "Was Youth Ordered Essay?"
        Case "Family Therapy"
            columnName = "Was Youth Ordered Family Therapy?"
        Case "GED"
            columnName = "Was Youth Ordered GED?"
        Case "Grief Counseling"
            columnName = "Was Youth Ordered Grief Counseling?"
        Case "Mental Health"
            columnName = "Was Youth Ordered Mental Health?"
        Case "Inpatient D&A"
            columnName = "Was Youth Ordered Inpatient D&A?"
        Case "IOP"
            columnName = "Was Youth Ordered IOP?"
        Case "Restitution"
            columnName = "Was Youth Ordered Restitution?"
        Case "Sexual Couns."
            columnName = "Was Youth Ordered Sexual Counseling?"
        Case "Victim Conf."
            columnName = "Was Youth Ordered Victim Conference?"
        Case "1st Viol. Hold"
            columnName = "Was Youth Ordered 1st Violation Hold?"
        Case Else
            columnName = "Was Youth Ordered Other Condition?"
    End Select
    
    Select Case Courtroom
        Case "4G", "4G", "4E", "6F", "6H", "3E", "Crossover", "WRAP", "JTC"
            Call flagYes(Range(hFind(columnName, Courtroom) & userRow))
        Case "Adult"
            Call flagYes(Range(hFind(columnName, "ADULT") & userRow))
    End Select

    Call flagYes(Range(hFind(columnName, "Conditions", "AGGREGATES") & userRow))
    Call flagYes(Range(hFind(columnName, "Aggregate Conditions", "AGGREGATES") & userRow))
End Sub

Sub aggFlagSupervision( _
    ByVal userRow As Long, _
    ByVal Courtroom As String, _
    ByVal supervisionName As String)
    
    Dim columnName As String

    Select Case supervisionName
        Case "IPS"
            columnName = "Did Youth Have IPS?"
        Case "Pre-ERC"
            columnName = "Did Youth Have Pre-ERC?"
        Case "IHD"
            columnName = "Did Youth Have IHD?"
        Case "ISP"
            columnName = "Did Youth Have ISP?"
        Case "GPS"
            columnName = "Did Youth Have GPS?"
        Case "Post-ERC"
            columnName = "Did Youth Have Post-ERC?"
        Case "Detention (respite)"
            columnName = "Did Youth Have Detention?"
        Case "Reintegration"
            columnName = "Did Youth Have Reintegration?"
        Case "CUA"
            columnName = "Did Youth Have CUA?"
        Case "RTF"
            columnName = "Did Youth Have RTF?"
        Case "Inpatient D&A"
            columnName = "Did Youth Have Inpatient D&A?"
        Case Else
            columnName = "Did Youth Have Other Supervision?"
    End Select
    
    Select Case Courtroom
        Case "4G", "4G", "4E", "6F", "6H", "3E", "Crossover", "WRAP", "JTC"
            Call flagYes(Range(hFind(columnName, Courtroom) & userRow))
        Case "Adult"
            Call flagYes(Range(hFind(columnName, "ADULT") & userRow))
    End Select

    Call flagYes(Range(hFind(columnName, "Supervision Programs", "AGGREGATES") & userRow))
    Call flagYes(Range(hFind(columnName, "Aggregate Supervision Programs", "AGGREGATES") & userRow))
End Sub

