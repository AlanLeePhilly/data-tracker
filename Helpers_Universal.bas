Attribute VB_Name = "Helpers_Universal"

Sub legalStatusStart( _
    ByVal clientRow As Long, _
    ByVal statusType As String, _
    ByVal Courtroom As String, _
    ByVal DA As String, _
    ByVal startDate As String, _
    Optional courtroomOfOrigin As String, _
    Optional Notes As String = "", _
    Optional zeroLocal As Boolean = False _
)
    Worksheets("Entry").Activate
    Dim bucketHead As String
    Dim canWriteLocal As Boolean
    Dim canWriteAgg As Boolean
    Dim i As Integer

    canWriteLocal = False
    canWriteAgg = False



    Select Case statusType
        Case "Pretrial", "Pretrial 2", "Consent Decree", "Interim Probation", "Probation", "Aftercare Probation"
            If isEmptyOrZero(Range(hFind("Start Date", statusType, "LEGAL STATUS", "AGGREGATES") & clientRow)) Then
                canWriteAgg = True
            End If
            Select Case Courtroom
                Case "4G", "4E", "6F", "6H", "3E"
                    If isEmptyOrZero(Range(hFind("Start Date", statusType, "Legal Status", Courtroom) & clientRow)) Then
                        canWriteLocal = True
                    End If
            End Select
    End Select

    Range(headerFind("Legal Status") & clientRow).value = Lookup("Legal_Status_Name")(statusType)

    For i = 1 To 2
        If canWriteAgg And i = 1 Then
            MsgBox "Opening Agg bucket for " & statusType
        Else
            i = 2
        End If

        Select Case i
            Case 1
                bucketHead = hFind(statusType, "LEGAL STATUS", "AGGREGATES")
            Case 2
                If canWriteLocal Then
                    MsgBox "Opening local bucket for " & statusType & " in " & Courtroom
                    bucketHead = hFind(statusType, "Legal Status", Courtroom)
                Else
                    Exit Sub
                End If
        End Select



        Range(headerFind("Was Youth on " & statusType & "?", bucketHead) & clientRow) _
            = Lookup("Generic_YNOU_Name")("Yes")

        If courtroomOfOrigin = "" Then
            Range(headerFind("Courtroom of Origin", bucketHead) & clientRow) _
                = Lookup("Courtroom_Name")(Courtroom)
        Else
            Range(headerFind("Courtroom of Origin", bucketHead) & clientRow) _
                = Lookup("Courtroom_Name")(courtroomOfOrigin)
        End If

        Range(headerFind("DA", bucketHead) & clientRow) _
            = Lookup("DA_Last_Name_Name")(DA)

        Range(headerFind("Age at Start of Status", bucketHead) & clientRow) _
            = ageAtTime(startDate, clientRow)
        Range(headerFind("Start Date", bucketHead) & clientRow) _
            = startDate
        Call append( _
            Range(headerFind("Notes on " & statusType, bucketHead) & clientRow), Notes)

    Next i
    
    If canWriteLocal And zeroLocal Then
        Call legalStatusEnd( _
            clientRow:=clientRow, _
            statusType:=statusType, _
            Courtroom:=Courtroom, _
            DA:=DA, _
            endDate:=startDate, _
            Nature:="Neutral", _
            withAgg:=False, _
            detailed:="Neutral Transfer of Status", _
            Notes:="Youth offered " & statusType & " transferred to new courtroom")
    End If

End Sub

Sub legalStatusEnd( _
    clientRow As Long, _
    ByVal statusType As String, _
    ByVal Courtroom As String, _
    DA As String, _
    endDate As String, _
    Nature As String, _
    detailed As String, _
    Optional Reason1 As String = "N/A", Optional Reason2 As String = "N/A", Optional Reason3 As String = "N/A", Optional Reason4 As String = "N/A", Optional Reason5 As String = "N/A", _
    Optional withAgg As Boolean = False, _
    Optional dischargingCourtroom As String = "", _
    Optional Notes As String = "" _
)
    Worksheets("Entry").Activate
    Dim bucketHead As String
    Dim canWriteLocal As Boolean
    Dim canWriteAgg As Boolean
    Dim i As Integer

    canWriteLocal = False
    canWriteAgg = False

    Select Case statusType
        Case "Pretrial", "Pretrial 2", "Consent Decree", "Interim Probation", "Probation", "Aftercare Probation"
            canWriteAgg = True

            Select Case Courtroom
                Case "4G", "4E", "6F", "6H", "3E"
                    If isEmptyOrZero(Range(headerFind("End Date", hFind(statusType, "Legal Status", Courtroom)) & clientRow)) Then
                        canWriteLocal = True
                    End If
            End Select
    End Select

    For i = 1 To 2
        If canWriteAgg And withAgg And i = 1 Then
            MsgBox "Closing Agg bucket for " & statusType
        Else
            i = 2
        End If

        Select Case i
            Case 1
                bucketHead = hFind(statusType, "LEGAL STATUS", "AGGREGATES")
            Case 2
                If canWriteLocal Then
                    MsgBox "Closing Local bucket for " & statusType & " in " & Courtroom
                    bucketHead = hFind(statusType, "Legal Status", Courtroom)
                Else
                    Exit Sub
                End If
        End Select


        If dischargingCourtroom = "" Then
            Range(headerFind("Discharging Courtroom", bucketHead) & clientRow) _
                = Lookup("Courtroom_Name")(Courtroom)
        Else
            Range(headerFind("Discharging Courtroom", bucketHead) & clientRow) _
                = Lookup("Courtroom_Name")(dischargingCourtroom)
        End If

        Range(headerFind("Discharging DA", bucketHead) & clientRow) _
            = Lookup("DA_Last_Name_Name")(DA)
        Range(headerFind("End Date", bucketHead) & clientRow) _
            = endDate
        Range(headerFind("Nature of Discharge", bucketHead) & clientRow) _
            = Lookup("Nature_of_Discharge_Name")(Nature)
        Range(headerFind("Detailed Status Outcome", bucketHead) & clientRow) _
            = Lookup("Detailed_Legal_Status_Outcome_Name")(detailed)

        Range(headerFind("Reason #1 for Negative Discharge", bucketHead) & clientRow) _
            = Lookup("Negative_Discharge_Reason_Name")(Reason1)
        Range(headerFind("Reason #2 for Negative Discharge", bucketHead) & clientRow) _
            = Lookup("Negative_Discharge_Reason_Name")(Reason2)
        Range(headerFind("Reason #3 for Negative Discharge", bucketHead) & clientRow) _
            = Lookup("Negative_Discharge_Reason_Name")(Reason3)
        Range(headerFind("Reason #4 for Negative Discharge", bucketHead) & clientRow) _
            = Lookup("Negative_Discharge_Reason_Name")(Reason4)
        Range(headerFind("Reason #5 for Negative Discharge", bucketHead) & clientRow) _
            = Lookup("Negative_Discharge_Reason_Name")(Reason5)
        Range(headerFind("LOS", bucketHead) & clientRow) _
            = calcLOS(Range(headerFind("Start Date", bucketHead) & clientRow).value, endDate)
        Call append( _
            Range(headerFind("Notes on " & statusType, bucketHead) & clientRow), Notes)

    Next i
End Sub

Sub closeOpenLegalStatuses( _
    clientRow As Long, _
    DateOf As String, _
    Courtroom As String, _
    DA As String, _
    legalStatus As String)

    Dim i As Integer, j As Integer
    Dim section As String, statusType As String
    Dim bucketHead As String

    For i = 1 To 6
        Select Case i
            Case 1
                section = "4G"
            Case 2
                section = "4E"
            Case 3
                section = "6F"
            Case 4
                section = "6H"
            Case 5
                section = "3E"
            Case 6
                section = "AGGREGATES"
        End Select

        For j = 1 To 5
            Select Case j
                Case 1
                    statusType = "Pretrial"
                Case 2
                    statusType = "Consent Decree"
                Case 3
                    statusType = "Interim Probation"
                Case 4
                    statusType = "Probation"
                Case 5
                    statusType = "Aftercare Probation"
            End Select

            'If statusType = legalStatus Then
            'If section = Courtroom Or section = "AGGREGATES" Then
            'GoTo NextJ
            'End If
            'End If

            bucketHead = hFind(statusType, section)

            If isNotEmptyOrZero(Range(headerFind("Start Date", bucketHead) & clientRow)) And _
               isEmptyOrZero(Range(headerFind("End Date", bucketHead) & clientRow)) Then

                MsgBox "Closing bucket for " & statusType & " in " & section & " (automated)"

                Range(headerFind("Discharging Courtroom", bucketHead) & clientRow) _
                    = Lookup("Courtroom_Name")(Courtroom)
                Range(headerFind("Discharging DA", bucketHead) & clientRow) _
                    = Lookup("DA_Last_Name_Name")(DA)
                Range(headerFind("End Date", bucketHead) & clientRow) _
                    = DateOf
                Range(headerFind("Nature of Discharge", bucketHead) & clientRow) _
                    = Lookup("Nature_of_Discharge_Name")("Neutral")
                Range(headerFind("Detailed Status Outcome", bucketHead) & clientRow) _
                    = Lookup("Detailed_Legal_Status_Outcome_Name")("N/A")

                Range(headerFind("Reason #1 for Negative Discharge", bucketHead) & clientRow) _
                    = Lookup("Negative_Discharge_Reason_Name")("N/A")
                Range(headerFind("Reason #2 for Negative Discharge", bucketHead) & clientRow) _
                    = Lookup("Negative_Discharge_Reason_Name")("N/A")
                Range(headerFind("Reason #3 for Negative Discharge", bucketHead) & clientRow) _
                    = Lookup("Negative_Discharge_Reason_Name")("N/A")
                Range(headerFind("Reason #4 for Negative Discharge", bucketHead) & clientRow) _
                    = Lookup("Negative_Discharge_Reason_Name")("N/A")
                Range(headerFind("Reason #5 for Negative Discharge", bucketHead) & clientRow) _
                    = Lookup("Negative_Discharge_Reason_Name")("N/A")
                Range(headerFind("LOS", bucketHead) & clientRow) _
                    = calcLOS(Range(headerFind("Start Date", bucketHead) & clientRow).value, DateOf)
                Call append( _
                    Range(headerFind("Notes on " & statusType, bucketHead) & clientRow), "This bucket was closed on a sweep by closeOpenLegalStatuses")

            End If
NextJ:
        Next j
    Next i
End Sub

Sub certificationStart( _
    ByVal clientRow As Long, _
    ByVal bucketHead As String, _
    ByVal statusType As String, _
    ByVal Courtroom As String, _
    ByVal DA As String, _
    ByVal startDate As String _
)
    Worksheets("Entry").Activate

    Range(headerFind("Was Notice of Certification Given?", bucketHead) & clientRow) _
        = Lookup("Generic_YNOU_Name")("Yes")
    Range(headerFind("Date of Certification Motion", bucketHead) & clientRow) _
        = startDate
    Range(headerFind("Courtroom of Certification Motion", bucketHead) & clientRow) _
        = Lookup("Courtroom_Name")(Courtroom)
    Range(headerFind("DA", bucketHead) & clientRow) _
        = Lookup("DA_Last_Name_Name")(DA)
    Range(headerFind("Legal Status of Certification Motion", bucketHead) & clientRow) _
        = Lookup("Legal_Status_Name")(statusType)
    Range(headerFind("Result of Certification Motion", bucketHead) & clientRow) _
        = Lookup("Certification_Status_Name")("Filed")
End Sub

Sub certificationUpdate( _
    ByVal clientRow As Long, _
    ByVal bucketHead As String, _
    ByVal Result As String, _
    ByVal newDate As String _
)
    Worksheets("Entry").Activate

    Range(headerFind("Result of Certification Motion", bucketHead) & clientRow) _
        = Lookup("Certification_Status_Name")(Result)
    Range(headerFind("LOS Certification Motion", bucketHead) & clientRow) _
        = calcLOS(Range(headerFind("Date of Certification Motion", bucketHead) & clientRow).value, newDate)
End Sub

Sub decertificationStart( _
    ByVal clientRow As Long, _
    ByVal bucketHead As String, _
    ByVal statusType As String, _
    ByVal Courtroom As String, _
    ByVal DA As String, _
    ByVal startDate As String _
)
    Worksheets("Entry").Activate

    Range(headerFind("Was Notice of De-Certification Given?", bucketHead) & clientRow) _
        = Lookup("Generic_YNOU_Name")("Yes")
    Range(headerFind("Date of De-Certification Motion", bucketHead) & clientRow) _
        = startDate
    Range(headerFind("Courtroom of De-Certification Motion", bucketHead) & clientRow) _
        = Lookup("Courtroom_Name")(Courtroom)
    Range(headerFind("DA", bucketHead) & clientRow) _
        = Lookup("DA_Last_Name_Name")(DA)
    Range(headerFind("Legal Status of De-Certification Motion", bucketHead) & clientRow) _
        = Lookup("Legal_Status_Name")(statusType)
    Range(headerFind("Result of De-Certification Motion", bucketHead) & clientRow) _
        = Lookup("Certification_Status_Name")("Filed")
End Sub

Sub decertificationUpdate( _
    ByVal clientRow As Long, _
    ByVal bucketHead As String, _
    ByVal Result As String, _
    ByVal newDate As String _
)
    Worksheets("Entry").Activate

    Range(headerFind("Result of De-Certification Motion", bucketHead) & clientRow) _
        = Lookup("Certification_Status_Name")(Result)
    Range(headerFind("LOS De-Certification Motion", bucketHead) & clientRow) _
        = calcLOS(newDate, Range(headerFind("Date of De-Certification Motion", bucketHead) & clientRow).value)
End Sub

Sub admissionStart( _
    ByVal clientRow As Long, _
    ByVal petitionNum As String, _
    ByVal statusType As String, _
    ByVal Courtroom As String, _
    ByVal DA As String, _
    ByVal startDate As String, _
    ByVal Result As String, _
    ByVal detailed As String, _
    leadChargeCode As String, _
    leadChargeName As String, _
    chargeCategory As String, _
    chargeGradeSpecific As String _
)
    Worksheets("Entry").Activate


    Dim bucketHead As String
    Dim petitionHead As String
    Dim i As Integer
    Dim j As Integer


    For i = 1 To 2
        Select Case i
            Case 1
                bucketHead = getCourtroomHead(Courtroom)
                bucketHead = headerFind("Admissions", bucketHead)
            Case 2
                bucketHead = hFind("Admissions", "COURT PROCEEDINGS", "AGGREGATES")
        End Select

        Range(headerFind("Did Youth Enter an Admission?", bucketHead) & clientRow) _
            = Lookup("Generic_YNOU_Name")("Yes")
        Range(headerFind("Date", bucketHead) & clientRow) _
            = startDate
        Range(headerFind("Courtroom", bucketHead) & clientRow) _
            = Lookup("Courtroom_Name")(Courtroom)
        Range(headerFind("DA", bucketHead) & clientRow) _
            = Lookup("DA_Last_Name_Name")(DA)
        Range(headerFind("Legal Status", bucketHead) & clientRow) _
            = Lookup("Legal_Status_Name")(statusType)
        Range(headerFind("Result", bucketHead) & clientRow) _
            = Lookup("Result_of_Admission_Name")(Result)
        Range(headerFind("Detailed Result", bucketHead) & clientRow) _
            = Lookup("Detailed_Result_of_Admission_Name")(detailed)

        Range(headerFind("Petition #", bucketHead) & clientRow) _
            = petitionNum

        Range(headerFind("Lead Charge Code", bucketHead) & clientRow) _
            = leadChargeCode
            
        Range(headerFind("Lead Charge Name", bucketHead) & clientRow) _
            = leadChargeName

        Range(headerFind("Charge Category", bucketHead) & clientRow) _
            = Lookup("Charge_Name")(chargeCategory)
            
        Range(headerFind("Charge Grade (specific)", bucketHead) & clientRow) _
            = Lookup("Charge_Grade_Specific_Name")(chargeGradeSpecific)

        Range(headerFind("Charge Grade (broad)", bucketHead) & clientRow) _
            = calcChargeBroad(chargeGradeSpecific)
            
        Range(headerFind("LOS Admission", bucketHead) & clientRow) _
            = calcLOS(Range(headerFind("Arrest Date") & clientRow).value, startDate)
    Next i
End Sub


Sub adjudicationStart( _
    ByVal clientRow As Long, _
    ByVal petitionNum As String, _
    ByVal Courtroom As String, _
    ByVal DA As String, _
    ByVal startDate As String, _
    ByVal Type_of As String, _
    re1 As String, re2 As String, re3 As String, re4 As String, re5 As String, _
    leadChargeCode As String, _
    leadChargeName As String, _
    chargeCategory As String, _
    chargeGradeSpecific As String _
)
    Worksheets("Entry").Activate

    Dim bucketHead As String
    Dim petitionHead As String
    Dim i As Integer
    Dim j As Integer

    For i = 1 To 2
        Select Case i
            Case 1
                Select Case Courtroom
                    Case "5E"
                        bucketHead = hFind("Adjudications", "COURT PROCEEDINGS", "Crossover")
                    Case "4G", "4E", "6F", "6H", "3E", "WRAP", "JTC"
                        bucketHead = hFind("Adjudications", "COURT PROCEEDINGS", Courtroom)
                End Select
            Case 2
                bucketHead = hFind("Adjudications", "COURT PROCEEDINGS", "AGGREGATES")
        End Select


        Range(headerFind("Adjudicated Delinquent?", bucketHead) & clientRow) _
            = Lookup("Generic_YNOU_Name")("Yes")
        Range(headerFind("Date of Adjudication", bucketHead) & clientRow) _
            = startDate
        Range(headerFind("Adjudicating Courtroom", bucketHead) & clientRow) _
            = Lookup("Courtroom_Name")(Courtroom)
        Range(headerFind("DA", bucketHead) & clientRow) _
            = Lookup("DA_Last_Name_Name")(DA)

        Range(headerFind("Guilty or Supervision Adjudication", bucketHead) & clientRow) _
            = Lookup("Reason_for_Adjudication_Name")(Type_of)
        Range(headerFind("Reason #1 for Technical Adjudication", bucketHead) & clientRow) _
            = Lookup("Negative_Discharge_Reason_Name")(re1)
        Range(headerFind("Reason #2 for Technical Adjudication", bucketHead) & clientRow) _
            = Lookup("Negative_Discharge_Reason_Name")(re2)
        Range(headerFind("Reason #3 for Technical Adjudication", bucketHead) & clientRow) _
            = Lookup("Negative_Discharge_Reason_Name")(re3)
        Range(headerFind("Reason #4 for Technical Adjudication", bucketHead) & clientRow) _
            = Lookup("Negative_Discharge_Reason_Name")(re4)
        Range(headerFind("Reason #5 for Technical Adjudication", bucketHead) & clientRow) _
            = Lookup("Negative_Discharge_Reason_Name")(re5)
            
        Range(headerFind("Petition #1", bucketHead) & clientRow) _
            = petitionNum

        Range(headerFind("Lead Charge Code", bucketHead) & clientRow) _
            = leadChargeCode
            
        Range(headerFind("Lead Charge Name", bucketHead) & clientRow) _
            = leadChargeName

        Range(headerFind("Charge Category", bucketHead) & clientRow) _
            = Lookup("Charge_Name")(chargeCategory)
            
        Range(headerFind("Charge Grade (specific)", bucketHead) & clientRow) _
            = Lookup("Charge_Grade_Specific_Name")(chargeGradeSpecific)

        Range(headerFind("Charge Grade (broad)", bucketHead) & clientRow) _
            = calcChargeBroad(chargeGradeSpecific)

        Range(headerFind("LOS Adjudication", bucketHead) & clientRow) _
            = calcLOS(Range(headerFind("Arrest Date") & clientRow).value, startDate)
    Next i
End Sub

Sub continuanceStart( _
    ByVal clientRow As Long, _
    ByVal listingStatus As String, _
    ByVal Courtroom As String, _
    ByVal DA As String, _
    ByVal startDate As String, _
    ByVal nextDate As String, _
    ByVal contType As String, _
    re1 As String, re2 As String, re3 As String _
)
    Worksheets("Entry").Activate

    ' Should work for: Standard, JTC, Adult (& AGG)

    Dim Num As Integer
    Dim i As Integer
    Dim section As String
    Dim bucketHead As String
    Dim tmpHead As String

    section = Courtroom

    Call flagYes(Range(hFind("Did Youth Have a Continuance?", "Continuances", Courtroom) & clientRow))
    Call flagYes(Range(hFind("Did Youth Have a Continuance?", "Continuances", "AGGREGATES") & clientRow))


    For Num = 1 To 2
        Select Case listingStatus
            Case "Pretrial"
                For i = 1 To 5
                    If isEmptyOrZero(Range(hFind("Date of Continuance #" & i, "PRETRIAL", "Continuances", section) & clientRow)) Then
                        bucketHead = hFind("Date of Continuance #" & i, "PRETRIAL", "Continuances", section)
                        i = 5
                    End If
                Next i
            Case "Adjudicatory"
                For i = 1 To 5
                    If isEmptyOrZero(Range(hFind("Date of Continuance #" & i, "ADJUDICATORY", "Continuances", section) & clientRow)) Then
                        bucketHead = hFind("Date of Continuance #" & i, "ADJUDICATORY", "Continuances", section)
                        i = 5
                    End If
                Next i
        End Select

        Range(bucketHead & clientRow) = startDate
        Range(headerFind("Next Court Date", bucketHead) & clientRow) = nextDate
        Range(headerFind("Courtroom of Continuance", bucketHead) & clientRow) _
            = Lookup("Courtroom_Name")(Courtroom)
        Range(headerFind("DA", bucketHead) & clientRow) _
            = Lookup("DA_Last_Name_Name")(DA)
        Range(headerFind("Listing Status of Continuance", bucketHead) & clientRow) _
            = Lookup("Continuance_Listing_Status_Name")(listingStatus)
        Range(headerFind("Type of Continuance", bucketHead) & clientRow) _
            = Lookup("Type_of_Continuance_Name")(contType)

        tmpHead = headerFind("Detailed Reason #1 for Commonwealth Continuance", bucketHead)
        Range(tmpHead & clientRow) = Lookup("Detailed_Reason_for_Commonwealth_Continuance_Name")(re1)
        Range(headerFind("Reason for Continuance Category", tmpHead) & clientRow) = commonwealthCat(re1)

        tmpHead = headerFind("Detailed Reason #2 for Commonwealth Continuance", bucketHead)
        Range(tmpHead & clientRow) = Lookup("Detailed_Reason_for_Commonwealth_Continuance_Name")(re2)
        Range(headerFind("Reason for Continuance Category", tmpHead) & clientRow) = commonwealthCat(re2)

        tmpHead = headerFind("Detailed Reason #3 for Commonwealth Continuance", bucketHead)
        Range(tmpHead & clientRow) = Lookup("Detailed_Reason_for_Commonwealth_Continuance_Name")(re3)
        Range(headerFind("Reason for Continuance Category", tmpHead) & clientRow) = commonwealthCat(re3)

        Range(headerFind("LOS Continuance", bucketHead) & clientRow) _
            = calcLOS(startDate, nextDate)

        section = "AGGREGATES"
    Next Num
End Sub

Sub loadPetitions(ByRef MyBox As Object, ByVal clientRow As String)
    Dim newIndex As Integer
    Dim Num As Integer
    Dim tmpHead As String

    Worksheets("Entry").Activate

    With MyBox
        .ColumnCount = 5
        .ColumnWidths = "0;30;50;80;50"

        For Num = 1 To 5
            tmpHead = hFind("Petition #" & Num, "PETITION")
            If isNotEmptyOrZero(Range(tmpHead & clientRow)) Then


                .AddItem tmpHead
                newIndex = MyBox.ListCount - 1

                .List(newIndex, 1) = Range(headerFind("Petition #" & Num, tmpHead) & clientRow).value
                .List(newIndex, 2) = Range(headerFind("Date Filed", tmpHead) & clientRow).value
                .List(newIndex, 3) = Range(headerFind("Lead Charge Name", tmpHead) & clientRow).value
            End If
        Next Num
    End With
End Sub

Sub addFTA( _
    ByVal clientRow As Long, _
    ByVal DateOf As String, _
    ByVal Courtroom As String, _
    ByVal legalStatus As String _
)
    Worksheets("Entry").Activate

    Dim i As Integer
    Dim fHead As String
    Dim bucketHead As String

    ftaHead = hFind("FTA", "AGGREGATES")
    For i = 1 To 15
        bucketHead = headerFind("FTA #" & i & " Date", ftaHead)
        If isEmptyOrZero(Range(bucketHead & clientRow)) Then
            Range(bucketHead & clientRow).value = DateOf
            Range(headerFind("Day of FTA", bucketHead) & clientRow).value = DateOf
            Range(headerFind("Courtroom", bucketHead) & clientRow).value = Lookup("Courtroom_Name")(Courtroom)
            Range(headerFind("Legal Status", bucketHead) & clientRow).value = Lookup("Legal_Status_Name")(legalStatus)
            If i = 1 Then
                Range(headerFind("LOS to FTA", bucketHead) & clientRow).value _
                 = calcLOS(Range(headerFind("Arrest Date") & clientRow).value, DateOf)
            Else
                Range(headerFind("LOS Between FTAs", bucketHead) & clientRow).value _
                 = calcLOS(Range(headerFind("FTA #" & (i - 1) & " Date", ftaHead) & clientRow).value, DateOf)
            End If
        End If
    Next i
End Sub


Sub startPlacement( _
    ByVal clientRow As Long, _
    ByVal DA As String, _
    ByVal Courtroom As String, _
    ByVal legalStatus As String, _
    ByVal NextCourtDate As String, _
    ByVal agency As String, _
    ByVal startDate As String, _
    Optional Notes As String = "", _
    Optional re1 As String, Optional re2 As String, Optional re3 As String _
)
    Dim bucketHead As String
    Dim courtHead As String
    Dim section As String
    Dim i As Integer


    'Hit section flag
    For i = 1 To 2
        Select Case i
            Case 1
                section = Courtroom
                If section = "5E" Then
                    section = "Crossover"
                End If
            Case 2
                section = "AGGREGATES"
        End Select

        Call flagYes(Range(hFind("Was Youth Placed?", section) & clientRow))

        Select Case section
            Case "4G", "4E", "6F", "6H", "3E", "Crossover", "WRAP"



                For Num = 1 To 5
                    If isEmptyOrZero(Range(hFind("Legal Status Prior to Commit #" & Num, section) & clientRow)) Then
                        bucketHead = hFind("Legal Status Prior to Commit #" & Num, section)
                        Num = 5
                    End If
                Next Num

            Case "JTC"
                For Num = 1 To 5
                    If isEmptyOrZero(Range(hFind("Placement #" & Num & " Phase", section) & clientRow)) Then
                        bucketHead = hFind("Placement #" & Num & " Phase", section)
                        Num = 5
                    Else
                        If Num = 5 Then
                            err.Raise vbObjectError + 1101, "AddPlacement", "Tried to add a 6th placement to JTC"
                        End If
                    End If
                Next Num

            Case "AGGREGATES"
                For Num = 1 To 10
                    If isEmptyOrZero(Range(hFind("Legal Status Prior to Commit #" & Num, section) & clientRow)) Then
                        bucketHead = hFind("Legal Status Prior to Commit #" & Num, section)
                        Num = 10
                    End If
                Next Num
        End Select

        If section = "JTC" Then
            Range(bucketHead & clientRow).value = Range(hFind("Phase", "JTC") & clientRow).value
        Else
            Range(bucketHead & clientRow).value = Lookup("Legal_Status_Name")(legalStatus)
        End If

        Range(headerFind("Committing Courtroom", bucketHead) & clientRow).value = Lookup("Courtroom_Name")(Courtroom)
        Range(headerFind("DA", bucketHead) & clientRow).value = Lookup("DA_Last_Name_Name")(DA)

        Range(headerFind("LOS Original Order", bucketHead) & clientRow).value = calcLOS(startDate, NextCourtDate)
        Range(headerFind("Residential Agency", bucketHead) & clientRow).value = Lookup("Residential_Supervision_Provider_Name")(agency)
        Range(headerFind("Start Date", bucketHead) & clientRow).value = startDate
        Range(headerFind("Reason #1 for Placement", bucketHead) & clientRow).value = Lookup("Supervision_Referral_Reason_Name")(re1)
        Range(headerFind("Reason #2 for Placement", bucketHead) & clientRow).value = Lookup("Supervision_Referral_Reason_Name")(re2)
        Range(headerFind("Reason #3 for Placement", bucketHead) & clientRow).value = Lookup("Supervision_Referral_Reason_Name")(re3)
        Range(headerFind("Placement Description", bucketHead) & clientRow).value = Notes
    Next i

    Range(headerFind("Active Supervision") & clientRow).value = Lookup("Supervision_Program_Name")("Placement")
    Range(headerFind("Active Supervision Provider") & clientRow).value = Lookup("Residential_Supervision_Provider_Name")(agency)

End Sub

Sub endPlacement( _
    ByVal clientRow As Long, _
    ByVal serviceType As String, _
    ByVal startDate As String, _
    ByVal endDate As String, _
    ByVal Nature As String, _
    Optional re1 As String, Optional re2 As String, Optional re3 As String, _
    Optional Notes As String = "")

    Worksheets("Entry").Activate
    'WORKS FOR STANDARD (+Cross & WRAP) AND JTC (+ AGG)

    Dim bucketHead
    Dim section As String
    Dim i As Integer
    Dim Courtroom As String

    For i = 1 To 2
        Select Case i
            Case 1
                section = "AGGREGATES"
            Case 2
                section = Courtroom
        End Select

        Select Case section

            Case "AGGREGATES"
                For Num = 1 To 10
                    bucketHead = hFind("Legal Status Prior to Commit #" & Num, "AGGREGATES")

                    If Range(headerFind("Start Date", bucketHead) & clientRow) = startDate Then
                        Num = 10
                        Courtroom = Lookup("Courtroom_Num")(Range(headerFind("Committing Courtroom", bucketHead) & clientRow).value)
                    End If
                Next Num

            Case "4G", "4E", "6F", "6H", "3E", "Crossover", "WRAP"
                For Num = 1 To 5
                    bucketHead = hFind("Legal Status Prior to Commit #" & Num, section)
                    If Range(headerFind("Start Date", bucketHead) & clientRow).value = startDate Then
                        Num = 5
                    End If
                Next Num

            Case "JTC"
                For Num = 1 To 5
                    bucketHead = hFind("Placement #" & Num & " Phase", section)
                    If Range(headerFind("Start Date", bucketHead) & clientRow).value = startDate Then
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
        End If

        Call append(Range(headerFind("Discharge Description", bucketHead) & clientRow), Notes)
        Range(headerFind("LOS", bucketHead) & clientRow) = calcLOS(startDate, endDate)

    Next i
End Sub

Sub closeCallIn(ByVal DateOf As String, ByVal userRow As Long)
    tempHead = hFind("CALL-IN")

    If Lookup("Generic_NYNOU_Num")(Range(headerFind("Did Youth Have Call-In?", tempHead) & userRow).value) = "Yes" Then

        If Lookup("DRAI_Action_Num")(Range(headerFind("DRAI Action", tempHead) & userRow).value) = "Follow - Hold" _
        Or Lookup("DRAI_Action_Num")(Range(headerFind("DRAI Action", tempHead) & userRow).value) = "Override - Hold" Then
            If isEmptyOrZero(Range(headerFind("End Date", tempHead) & userRow)) Then
                Range(headerFind("End Date", tempHead) & userRow).value = DateOf
            End If
            If isEmptyOrZero(Range(headerFind("LOS in Detention", tempHead) & userRow)) _
                And isNotEmptyOrZero(Range(hFind("Date of Call-In", "CALL-IN") & userRow)) Then
                Range(headerFind("LOS in Detention", tempHead) & userRow).value _
                    = calcLOS(Range(headerFind("Date of Call-In", tempHead) & userRow).value, Range(headerFind("End Date", tempHead) & userRow).value)
            End If
        End If
    End If
End Sub




