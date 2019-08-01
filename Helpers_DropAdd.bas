Attribute VB_Name = "Helpers_DropAdd"

Sub addSupervision( _
    ByVal clientRow As Long, _
    ByVal serviceType As String, _
    ByVal legalStatus As String, _
    ByVal Courtroom As String, _
    ByVal DA As String, _
    ByVal agency As String, _
    ByVal startDate As String, _
    Optional Re1 As String, Optional Re2 As String, Optional Re3 As String, _
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
            Re1:=Re1, _
            Re2:=Re2, _
            Re3:=Re3 _
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
        Range(headerFind("Reason #1 for Referral", bucketHead) & clientRow).value = Lookup("Supervision_Referral_Reason_Name")(Re1)
        Range(headerFind("Reason #2 for Referral", bucketHead) & clientRow).value = Lookup("Supervision_Referral_Reason_Name")(Re2)
        Range(headerFind("Reason #3 for Referral", bucketHead) & clientRow).value = Lookup("Supervision_Referral_Reason_Name")(Re3)
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
    ByVal head As String, _
    ByVal serviceType As String, _
    ByVal startDate As String, _
    ByVal endDate As String, _
    ByVal Nature As String, _
    Optional Re1 As String, Optional Re2 As String, Optional Re3 As String, _
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
            Re1:=Re1, _
            Re2:=Re2, _
            Re3:=Re3 _
        )
    End If

    For i = 1 To 2


        Select Case i
            Case 1
                bucketHead = head
            Case 2
                section = "AGGREGATES"
        End Select

        Select Case section

            Case "AGGREGATES"
                For Num = 1 To 30
                    bucketHead = hFind("Supervision Ordered #" & Num, "AGGREGATES")

                    If Range(bucketHead & clientRow) = Lookup("Supervision_Program_Name")(serviceType) _
                    And Range(headerFind("Start Date", bucketHead) & clientRow) = startDate Then
                        Num = 30
                    End If
                Next Num
        End Select

        Range(headerFind("End Date", bucketHead) & clientRow).value = endDate
        Range(headerFind("Nature of Discharge", bucketHead) & clientRow) = Lookup("Nature_of_Discharge_Name")(Nature)

        If Not Nature = "Positive" Then
            Range(headerFind("Reason #1 for Negative D/C", bucketHead) & clientRow) = Lookup("Negative_Discharge_Reason_Name")(Re1)
            Range(headerFind("Reason #2 for Negative D/C", bucketHead) & clientRow) = Lookup("Negative_Discharge_Reason_Name")(Re2)
            Range(headerFind("Reason #3 for Negative D/C", bucketHead) & clientRow) = Lookup("Negative_Discharge_Reason_Name")(Re3)
        End If

        Call append(Range(headerFind("Discharge Description", bucketHead) & clientRow), Notes)
        Range(headerFind("LOS", bucketHead) & clientRow) = calcLOS(startDate, endDate)

        If alphaToNum(head) > alphaToNum(hFind("AGGREGATES")) Then
            i = 2
        End If


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
    Optional Re1 As String, Optional Re2 As String, Optional Re3 As String, _
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

    For i = 1 To 2
        If Courtroom = "Intake Conf." _
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
        Range(headerFind("Reason #1 for Referral", bucketHead) & clientRow).value = Lookup("Supervision_Referral_Reason_Name")(Re1)
        Range(headerFind("Reason #2 for Referral", bucketHead) & clientRow).value = Lookup("Supervision_Referral_Reason_Name")(Re2)
        Range(headerFind("Reason #3 for Referral", bucketHead) & clientRow).value = Lookup("Supervision_Referral_Reason_Name")(Re3)
        Range(headerFind("Condition Description", bucketHead) & clientRow).value = Notes
    Next i
End Sub

Sub dropCondition( _
    ByVal clientRow As Long, _
    ByVal head As String, _
    ByVal condition As String, _
    ByVal startDate As String, _
    ByVal endDate As String, _
    ByVal Nature As String, _
    Optional Re1 As Variant, Optional Re2 As Variant, Optional Re3 As Variant, _
    Optional Notes As String = "")

    Worksheets("Entry").Activate

    'WORKS FOR STANDARD (+Cross & WRAP) AND JTC (+ AGG)

    Dim bucketHead
    Dim section As String
    Dim i As Integer

    For i = 1 To 2
        Select Case i
            Case 1
                bucketHead = head
            Case 2
                section = "AGGREGATES"
        End Select

        Select Case section

            Case "AGGREGATES"
                For Num = 1 To 20
                    bucketHead = hFind("Condition Ordered #" & Num, "AGGREGATES")

                    If Range(bucketHead & clientRow) = Lookup("Condition_Name")(condition) _
                    And Range(headerFind("Start Date", bucketHead) & clientRow) = startDate Then
                        Num = 20
                    End If
                Next Num
        End Select

        Range(headerFind("End Date", bucketHead) & clientRow).value = endDate
        Range(headerFind("Nature of Discharge", bucketHead) & clientRow) = Lookup("Nature_of_Discharge_Name")(Nature)

        If Not Nature = "Positive" Then
            Range(headerFind("Reason #1 for Negative D/C", bucketHead) & clientRow) = Lookup("Negative_Discharge_Reason_Name")(Re1)
            Range(headerFind("Reason #2 for Negative D/C", bucketHead) & clientRow) = Lookup("Negative_Discharge_Reason_Name")(Re2)
            Range(headerFind("Reason #3 for Negative D/C", bucketHead) & clientRow) = Lookup("Negative_Discharge_Reason_Name")(Re3)
        End If

        Call append(Range(headerFind("Discharge Description", bucketHead) & clientRow), Notes)
        Range(headerFind("LOS", bucketHead) & clientRow) = calcLOS(startDate, endDate)

        If alphaToNum(head) > alphaToNum(hFind("AGGREGATES")) Then
            i = 2
        End If
    Next i
End Sub


