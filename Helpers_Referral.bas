Attribute VB_Name = "Helpers_Referral"
Option Explicit
Sub ReferClientFrom( _
    fromCR As String, _
    referralDate As String, _
    clientRow As Long, _
    Optional toCR As String = "N/A", _
    Optional Notes As String = "")
    Dim toHead As String
    Dim fromHead As String
    
    Select Case fromCR
        Case "4G", "4E", "6F", "6H", "3E", "JTC", "WRAP", "Adult"
            fromHead = headerFind(fromCR)
        Case Else
            fromHead = "A"
    End Select
    
    Select Case toCR
        Case "4G", "4E", "6F", "6H", "3E", "JTC", "WRAP", "Adult"
            toHead = headerFind(toCR)
        Case Else
            toHead = "A"
    End Select
    
        Select Case fromCR
        Case "4G", "4E", "6F", "6H", "3E", "WRAP", "5E"
            Range(headerFind("End Date", fromHead) & clientRow).value _
                = referralDate
            Range(headerFind("LOS", fromHead) & clientRow).value _
                = DateDiff("d", _
                    Range(headerFind("Start Date", fromHead) & clientRow).value, _
                    Range(headerFind("End Date", fromHead) & clientRow).value)
            Range(headerFind("Courtroom of Transfer (if relevant)", fromHead) & clientRow).value _
                = Lookup("Courtroom_Name")(toCR)
        Case "JTC"
        'check phase and pull out
            
        Case "Adult"
            Range(headerFind("End Date", fromHead) & clientRow).value _
                = referralDate
            Range(headerFind("LOS", fromHead) & clientRow).value _
                = DateDiff("d", _
                    Range(headerFind("Start Date", fromHead) & clientRow).value, _
                    Range(headerFind("End Date", fromHead) & clientRow).value)
                    
        
        Case Else
            MsgBox "Referring Courtroom not recognized. Contact your admin."
            Exit Sub
    End Select

End Sub



Sub ReferClientTo( _
    referralDate As String, _
    clientRow As Long, _
    Optional toCR As String = "N/A", _
    Optional fromCR As String = "N/A", _
    Optional Notes As String = "", _
    Optional newLegalStatus As String = "")
    
    Worksheets("Entry").Activate
    
    Dim toHead As String
    Dim fromHead As String
    Dim legalStatusVar As String
    
                                '''''''''''
                                'SET HEADS'
                                '''''''''''
    Select Case fromCR
        Case "4G", "4E", "6F", "6H", "3E", "JTC", "WRAP", "Adult"
            fromHead = headerFind(fromCR)
        Case "PJJSC"
            fromHead = headerFind("DETENTION")
        Case "5E"
            fromHead = headerFind("Crossover")
        Case "N/A"
            fromHead = "A"
        Case Else
            MsgBox ("Courtroom " & fromCR & " not recognized. Contact your admin")
            
    End Select
    
    Select Case toCR
        Case "4G", "4E", "6F", "6H", "3E", "JTC", "WRAP", "Adult"
            toHead = headerFind(toCR)
        Case "PJJSC"
            toHead = headerFind("DETENTION")
        Case "5E"
            toHead = headerFind("Crossover")
        Case "N/A"
            toHead = "A"
        Case Else
            MsgBox ("Courtroom " & toCR & " not recognized. Contact your admin")
            
    End Select
    
                                ''''''''''''''''
                                'CLOSE OLD ROOM'
                                ''''''''''''''''
    If Not fromCR = "N/A" Then
        Select Case fromCR
            Case "4G", "4E", "6F", "6H", "3E"
                Range(headerFind("End Date", fromHead) & clientRow).value _
                    = referralDate
                Range(headerFind("LOS", fromHead) & clientRow).value _
                    = DateDiff("d", _
                        Range(headerFind("Start Date", fromHead) & clientRow).value, _
                        Range(headerFind("End Date", fromHead) & clientRow).value)
                Range(headerFind("Courtroom of Transfer (if relevant)", fromHead) & clientRow).value _
                    = Lookup("Courtroom_Name")(toCR)
            
            Case "WRAP", "5E"
            
            'TODO Close out
            Range(headerFind("End Date", fromHead) & clientRow).value _
                    = referralDate
                Range(headerFind("LOS", fromHead) & clientRow).value _
                    = DateDiff("d", _
                        Range(headerFind("Start Date", fromHead) & clientRow).value, _
                        Range(headerFind("End Date", fromHead) & clientRow).value)
                Range(headerFind("Courtroom of Transfer (if relevant)", fromHead) & clientRow).value _
                    = Lookup("Courtroom_Name")(toCR)
                
            Case "Adult"
                Range(headerFind("End Date", fromHead) & clientRow).value _
                    = referralDate
                Range(headerFind("LOS", fromHead) & clientRow).value _
                    = DateDiff("d", _
                        Range(headerFind("Start Date", fromHead) & clientRow).value, _
                        Range(headerFind("End Date", fromHead) & clientRow).value)
        End Select
    End If
    
                                ''''''''''''''''
                                'START NEW ROOM'
                                ''''''''''''''''
    If Not toCR = "N/A" Then
        Select Case toCR
            Case "4G", "4E", "6F", "6H", "3E"
                Range(headerFind("Was Youth in " & toCR & "?", toHead) & clientRow).value _
                    = Lookup("Generic_YN_Name")("Yes")
                Range(headerFind("Courtroom of Origin", toHead) & clientRow).value _
                    = Lookup("Courtroom_Name")(fromCR)
                Range(headerFind("Start Date", toHead) & clientRow).value _
                    = referralDate
                Range(headerFind("Age at Start of Courtroom", toHead) & clientRow).value _
                    = ageAtTime(referralDate, clientRow)
                
                Call append(Range(headerFind("Notes on " & toCR, toHead) & clientRow), Notes, referralDate)
                
                'Pre-fill Booleans
                Call flagNo(Range(headerFind("Was Youth on Pretrial?", toHead) & clientRow))
                Call flagNo(Range(headerFind("Was Youth on Consent Decree?", toHead) & clientRow))
                Call flagNo(Range(headerFind("Was Youth on Interim Probation?", toHead) & clientRow))
                Call flagNo(Range(headerFind("Was Youth on Probation?", toHead) & clientRow))
                Call flagNo(Range(headerFind("Was Youth on Aftercare Probation?", toHead) & clientRow))
                
                Call flagNo(Range(headerFind("Was Notice of Certification Given?", toHead) & clientRow))
                Call flagNo(Range(headerFind("Was Notice of De-Certification Given?", toHead) & clientRow))
                Call flagNo(Range(headerFind("Did Youth Enter an Admission?", toHead) & clientRow))
                Call flagNo(Range(headerFind("Adjudicated Delinquent?", toHead) & clientRow))
                Call flagNo(Range(headerFind("Did Youth Have a Continuance?", toHead) & clientRow))
                Call flagNo(Range(headerFind("Was Youth Placed?", toHead) & clientRow))
                
                
            Case "JTC"
                Call flagYes(Range(headerFind("Was Youth in JTC?", toHead) & clientRow))
                Range(headerFind("Courtroom of Origin", toHead) & clientRow).value _
                    = Lookup("Courtroom_Name")(fromCR)
                Range(headerFind("Referral Date", toHead) & clientRow).value _
                    = referralDate
                Range(headerFind("Age at Courtroom Referral", toHead) & clientRow).value _
                    = ageAtTime(referralDate, clientRow)
                Range(headerFind("Next Court Date", toHead) & clientRow).value _
                    = referralDate
                Range(headerFind("Phase", toHead) & clientRow).value = Lookup("JTC_Phase_Name")("Referred")
                Range(headerFind("Legal Status") & clientRow).value = Lookup("Legal_Status_Name")("JTC")
                
                Range(headerFind("Active or Discharged", toHead) & clientRow).value _
                    = Lookup("Active_Name")("Active")
                Range(headerFind("Nature of Discharge", toHead) & clientRow).value _
                    = Lookup("Nature_of_Discharge_Name")("Still Active")
                    
                Call flagNo(Range(headerFind("Was Youth Placed?", toHead) & clientRow))
                Call flagNo(Range(headerFind("Was Notice of Certification Given?", toHead) & clientRow))
                Call flagNo(Range(headerFind("Was Notice of De-Certification Given?", toHead) & clientRow))
                Call flagNo(Range(headerFind("Did Youth Enter an Admission?", toHead) & clientRow))
                Call flagNo(Range(headerFind("Adjudicated Delinquent?", toHead) & clientRow))
                Call flagNo(Range(headerFind("Did Youth Have a Continuance?", toHead) & clientRow))
                    
                
                    
            Case "WRAP", "5E"
                Dim altName As String
                
                If toCR = "5E" Then
                    altName = "Crossover"
                Else
                    altName = "WRAP"
                End If
                
                Call flagYes(Range(headerFind("Was Youth on " & altName & " Status?", toHead) & clientRow))
                Range(headerFind("Courtroom of Origin", toHead) & clientRow).value _
                    = Lookup("Courtroom_Name")(fromCR)
                Range(headerFind("Age at Courtroom Referral", toHead) & clientRow).value _
                    = DateDiff("d", Range(headerFind("DOB") & clientRow).value, referralDate) / 365
                Range(headerFind("Referral Date", toHead) & clientRow).value _
                    = referralDate
                Call append(Range(headerFind("Notes on " & altName, toHead) & clientRow), Notes, referralDate)
                
                'Pre-fill Booleans
                Call flagNo(Range(headerFind("Was Youth on Pretrial?", toHead) & clientRow))
                Call flagNo(Range(headerFind("Was Youth on Consent Decree?", toHead) & clientRow))
                Call flagNo(Range(headerFind("Was Youth on Interim Probation?", toHead) & clientRow))
                Call flagNo(Range(headerFind("Was Youth on Probation?", toHead) & clientRow))
                Call flagNo(Range(headerFind("Was Youth on Aftercare Probation?", toHead) & clientRow))
                
                Call flagNo(Range(headerFind("Was Notice of Certification Given?", toHead) & clientRow))
                Call flagNo(Range(headerFind("Was Notice of De-Certification Given?", toHead) & clientRow))
                Call flagNo(Range(headerFind("Did Youth Enter an Admission?", toHead) & clientRow))
                Call flagNo(Range(headerFind("Adjudicated Delinquent?", toHead) & clientRow))
                Call flagNo(Range(headerFind("Did Youth Have a Continuance?", toHead) & clientRow))
                Call flagNo(Range(headerFind("Was Youth Placed?", toHead) & clientRow))
                
            Case "Adult"
                Call flagYes(Range(headerFind("Was Youth in Adult?", toHead) & clientRow))
                Range(headerFind("Age at Start of Status", toHead) & clientRow).value _
                    = DateDiff("d", Range(headerFind("DOB") & clientRow).value, referralDate) / 365
                Range(headerFind("Start Date", toHead) & clientRow).value _
                    = referralDate
                Call append(Range(headerFind("Notes on " & toCR, toHead) & clientRow), Notes, referralDate)
            
            Case "PJJSC"
                Range(headerFind("Did Youth Have Initial Detention Hearing?", toHead) & clientRow).value _
                    = Lookup("Generic_YN_Name")("No")
        End Select
    End If
    
                                ''''''''''''''''''''''
                                'SET ACTIVE COURTROOM'
                                ''''''''''''''''''''''
    Select Case toCR
        Case "4G", "4E", "6F", "6H", "3E", "JTC", "WRAP", "Adult", "PJJSC", "5E"
            Range(headerFind("Active Courtroom") & clientRow).value _
            = Lookup("Courtroom_Name")(toCR)
    End Select
    
    If Not toCR = "N/A" And Not fromCR = "N/A" Then
        
        Dim i As Integer
        Dim bucketCount As Integer
        Dim sectionHead As String
        Dim bucketHead As String
        Dim agency As String
        
        
        'Duplicate ongoing Supervision Programs
        sectionHead = hFind("Supervision Programs", "AGGREGATES")
        For i = 20 To 1 Step -1
            bucketHead = headerFind("Supervision Ordered #" & i, sectionHead)
            
            If isNotEmptyOrZero(Range(bucketHead & clientRow)) _
                And isEmptyOrZero(Range(headerFind("End Date", bucketHead) & clientRow)) Then
                
                If isResidential(Lookup("Supervision_Program_Num")(Range(bucketHead & clientRow).value)) Then
                    agency = Lookup("Residential_Supervision_Provider_Num") _
                        (Range(headerFind("Residential Agency", bucketHead) & clientRow).value)
                Else
                    agency = Lookup("Community_Based_Supervision_Provider_Num") _
                        (Range(headerFind("Community-Based Agency", bucketHead) & clientRow).value)
                End If
                
                
                
                Call addSupervision( _
                    clientRow:=clientRow, _
                    serviceType:=Lookup("Supervision_Program_Num")(Range(bucketHead & clientRow).value), _
                    legalStatus:=Lookup("Legal_Status_Num")(Range(hFind("Legal Status") & clientRow).value), _
                    Courtroom:=Lookup("Courtroom_Num")(Range(hFind("Active Courtroom") & clientRow).value), _
                    DA:=Lookup("DA_Last_Name_Num")(Range(headerFind("DA", bucketHead) & clientRow).value), _
                    agency:=agency, _
                    startDate:=referralDate, _
                    NextCourtDate:=Range(headerFind("Next Court Date") & clientRow).value, _
                    Re1:="Other", _
                    Re2:="N/A", _
                    Re3:="N/A", _
                    Notes:="Tranferred courtroom during program")
            End If
        Next i
        
        'Drop old ongoing Supervision Programs
        Select Case toCR
            Case "4G", "4E", "6F", "6H", "3E", "WRAP", "5E", "JTC", "Adult"
                sectionHead = headerFind("Supervision Programs", fromHead)
                For i = 15 To 1 Step -1
                    If i > 5 And fromCR = "Adult" Then
                            i = 5
                    End If
                
                    bucketHead = headerFind("Supervision Ordered #" & i, sectionHead)
                    If isNotEmptyOrZero(Range(bucketHead & clientRow)) _
                        And isEmptyOrZero(Range(headerFind("End Date", bucketHead) & clientRow)) Then
                        
                        Call dropSupervision( _
                            clientRow:=clientRow, _
                            head:=bucketHead, _
                            serviceType:=Lookup("Supervision_Program_Num")(Range(bucketHead & clientRow).value), _
                            startDate:=Range(headerFind("Start Date", bucketHead) & clientRow).value, _
                            endDate:=referralDate, _
                            Nature:="Neutral", _
                            Re1:="Other", _
                            Re2:="N/A", _
                            Re3:="N/A", _
                            Notes:="Tranferred courtroom during program")
                        
                        'short-circuit Adult court b/c it only has 5 buckets
                        If i = 5 And fromCR = "Adult" Then
                            i = 20
                        End If
                    End If
                Next i
        End Select
                
              
              
        'Duplicate ongoing Conditions
        sectionHead = hFind("Conditions", "AGGREGATES")
        For i = 20 To 1 Step -1
            bucketHead = headerFind("Condition Ordered #" & i, sectionHead)
            
            If isNotEmptyOrZero(Range(bucketHead & clientRow)) _
                And isEmptyOrZero(Range(headerFind("End Date", bucketHead) & clientRow)) Then
                
                Call addCondition( _
                    clientRow:=clientRow, _
                    condition:=Lookup("Condition_Num")(Range(bucketHead & clientRow).value), _
                    legalStatus:=Lookup("Legal_Status_Num")(Range(hFind("Legal Status") & clientRow).value), _
                    Courtroom:=Lookup("Courtroom_Num")(Range(hFind("Active Courtroom") & clientRow).value), _
                    DA:=Lookup("DA_Last_Name_Num")(Range(headerFind("DA", bucketHead) & clientRow).value), _
                    agency:=Lookup("Condition_Provider_Num")(Range(headerFind("Condition Agency", bucketHead) & clientRow).value), _
                    startDate:=referralDate, _
                    Re1:="Other", _
                    Re2:="N/A", _
                    Re3:="N/A", _
                    Notes:="Tranferred courtroom during program")
            End If
        Next i
        
        'Drop old ongoing Conditions
        Select Case toCR
            Case "4G", "4E", "6F", "6H", "3E", "WRAP", "5E", "JTC", "Adult"
                sectionHead = headerFind("Conditions", fromHead)
                For i = 15 To 1 Step -1
                    If i > 5 And fromCR = "Adult" Then
                            i = 5
                    End If
                    bucketHead = headerFind("Condition Ordered #" & i, sectionHead)
                    If isNotEmptyOrZero(Range(bucketHead & clientRow)) _
                        And isEmptyOrZero(Range(headerFind("End Date", bucketHead) & clientRow)) Then
                        
                        Call dropCondition( _
                            clientRow:=clientRow, _
                            head:=bucketHead, _
                            condition:=Lookup("Condition_Num")(Range(bucketHead & clientRow).value), _
                            startDate:=Range(headerFind("Start Date", bucketHead) & clientRow).value, _
                            endDate:=referralDate, _
                            Nature:="Neutral", _
                            Re1:="Other", _
                            Re2:="N/A", _
                            Re3:="N/A", _
                            Notes:="Tranferred courtroom during program")
                    End If
                Next i
        End Select
                
        'Update Legal Status
        Select Case fromCR
            Case "4G", "4E", "6F", "6H", "3E"
                bucketHead = headerFind("Was Youth on Pretrial?", fromHead)
                If Range(bucketHead & clientRow).value = Lookup("Generic_YN_Name")("Yes") _
                    And isEmptyOrZero(Range(headerFind("End Date", bucketHead) & clientRow)) Then
                    
                    Range(headerFind("End Date", bucketHead) & clientRow).value = referralDate
                    
                    If newLegalStatus = "" Then
                        legalStatusVar = "Pretrial"
                    Else
                        legalStatusVar = newLegalStatus
                    End If
                    Range(hFind("Legal Status") & clientRow).value = Lookup("Legal_Status_Name")(legalStatusVar)
                    
                    Select Case toCR
                        Case "4G", "4E", "6F", "6H", "3E"
                            bucketHead = headerFind("Was Youth on " & legalStatusVar & "?", toHead)
                            Call flagYes(Range(bucketHead & clientRow))
                            Range(headerFind("Start Date", bucketHead) & clientRow).value = referralDate
                    End Select
                End If
                
                bucketHead = headerFind("Was Youth on Consent Decree?", fromHead)
                If Range(bucketHead & clientRow).value = Lookup("Generic_YN_Name")("Yes") _
                    And isEmptyOrZero(Range(headerFind("End Date", bucketHead) & clientRow)) Then
                    
                    Range(headerFind("End Date", bucketHead) & clientRow).value = referralDate
                    
                    If newLegalStatus = "" Then
                        legalStatusVar = "Consent Decree"
                    Else
                        legalStatusVar = newLegalStatus
                    End If
                    Range(hFind("Legal Status") & clientRow).value = Lookup("Legal_Status_Name")(legalStatusVar)
                    
                    Select Case toCR
                        Case "4G", "4E", "6F", "6H", "3E"
                            bucketHead = headerFind("Was Youth on " & legalStatusVar & "?", toHead)
                            Call flagYes(Range(bucketHead & clientRow))
                            Range(headerFind("Start Date", bucketHead) & clientRow).value = referralDate
                    End Select
                End If
                
                bucketHead = headerFind("Was Youth on Interim Probation?", fromHead)
                If Range(bucketHead & clientRow).value = Lookup("Generic_YN_Name")("Yes") _
                    And isEmptyOrZero(Range(headerFind("End Date", bucketHead) & clientRow)) Then
                    
                    Range(headerFind("End Date", bucketHead) & clientRow).value = referralDate
                    
                    If newLegalStatus = "" Then
                        legalStatusVar = "Interim Probation"
                    Else
                        legalStatusVar = newLegalStatus
                    End If
                    Range(hFind("Legal Status") & clientRow).value = Lookup("Legal_Status_Name")(legalStatusVar)
                    
                    Select Case toCR
                        Case "4G", "4E", "6F", "6H", "3E"
                            bucketHead = headerFind("Was Youth on " & legalStatusVar & "?", toHead)
                            Call flagYes(Range(bucketHead & clientRow))
                            Range(headerFind("Start Date", bucketHead) & clientRow).value = referralDate
                    End Select
                End If
                
                
                bucketHead = headerFind("Was Youth on Probation?", fromHead)
                If Range(bucketHead & clientRow).value = Lookup("Generic_YN_Name")("Yes") _
                    And isEmptyOrZero(Range(headerFind("End Date", bucketHead) & clientRow)) Then
                    
                    Range(headerFind("End Date", bucketHead) & clientRow).value = referralDate
                    
                    If newLegalStatus = "" Then
                        legalStatusVar = "Probation"
                    Else
                        legalStatusVar = newLegalStatus
                    End If
                    Range(hFind("Legal Status") & clientRow).value = Lookup("Legal_Status_Name")(legalStatusVar)
                    
                    Select Case toCR
                        Case "4G", "4E", "6F", "6H", "3E"
                            bucketHead = headerFind("Was Youth on " & legalStatusVar & "?", toHead)
                            Call flagYes(Range(bucketHead & clientRow))
                            Range(headerFind("Start Date", bucketHead) & clientRow).value = referralDate
                    End Select
                End If
                
                
                bucketHead = headerFind("Was Youth on Aftercare Probation?", fromHead)
                If Range(bucketHead & clientRow).value = Lookup("Generic_YN_Name")("Yes") _
                    And isEmptyOrZero(Range(headerFind("End Date", bucketHead) & clientRow)) Then
                    
                    Range(headerFind("End Date", bucketHead) & clientRow).value = referralDate
                    
                    If newLegalStatus = "" Then
                        legalStatusVar = "Aftercare Probation"
                    Else
                        legalStatusVar = newLegalStatus
                    End If
                    Range(hFind("Legal Status") & clientRow).value = Lookup("Legal_Status_Name")(legalStatusVar)
                    
                    Select Case toCR
                        Case "4G", "4E", "6F", "6H", "3E"
                            bucketHead = headerFind("Was Youth on " & legalStatusVar & "?", toHead)
                            Call flagYes(Range(bucketHead & clientRow))
                            Range(headerFind("Start Date", bucketHead) & clientRow).value = referralDate
                    End Select
                End If
        End Select
        
        Select Case toCR
            Case "WRAP"
                Range(hFind("Legal Status") & clientRow).value = Lookup("Legal_Status_Name")("Adult")
            Case "5E"
                Range(hFind("Legal Status") & clientRow).value = Lookup("Legal_Status_Name")("Crossover")
            Case "JTC"
                Range(hFind("Legal Status") & clientRow).value = Lookup("Legal_Status_Name")("JTC")
        End Select
    End If
    If Not newLegalStatus = "" Then
        Range(hFind("Legal Status") & clientRow).value = Lookup("Legal_Status_Name")(newLegalStatus)
    End If
End Sub


