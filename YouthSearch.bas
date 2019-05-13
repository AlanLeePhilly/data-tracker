Attribute VB_Name = "YouthSearch"
Sub ReturnToNavigationYouthSearch()
    Range("K5").Select
End Sub

Sub JumptoCourtroomHistory()
    Range("E81").Select
End Sub

Sub JumptoLegalStatusHistory()
    Range("T81").Select
End Sub
Sub JumptoSupervisionHistory()
    Range("F197").Select
End Sub
Sub JumptoConditionsHistory()
    Range("T197").Select
End Sub
Sub JumptoListingsHisory()
    Range("I483").Select
End Sub


Sub YouthSearchPrint()
    Call RefreshNamedRanges
    Call Generate_Dictionaries
    
    Dim PrintSheet As Worksheet
    Dim DataSheet As Worksheet
    Set PrintSheet = Worksheets("Youth Search")
    Set DataSheet = Worksheets("Entry")
    
    Dim userRow As Long
    userRow = PrintSheet.Range("J5").value
    
    'PRINT IDENTIFIERS BOX
    
    'example stringing together two different cell values to print
    PrintSheet.Range("D12").value = DataSheet.Range(hFind("First Name") & userRow).value & " " & DataSheet.Range(hFind("Last Name") & userRow).value
    
    'basic examples
    PrintSheet.Range("D14").value = DataSheet.Range(hFind("DOB") & userRow).value
    PrintSheet.Range("I12").value = DataSheet.Range(hFind("PID #") & userRow).value
    PrintSheet.Range("I14").value = DataSheet.Range(hFind("SID #") & userRow).value
    
    'examples with Lookup dictionary
    PrintSheet.Range("D19").value = Lookup("Sex_Num")(DataSheet.Range(hFind("Sex") & userRow).value)
    PrintSheet.Range("D21").value = Lookup("Race_Num")(DataSheet.Range(hFind("Race") & userRow).value)
    
    'basic example
    PrintSheet.Range("D23").value = DataSheet.Range(hFind("Address") & userRow).value
    PrintSheet.Range("D26").value = DataSheet.Range(hFind("Zipcode", "DEMOGRAPHICS") & userRow).value
    
    'example using a custom function
    PrintSheet.Range("I19").value = ageAtTime(Format(Now(), "mm/dd/yyyy"), userRow)
    
    PrintSheet.Range("I21").value = DataSheet.Range(hFind("Age @ Intake") & userRow).value
    PrintSheet.Range("I23").value = DataSheet.Range(hFind("Guardian First") & userRow).value & " " & DataSheet.Range(hFind("Guardian Last") & userRow).value

    PrintSheet.Range("E28").value = DataSheet.Range(hFind("School") & userRow).value
    PrintSheet.Range("E30").value = DataSheet.Range(hFind("Grade") & userRow).value
    PrintSheet.Range("I26").value = DataSheet.Range(hFind("Phone #") & userRow).value
    
    
    'PRINT ARREST & PETITION INFO BOX
    'Arrest info
    PrintSheet.Range("O12").value = DataSheet.Range(hFind("DC #") & userRow).value
    PrintSheet.Range("P14").value = DataSheet.Range(hFind("Incident District") & userRow).value
    PrintSheet.Range("P16").value = DataSheet.Range(hFind("Arrest Date") & userRow).value
    PrintSheet.Range("P18").value = Lookup("Generic_YNOU_Num")(DataSheet.Range(hFind("Active in System at Time of Arrest?", "Petition") & userRow).value)
    PrintSheet.Range("P20").value = DataSheet.Range(hFind("# of Prior Arrests") & userRow).value
    
    'Notes from Intake
    PrintSheet.Range("N23").value = DataSheet.Range(hFind("General Notes from Intake") & userRow).value
    
    'Petition #1
    PrintSheet.Range("T12").value = DataSheet.Range(hFind("Petition #1") & userRow).value
    PrintSheet.Range("Y12").value = DataSheet.Range(hFind("Lead Charge Name", "Petition #1") & userRow).value
    PrintSheet.Range("AC12").value = Lookup("Charge_Grade_Specific_Num")(DataSheet.Range(hFind("Charge Grade (specific) #1", "Petition #1") & userRow).value)
    PrintSheet.Range("Y14").value = DataSheet.Range(hFind("Charge Name #2", "Petition #1") & userRow).value
    PrintSheet.Range("AC14").value = Lookup("Charge_Grade_Specific_Num")(DataSheet.Range(hFind("Charge Grade (specific) #2", "Petition #1") & userRow).value)
    PrintSheet.Range("Y16").value = DataSheet.Range(hFind("Charge Name #3", "Petition #1") & userRow).value
    PrintSheet.Range("AC16").value = Lookup("Charge_Grade_Specific_Num")(DataSheet.Range(hFind("Charge Grade (specific) #3", "Petition #1") & userRow).value)
    PrintSheet.Range("Y18").value = DataSheet.Range(hFind("Charge Name #4", "Petition #1") & userRow).value
    PrintSheet.Range("AC18").value = Lookup("Charge_Grade_Specific_Num")(DataSheet.Range(hFind("Charge Grade (specific) #4", "Petition #1") & userRow).value)
    PrintSheet.Range("Y20").value = DataSheet.Range(hFind("Charge Name #5", "Petition #1") & userRow).value
    PrintSheet.Range("AC20").value = Lookup("Charge_Grade_Specific_Num")(DataSheet.Range(hFind("Charge Grade (specific) #5", "Petition #1") & userRow).value)
    
    PrintSheet.Range("T14").value = DataSheet.Range(hFind("Date Filed", "Petition #1") & userRow).value
    
    
    'Petition #2
    PrintSheet.Range("T23").value = DataSheet.Range(hFind("Petition #2") & userRow).value
    PrintSheet.Range("Y23").value = DataSheet.Range(hFind("Lead Charge Name", "Petition #2") & userRow).value
    PrintSheet.Range("AC23").value = Lookup("Charge_Grade_Specific_Num")(DataSheet.Range(hFind("Charge Grade (specific) #1", "Petition #2") & userRow).value)
    PrintSheet.Range("Y25").value = DataSheet.Range(hFind("Charge Name #2", "Petition #2") & userRow).value
    PrintSheet.Range("AC25").value = Lookup("Charge_Grade_Specific_Num")(DataSheet.Range(hFind("Charge Grade (specific) #2", "Petition #2") & userRow).value)
    PrintSheet.Range("Y27").value = DataSheet.Range(hFind("Charge Name #3", "Petition #2") & userRow).value
    PrintSheet.Range("AC27").value = Lookup("Charge_Grade_Specific_Num")(DataSheet.Range(hFind("Charge Grade (specific) #3", "Petition #2") & userRow).value)
    PrintSheet.Range("Y29").value = DataSheet.Range(hFind("Charge Name #4", "Petition #2") & userRow).value
    PrintSheet.Range("AC29").value = Lookup("Charge_Grade_Specific_Num")(DataSheet.Range(hFind("Charge Grade (specific) #4", "Petition #2") & userRow).value)
    PrintSheet.Range("Y31").value = DataSheet.Range(hFind("Charge Name #5", "Petition #2") & userRow).value
    PrintSheet.Range("AC31").value = Lookup("Charge_Grade_Specific_Num")(DataSheet.Range(hFind("Charge Grade (specific) #5", "Petition #2") & userRow).value)

    PrintSheet.Range("T25").value = DataSheet.Range(hFind("Date Filed", "Petition #2") & userRow).value


    'Petition #3
    PrintSheet.Range("T34").value = DataSheet.Range(hFind("Petition #3") & userRow).value
    PrintSheet.Range("Y34").value = DataSheet.Range(hFind("Lead Charge Name", "Petition #3") & userRow).value
    PrintSheet.Range("AC34").value = Lookup("Charge_Grade_Specific_Num")(DataSheet.Range(hFind("Charge Grade (specific) #1", "Petition #3") & userRow).value)
    PrintSheet.Range("Y36").value = DataSheet.Range(hFind("Charge Name #2", "Petition #3") & userRow).value
    PrintSheet.Range("AC36").value = Lookup("Charge_Grade_Specific_Num")(DataSheet.Range(hFind("Charge Grade (specific) #2", "Petition #3") & userRow).value)
    PrintSheet.Range("Y38").value = DataSheet.Range(hFind("Charge Name #3", "Petition #3") & userRow).value
    PrintSheet.Range("AC38").value = Lookup("Charge_Grade_Specific_Num")(DataSheet.Range(hFind("Charge Grade (specific) #3", "Petition #3") & userRow).value)
    PrintSheet.Range("Y40").value = DataSheet.Range(hFind("Charge Name #4", "Petition #3") & userRow).value
    PrintSheet.Range("AC40").value = Lookup("Charge_Grade_Specific_Num")(DataSheet.Range(hFind("Charge Grade (specific) #4", "Petition #3") & userRow).value)
    PrintSheet.Range("Y42").value = DataSheet.Range(hFind("Charge Name #5", "Petition #3") & userRow).value
    PrintSheet.Range("AC42").value = Lookup("Charge_Grade_Specific_Num")(DataSheet.Range(hFind("Charge Grade (specific) #5", "Petition #3") & userRow).value)

    PrintSheet.Range("T36").value = DataSheet.Range(hFind("Date Filed", "Petition #2") & userRow).value


    'Petition #4
    PrintSheet.Range("T45").value = DataSheet.Range(hFind("Petition #4") & userRow).value
    PrintSheet.Range("Y45").value = DataSheet.Range(hFind("Lead Charge Name", "Petition #4") & userRow).value
    PrintSheet.Range("AC45").value = Lookup("Charge_Grade_Specific_Num")(DataSheet.Range(hFind("Charge Grade (specific) #1", "Petition #4") & userRow).value)
    PrintSheet.Range("Y47").value = DataSheet.Range(hFind("Charge Name #2", "Petition #4") & userRow).value
    PrintSheet.Range("AC47").value = Lookup("Charge_Grade_Specific_Num")(DataSheet.Range(hFind("Charge Grade (specific) #2", "Petition #4") & userRow).value)
    PrintSheet.Range("Y49").value = DataSheet.Range(hFind("Charge Name #3", "Petition #4") & userRow).value
    PrintSheet.Range("AC49").value = Lookup("Charge_Grade_Specific_Num")(DataSheet.Range(hFind("Charge Grade (specific) #3", "Petition #4") & userRow).value)
    PrintSheet.Range("Y51").value = DataSheet.Range(hFind("Charge Name #4", "Petition #4") & userRow).value
    PrintSheet.Range("AC51").value = Lookup("Charge_Grade_Specific_Num")(DataSheet.Range(hFind("Charge Grade (specific) #4", "Petition #4") & userRow).value)
    PrintSheet.Range("Y53").value = DataSheet.Range(hFind("Charge Name #5", "Petition #4") & userRow).value
    PrintSheet.Range("AC53").value = Lookup("Charge_Grade_Specific_Num")(DataSheet.Range(hFind("Charge Grade (specific) #5", "Petition #4") & userRow).value)

    PrintSheet.Range("T47").value = DataSheet.Range(hFind("Date Filed", "Petition #2") & userRow).value


    'STATUS OF ARREST INCIDENT
    Dim activeStatus As String
    activeStatus = Lookup("Active_Num")(DataSheet.Range(hFind("Active or Discharged (in courtroom)?") & userRow).value)
    PrintSheet.Range("D37").value = activeStatus
    PrintSheet.Range("G37").value = Lookup("Generic_YNOU_Num")(DataSheet.Range(hFind("Did Youth Enter an Admission?", "COURT PROCEEDINGS", "AGGREGATES") & userRow).value)
    PrintSheet.Range("J37").value = Lookup("Generic_YNOU_Num")(DataSheet.Range(hFind("Adjudicated Delinquent?", "COURT PROCEEDINGS", "AGGREGATES") & userRow).value)
    PrintSheet.Range("G39").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Admissions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("J39").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Adjudicating Courtroom", "Adjudications", "AGGREGATES") & userRow).value)
    
    
    'SHONA'S LOS CALCULATIONS
    
    If StrComp(activeStatus, "Active") = 0 Then
        'LoS for arrest and petition
        Dim losArrest As Integer
        Dim losPetition As Integer
        
        Dim ArrestDate As String
        ArrestDate = DataSheet.Range(headerFind("Arrest Date (current petition)") & userRow).value
        losArrest = DateDiff("d", ArrestDate, Format(Now(), "mm/dd/yyyy"))
        PrintSheet.Range("D39").value = losArrest & " days"
        
        Dim petitionDate As String
        petitionDate = DataSheet.Range(hFind("Date Filed", "Petition") & userRow).value
        losPetition = DateDiff("d", petitionDate, Format(Now(), "mm/dd/yyyy"))
        PrintSheet.Range("D41").value = losPetition & " days"
        
        'Active court proceedings
        Dim Courtroom As String
        Dim losCourtroom As Integer
        Dim courtroomOptions(1 To 9) As String
        courtroomOptions(1) = "4G"
        courtroomOptions(2) = "4E"
        courtroomOptions(3) = "6F"
        courtroomOptions(4) = "6H"
        courtroomOptions(5) = "3E"
        courtroomOptions(6) = "Crossover"
        courtroomOptions(7) = "WRAP"
        courtroomOptions(8) = "JTC"
        courtroomOptions(9) = "ADULT"
        Courtroom = findFirstValue(DataSheet, userRow, "4G", courtroomOptions, "Start Date", "End Date")
        PrintSheet.Range("G48") = Courtroom
        
        losCourtroom = DateDiff("d", DataSheet.Range(hFind("Start Date", Courtroom, "4G") & userRow).value, _
            Format(Now(), "mm/dd/yyyy"))
        PrintSheet.Range("J48") = losCourtroom & " days"
        
        Dim legalStatus As String
        Dim losLegalStatus As Integer
        Dim lostLegalStatus As Integer
        Dim legalStatusOptions(1 To 5) As String
        legalStatusOptions(1) = "Pretrial"
        legalStatusOptions(2) = "Consent Decree"
        legalStatusOptions(3) = "Interim Probation"
        legalStatusOptions(4) = "Probation"
        legalStatusOptions(5) = "Aftercare Probation"
        legalStatus = findFirstValue(DataSheet, userRow, "Aggregates", legalStatusOptions, "Start Date", "End Date")
        PrintSheet.Range("G50") = legalStatus
        
        losLegalStatus = DateDiff("d", DataSheet.Range(hFind("Start Date", legalStatus, "Aggregates") & userRow).value, _
            Format(Now(), "mm/dd/yyyy"))
        PrintSheet.Range("J50") = losLegalStatus & " days"
        
        'Supervision Programs
        Dim supervisionProgramColumns() As String
        supervisionProgramColumns = findAllValues(DataSheet, userRow, "Aggregates", "Supervision Ordered", "Start Date", "End Date")
        Dim supervisionArrLength As Integer
        
        If (Not supervisionProgramColumns) = -1 Then
            supervisionArrLength = 0
        Else
            supervisionArrLength = UBound(supervisionProgramColumns) - LBound(supervisionProgramColumns) + 1
        End If
        
        Dim supervisionI As Integer
        Dim supervisionStart As String
        
        For supervisionI = 1 To supervisionArrLength
            PrintSheet.Range("D" & 53 + 2 * supervisionI) = Lookup("Supervision_Program_Num")(DataSheet.Range(hFind(supervisionProgramColumns(supervisionI - 1), "Aggregates") & userRow).value)
            supervisionStart = DataSheet.Range(hFind("Start Date", supervisionProgramColumns(supervisionI - 1), "Aggregates") & userRow).value
            PrintSheet.Range("J" & 53 + 2 * supervisionI) = DateDiff("d", supervisionStart, Format(Now(), "mm/dd/yyyy")) & " days"
            If supervisionI = 3 Then Exit For
        Next supervisionI
        
        'Conditions
        Dim conditionsColumns() As String
        conditionsColumns = findAllValues(DataSheet, userRow, "Aggregates", "Condition Ordered", "Start Date", "End Date")
        Dim conditionArrLength As Integer
        
        If (Not conditionsColumns) = -1 Then
            conditionArrLength = 0
        Else
            conditionArrLength = UBound(conditionsColumns) - LBound(conditionsColumns) + 1
        End If
        Dim conditionI As Integer
        Dim conditionStart As String
        
        For conditionI = 1 To conditionArrLength
            PrintSheet.Range("D" & 62 + 2 * conditionI) = Lookup("Condition_Num")(DataSheet.Range(hFind(conditionsColumns(conditionI - 1), "Aggregates") & userRow).value)
            conditionStart = DataSheet.Range(hFind("Start Date", conditionsColumns(conditionI - 1), "Aggregates") & userRow).value
            PrintSheet.Range("J" & 62 + 2 * conditionI) = DateDiff("d", conditionStart, Format(Now(), "mm/dd/yyyy")) & " days"
            If conditionI = 6 Then Exit For
        Next conditionI
    End If
    
    PrintSheet.Range("D43").value = DataSheet.Range(hFind("Total LOS From Arrest", "Petition Outcomes") & userRow).value
    
    'COURT PROCEEDINGS
    PrintSheet.Range("D48").value = DataSheet.Range(hFind("Next Court Date") & userRow).value
    'Replaced by courtroom & legal status calculation above
    'PrintSheet.Range("G48").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Active Courtroom") & userRow).value)
    'PrintSheet.Range("G50").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status") & userRow).value)
    
    
    'SUPERVISION PROGRAMS
    'PrintSheet.Range("D55").value = Lookup("Supervision_Program_Num")(DataSheet.Range(hFind("Active Supervision") & userRow).value)
    'PrintSheet.Range("G55").value = Lookup("Residential_Supervision_Provider_Num")(DataSheet.Range(hFind("Active Supervision Provider") & userRow).value)
    
    'COURTROOM HISTORY
    'Detention
    PrintSheet.Range("D84").value = Lookup("Generic_YNOU_Num")(DataSheet.Range(hFind("Did Youth Have Initial Detention Hearing?", "DETENTION") & userRow).value)
    PrintSheet.Range("G84").value = DataSheet.Range(hFind("Date of Initial Detention Hearing") & userRow).value
    PrintSheet.Range("J84").value = DataSheet.Range(hFind("Date of Release", "DETENTION") & userRow).value
    PrintSheet.Range("N84").value = DataSheet.Range(hFind("LOS in Detention", "DETENTION") & userRow).value
    PrintSheet.Range("D86").value = DataSheet.Range(hFind("Notes on Detention", "DETENTION") & userRow).value
    
    '4G
    PrintSheet.Range("D94").value = Lookup("Generic_YNOU_Num")(DataSheet.Range(hFind("Was Youth in 4G?") & userRow).value)
    PrintSheet.Range("G94").value = DataSheet.Range(hFind("Start Date", "4G") & userRow).value
    PrintSheet.Range("J94").value = DataSheet.Range(hFind("End Date", "4G") & userRow).value
    PrintSheet.Range("N94").value = DataSheet.Range(hFind("LOS", "4G") & userRow).value
    PrintSheet.Range("D96").value = DataSheet.Range(hFind("Notes on 4G", "4G") & userRow).value
    
    '3E
    PrintSheet.Range("D104").value = Lookup("Generic_YNOU_Num")(DataSheet.Range(hFind("Was Youth in 3E?") & userRow).value)
    PrintSheet.Range("G104").value = DataSheet.Range(hFind("Start Date", "3E") & userRow).value
    PrintSheet.Range("J104").value = DataSheet.Range(hFind("End Date", "3E") & userRow).value
    PrintSheet.Range("N104").value = DataSheet.Range(hFind("LOS", "3E") & userRow).value
    PrintSheet.Range("D106").value = DataSheet.Range(hFind("Notes on 3E", "3E") & userRow).value
    
    'JTC
    PrintSheet.Range("D114").value = Lookup("Generic_YNOU_Num")(DataSheet.Range(hFind("Was Youth in JTC?") & userRow).value)
    PrintSheet.Range("G114").value = DataSheet.Range(hFind("Referral Date", "JTC") & userRow).value
    PrintSheet.Range("J114").value = DataSheet.Range(hFind("Date of Overall Discharge", "JTC") & userRow).value
    PrintSheet.Range("N114").value = DataSheet.Range(hFind("Total LOS in JTC", "JTC") & userRow).value
    PrintSheet.Range("D116").value = DataSheet.Range(hFind("Notes on Outcome", "JTC") & userRow).value
    
    'CROSSOVER
    PrintSheet.Range("D124").value = Lookup("Generic_YNOU_Num")(DataSheet.Range(hFind("Was Youth on Crossover Status?") & userRow).value)
    PrintSheet.Range("G124").value = DataSheet.Range(hFind("Referral Date", "Crossover") & userRow).value
    PrintSheet.Range("J124").value = DataSheet.Range(hFind("End Date", "Crossover") & userRow).value
    PrintSheet.Range("N124").value = DataSheet.Range(hFind("LOS", "Crossover") & userRow).value
    PrintSheet.Range("D126").value = DataSheet.Range(hFind("Notes on Outcome", "Crossover") & userRow).value
    
    'WRAP
    PrintSheet.Range("D134").value = Lookup("Generic_YNOU_Num")(DataSheet.Range(hFind("Was Youth on WRAP Status?") & userRow).value)
    PrintSheet.Range("G134").value = DataSheet.Range(hFind("Referral Date", "WRAP") & userRow).value
    PrintSheet.Range("J134").value = DataSheet.Range(hFind("End Date", "WRAP") & userRow).value
    PrintSheet.Range("N134").value = DataSheet.Range(hFind("LOS", "WRAP") & userRow).value
    PrintSheet.Range("D136").value = DataSheet.Range(hFind("Notes on Outcome", "WRAP") & userRow).value
    
    '4E
    PrintSheet.Range("D144").value = Lookup("Generic_YNOU_Num")(DataSheet.Range(hFind("Was Youth in 4E?") & userRow).value)
    PrintSheet.Range("G144").value = DataSheet.Range(hFind("Start Date", "4E") & userRow).value
    PrintSheet.Range("J144").value = DataSheet.Range(hFind("End Date", "4E") & userRow).value
    PrintSheet.Range("N144").value = DataSheet.Range(hFind("LOS", "4E") & userRow).value
    PrintSheet.Range("D146").value = DataSheet.Range(hFind("Notes on 4E", "4E") & userRow).value
    
    '5F
    
    
    '6F
    PrintSheet.Range("D164").value = Lookup("Generic_YNOU_Num")(DataSheet.Range(hFind("Was Youth in 6F?") & userRow).value)
    PrintSheet.Range("G164").value = DataSheet.Range(hFind("Start Date", "6F") & userRow).value
    PrintSheet.Range("J164").value = DataSheet.Range(hFind("End Date", "6F") & userRow).value
    PrintSheet.Range("N164").value = DataSheet.Range(hFind("LOS", "6F") & userRow).value
    PrintSheet.Range("D166").value = DataSheet.Range(hFind("Notes on 6F", "6F") & userRow).value
    
    '6H
    PrintSheet.Range("D174").value = Lookup("Generic_YNOU_Num")(DataSheet.Range(hFind("Was Youth in 6H?") & userRow).value)
    PrintSheet.Range("G174").value = DataSheet.Range(hFind("Start Date", "6H") & userRow).value
    PrintSheet.Range("J174").value = DataSheet.Range(hFind("End Date", "6H") & userRow).value
    PrintSheet.Range("N174").value = DataSheet.Range(hFind("LOS", "6H") & userRow).value
    PrintSheet.Range("D176").value = DataSheet.Range(hFind("Notes on 6H", "6H") & userRow).value
    
    'Adult
    PrintSheet.Range("D184").value = Lookup("Generic_YNOU_Num")(DataSheet.Range(hFind("Was Youth in Adult?") & userRow).value)
    PrintSheet.Range("G184").value = DataSheet.Range(hFind("Start Date", "Adult") & userRow).value
    PrintSheet.Range("J184").value = DataSheet.Range(hFind("End Date", "Adult") & userRow).value
    PrintSheet.Range("N184").value = DataSheet.Range(hFind("LOS", "Adult") & userRow).value
    PrintSheet.Range("D186").value = DataSheet.Range(hFind("Notes on Adult", "Adult") & userRow).value
    
    
    
    'LEGAL STATUS HISTORY
    'Diversion
    PrintSheet.Range("S84").value = Lookup("Generic_YNOU_Num")(DataSheet.Range(hFind("Referred to Diversion?") & userRow).value)
    PrintSheet.Range("V84").value = DataSheet.Range(hFind("Referral Date", "DIVERSION") & userRow).value
    PrintSheet.Range("Z84").value = DataSheet.Range(hFind("Discharge Date", "DIVERSION") & userRow).value
    PrintSheet.Range("AC84").value = DataSheet.Range(hFind("LOS Diversion") & userRow).value
    PrintSheet.Range("V88").value = DataSheet.Range(hFind("Diversion Notes") & userRow).value

    'Pretrial
    PrintSheet.Range("S96").value = Lookup("Generic_YNOU_Num")(DataSheet.Range(hFind("Was Youth on Pretrial?", "AGGREGATES") & userRow).value)
    PrintSheet.Range("V96").value = DataSheet.Range(hFind("Start Date", "Pretrial", "AGGREGATES") & userRow).value
    PrintSheet.Range("Z96").value = DataSheet.Range(hFind("End Date", "Pretrial", "AGGREGATES") & userRow).value
    PrintSheet.Range("AC96").value = DataSheet.Range(hFind("LOS", "Pretrial", "AGGREGATES") & userRow).value
    PrintSheet.Range("V98").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Origin", "Pretrial", "AGGREGATES") & userRow).value)
    PrintSheet.Range("Z98").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Discharging Courtroom", "Pretrial", "AGGREGATES") & userRow).value)
    PrintSheet.Range("V100").value = DataSheet.Range(hFind("Notes on Pretrial", "AGGREGATES") & userRow).value

    'Consent Decree
    PrintSheet.Range("S108").value = Lookup("Generic_YNOU_Num")(DataSheet.Range(hFind("Was Youth on Consent Decree?", "AGGREGATES") & userRow).value)
    PrintSheet.Range("V108").value = DataSheet.Range(hFind("Start Date", "Consent Decree", "AGGREGATES") & userRow).value
    PrintSheet.Range("Z108").value = DataSheet.Range(hFind("End Date", "Consent Decree", "AGGREGATES") & userRow).value
    PrintSheet.Range("AC108").value = DataSheet.Range(hFind("LOS", "Consent Decree", "AGGREGATES") & userRow).value
    PrintSheet.Range("V110").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Origin", "Consent Decree", "AGGREGATES") & userRow).value)
    PrintSheet.Range("Z110").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Discharging Courtroom", "Consent Decree", "AGGREGATES") & userRow).value)
    PrintSheet.Range("V112").value = DataSheet.Range(hFind("Notes on Consent Decree", "AGGREGATES") & userRow).value

    'Interim/Deferred
    PrintSheet.Range("S120").value = Lookup("Generic_YNOU_Num")(DataSheet.Range(hFind("Was Youth on Interim Probation?", "AGGREGATES") & userRow).value)
    PrintSheet.Range("V120").value = DataSheet.Range(hFind("Start Date", "Interim Probation", "AGGREGATES") & userRow).value
    PrintSheet.Range("Z120").value = DataSheet.Range(hFind("End Date", "Interim Probation", "AGGREGATES") & userRow).value
    PrintSheet.Range("AC120").value = DataSheet.Range(hFind("LOS", "Interim Probation", "AGGREGATES") & userRow).value
    PrintSheet.Range("V122").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Origin", "Interim Probation", "AGGREGATES") & userRow).value)
    PrintSheet.Range("Z122").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Discharging Courtroom", "Interim Probation", "AGGREGATES") & userRow).value)
    PrintSheet.Range("V124").value = DataSheet.Range(hFind("Notes on Interim Probation", "AGGREGATES") & userRow).value


    'Probation
    PrintSheet.Range("S132").value = Lookup("Generic_YNOU_Num")(DataSheet.Range(hFind("Was Youth on Probation?", "AGGREGATES") & userRow).value)
    PrintSheet.Range("V132").value = DataSheet.Range(hFind("Start Date", "Probation", "AGGREGATES") & userRow).value
    PrintSheet.Range("Z132").value = DataSheet.Range(hFind("End Date", "Probation", "AGGREGATES") & userRow).value
    PrintSheet.Range("AC132").value = DataSheet.Range(hFind("LOS", "Probation", "AGGREGATES") & userRow).value
    PrintSheet.Range("V134").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Origin", "Probation", "AGGREGATES") & userRow).value)
    PrintSheet.Range("Z134").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Discharging Courtroom", "Probation", "AGGREGATES") & userRow).value)
    PrintSheet.Range("V136").value = DataSheet.Range(hFind("Notes on Probation", "AGGREGATES") & userRow).value
    
    'Aftercare Probation
    PrintSheet.Range("S144").value = Lookup("Generic_YNOU_Num")(DataSheet.Range(hFind("Was Youth on Aftercare Probation?", "AGGREGATES") & userRow).value)
    PrintSheet.Range("V144").value = DataSheet.Range(hFind("Start Date", "Aftercare Probation", "AGGREGATES") & userRow).value
    PrintSheet.Range("Z144").value = DataSheet.Range(hFind("End Date", "Aftercare Probation", "AGGREGATES") & userRow).value
    PrintSheet.Range("AC144").value = DataSheet.Range(hFind("LOS", "Aftercare Probation", "AGGREGATES") & userRow).value
    PrintSheet.Range("V146").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Origin", "Aftercare Probation", "AGGREGATES") & userRow).value)
    PrintSheet.Range("Z146").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Discharging Courtroom", "Aftercare Probation", "AGGREGATES") & userRow).value)
    PrintSheet.Range("V148").value = DataSheet.Range(hFind("Notes on Aftercare Probation", "AGGREGATES") & userRow).value
    
    'Adult
    PrintSheet.Range("S144").value = Lookup("Generic_YNOU_Num")(DataSheet.Range(hFind("Was Youth on Aftercare Probation?", "AGGREGATES") & userRow).value)
    PrintSheet.Range("V144").value = DataSheet.Range(hFind("Start Date", "Aftercare Probation", "AGGREGATES") & userRow).value
    PrintSheet.Range("Z144").value = DataSheet.Range(hFind("End Date", "Aftercare Probation", "AGGREGATES") & userRow).value
    PrintSheet.Range("AC144").value = DataSheet.Range(hFind("LOS", "Aftercare Probation", "AGGREGATES") & userRow).value
    PrintSheet.Range("V146").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Origin", "Aftercare Probation", "AGGREGATES") & userRow).value)
    PrintSheet.Range("Z146").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Discharging Courtroom", "Aftercare Probation", "AGGREGATES") & userRow).value)
    PrintSheet.Range("V148").value = DataSheet.Range(hFind("Notes on Aftercare Probation", "AGGREGATES") & userRow).value

   
    Call YouthSearchPrint2

End Sub


Sub YouthSearchPrint2()

    Call RefreshNamedRanges
    Call Generate_Dictionaries
    
    Dim PrintSheet As Worksheet
    Dim DataSheet As Worksheet
    Set PrintSheet = Worksheets("Youth Search")
    Set DataSheet = Worksheets("Entry")
    
    Dim userRow As Long
    userRow = PrintSheet.Range("J5").value


    'SUPERVISION HISTORY
    '#1
    PrintSheet.Range("D200").value = Lookup("Supervision_Program_Num")(DataSheet.Range(hFind("Supervision Ordered #1", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("G200").value = DataSheet.Range(hFind("Start Date", "Supervision Ordered #1", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G202").value = DataSheet.Range(hFind("End Date", "Supervision Ordered #1", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G204").value = DataSheet.Range(hFind("LOS", "Supervision Ordered #1", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("E206").value = DataSheet.Range(hFind("Supervision Description", "Supervision Ordered #1", "Supervision Programs", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Supervision Ordered #1", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("K200").value = Lookup("Community_Based_Supervision_Provider_Num")(DataSheet.Range(hFind("Community-Based Agency", "Supervision Ordered #1", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("N200").value = Lookup("Residential_Supervision_Provider_Num")(DataSheet.Range(hFind("Residential Agency", "Supervision Ordered #1", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K202").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Supervision Ordered #1", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K204").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Supervision Ordered #1", "Supervision Programs", "AGGREGATES") & userRow).value)

    '#2
    PrintSheet.Range("D214").value = Lookup("Supervision_Program_Num")(DataSheet.Range(hFind("Supervision Ordered #2", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("G214").value = DataSheet.Range(hFind("Start Date", "Supervision Ordered #2", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G216").value = DataSheet.Range(hFind("End Date", "Supervision Ordered #2", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G218").value = DataSheet.Range(hFind("LOS", "Supervision Ordered #2", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("E220").value = DataSheet.Range(hFind("Supervision Description", "Supervision Ordered #2", "Supervision Programs", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Supervision Ordered #2", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("K214").value = Lookup("Community_Based_Supervision_Provider_Num")(DataSheet.Range(hFind("Community-Based Agency", "Supervision Ordered #2", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("N214").value = Lookup("Residential_Supervision_Provider_Num")(DataSheet.Range(hFind("Residential Agency", "Supervision Ordered #2", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K216").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Supervision Ordered #2", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K218").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Supervision Ordered #2", "Supervision Programs", "AGGREGATES") & userRow).value)
    
    '#3
    PrintSheet.Range("D228").value = Lookup("Supervision_Program_Num")(DataSheet.Range(hFind("Supervision Ordered #3", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("G228").value = DataSheet.Range(hFind("Start Date", "Supervision Ordered #3", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G230").value = DataSheet.Range(hFind("End Date", "Supervision Ordered #3", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G232").value = DataSheet.Range(hFind("LOS", "Supervision Ordered #3", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("E234").value = DataSheet.Range(hFind("Supervision Description", "Supervision Ordered #3", "Supervision Programs", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Supervision Ordered #3", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("K228").value = Lookup("Community_Based_Supervision_Provider_Num")(DataSheet.Range(hFind("Community-Based Agency", "Supervision Ordered #3", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("N228").value = Lookup("Residential_Supervision_Provider_Num")(DataSheet.Range(hFind("Residential Agency", "Supervision Ordered #3", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K230").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Supervision Ordered #3", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K232").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Supervision Ordered #3", "Supervision Programs", "AGGREGATES") & userRow).value)

    '#4
    PrintSheet.Range("D242").value = Lookup("Supervision_Program_Num")(DataSheet.Range(hFind("Supervision Ordered #4", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("G242").value = DataSheet.Range(hFind("Start Date", "Supervision Ordered #4", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G244").value = DataSheet.Range(hFind("End Date", "Supervision Ordered #4", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G246").value = DataSheet.Range(hFind("LOS", "Supervision Ordered #4", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("E248").value = DataSheet.Range(hFind("Supervision Description", "Supervision Ordered #4", "Supervision Programs", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Supervision Ordered #4", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("K242").value = Lookup("Community_Based_Supervision_Provider_Num")(DataSheet.Range(hFind("Community-Based Agency", "Supervision Ordered #4", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("N242").value = Lookup("Residential_Supervision_Provider_Num")(DataSheet.Range(hFind("Residential Agency", "Supervision Ordered #4", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K244").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Supervision Ordered #4", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K246").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Supervision Ordered #4", "Supervision Programs", "AGGREGATES") & userRow).value)

    '#5
    PrintSheet.Range("D256").value = Lookup("Supervision_Program_Num")(DataSheet.Range(hFind("Supervision Ordered #5", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("G256").value = DataSheet.Range(hFind("Start Date", "Supervision Ordered #5", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G258").value = DataSheet.Range(hFind("End Date", "Supervision Ordered #5", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G260").value = DataSheet.Range(hFind("LOS", "Supervision Ordered #5", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("E262").value = DataSheet.Range(hFind("Supervision Description", "Supervision Ordered #5", "Supervision Programs", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Supervision Ordered #5", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("K256").value = Lookup("Community_Based_Supervision_Provider_Num")(DataSheet.Range(hFind("Community-Based Agency", "Supervision Ordered #5", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("N256").value = Lookup("Residential_Supervision_Provider_Num")(DataSheet.Range(hFind("Residential Agency", "Supervision Ordered #5", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K258").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Supervision Ordered #5", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K260").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Supervision Ordered #5", "Supervision Programs", "AGGREGATES") & userRow).value)

    '#6
    PrintSheet.Range("D270").value = Lookup("Supervision_Program_Num")(DataSheet.Range(hFind("Supervision Ordered #6", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("G270").value = DataSheet.Range(hFind("Start Date", "Supervision Ordered #6", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G272").value = DataSheet.Range(hFind("End Date", "Supervision Ordered #6", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G274").value = DataSheet.Range(hFind("LOS", "Supervision Ordered #6", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("E276").value = DataSheet.Range(hFind("Supervision Description", "Supervision Ordered #6", "Supervision Programs", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Supervision Ordered #6", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("K270").value = Lookup("Community_Based_Supervision_Provider_Num")(DataSheet.Range(hFind("Community-Based Agency", "Supervision Ordered #6", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("N270").value = Lookup("Residential_Supervision_Provider_Num")(DataSheet.Range(hFind("Residential Agency", "Supervision Ordered #6", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K272").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Supervision Ordered #6", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K274").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Supervision Ordered #6", "Supervision Programs", "AGGREGATES") & userRow).value)

    '#7
    PrintSheet.Range("D284").value = Lookup("Supervision_Program_Num")(DataSheet.Range(hFind("Supervision Ordered #7", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("G284").value = DataSheet.Range(hFind("Start Date", "Supervision Ordered #7", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G286").value = DataSheet.Range(hFind("End Date", "Supervision Ordered #7", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G288").value = DataSheet.Range(hFind("LOS", "Supervision Ordered #7", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("E290").value = DataSheet.Range(hFind("Supervision Description", "Supervision Ordered #7", "Supervision Programs", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Supervision Ordered #7", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("K284").value = Lookup("Community_Based_Supervision_Provider_Num")(DataSheet.Range(hFind("Community-Based Agency", "Supervision Ordered #7", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("N284").value = Lookup("Residential_Supervision_Provider_Num")(DataSheet.Range(hFind("Residential Agency", "Supervision Ordered #7", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K286").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Supervision Ordered #7", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K288").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Supervision Ordered #7", "Supervision Programs", "AGGREGATES") & userRow).value)

    '#8
    PrintSheet.Range("D298").value = Lookup("Supervision_Program_Num")(DataSheet.Range(hFind("Supervision Ordered #8", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("G298").value = DataSheet.Range(hFind("Start Date", "Supervision Ordered #8", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G300").value = DataSheet.Range(hFind("End Date", "Supervision Ordered #8", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G302").value = DataSheet.Range(hFind("LOS", "Supervision Ordered #8", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("E304").value = DataSheet.Range(hFind("Supervision Description", "Supervision Ordered #8", "Supervision Programs", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Supervision Ordered #8", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("K298").value = Lookup("Community_Based_Supervision_Provider_Num")(DataSheet.Range(hFind("Community-Based Agency", "Supervision Ordered #8", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("N298").value = Lookup("Residential_Supervision_Provider_Num")(DataSheet.Range(hFind("Residential Agency", "Supervision Ordered #8", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K300").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Supervision Ordered #8", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K302").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Supervision Ordered #8", "Supervision Programs", "AGGREGATES") & userRow).value)

    '#9
    PrintSheet.Range("D312").value = Lookup("Supervision_Program_Num")(DataSheet.Range(hFind("Supervision Ordered #9", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("G312").value = DataSheet.Range(hFind("Start Date", "Supervision Ordered #9", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G314").value = DataSheet.Range(hFind("End Date", "Supervision Ordered #9", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G316").value = DataSheet.Range(hFind("LOS", "Supervision Ordered #9", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("E318").value = DataSheet.Range(hFind("Supervision Description", "Supervision Ordered #9", "Supervision Programs", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Supervision Ordered #9", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("K312").value = Lookup("Community_Based_Supervision_Provider_Num")(DataSheet.Range(hFind("Community-Based Agency", "Supervision Ordered #9", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("N312").value = Lookup("Residential_Supervision_Provider_Num")(DataSheet.Range(hFind("Residential Agency", "Supervision Ordered #9", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K314").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Supervision Ordered #9", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K316").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Supervision Ordered #9", "Supervision Programs", "AGGREGATES") & userRow).value)

    '#10
    PrintSheet.Range("D326").value = Lookup("Supervision_Program_Num")(DataSheet.Range(hFind("Supervision Ordered #10", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("G326").value = DataSheet.Range(hFind("Start Date", "Supervision Ordered #10", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G328").value = DataSheet.Range(hFind("End Date", "Supervision Ordered #10", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G330").value = DataSheet.Range(hFind("LOS", "Supervision Ordered #10", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("E332").value = DataSheet.Range(hFind("Supervision Description", "Supervision Ordered #10", "Supervision Programs", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Supervision Ordered #10", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("K326").value = Lookup("Community_Based_Supervision_Provider_Num")(DataSheet.Range(hFind("Community-Based Agency", "Supervision Ordered #10", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("N326").value = Lookup("Residential_Supervision_Provider_Num")(DataSheet.Range(hFind("Residential Agency", "Supervision Ordered #10", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K328").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Supervision Ordered #10", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K330").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Supervision Ordered #10", "Supervision Programs", "AGGREGATES") & userRow).value)

    '#11
    PrintSheet.Range("D340").value = Lookup("Supervision_Program_Num")(DataSheet.Range(hFind("Supervision Ordered #11", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("G340").value = DataSheet.Range(hFind("Start Date", "Supervision Ordered #11", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G342").value = DataSheet.Range(hFind("End Date", "Supervision Ordered #11", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G344").value = DataSheet.Range(hFind("LOS", "Supervision Ordered #11", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("E346").value = DataSheet.Range(hFind("Supervision Description", "Supervision Ordered #11", "Supervision Programs", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Supervision Ordered #11", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("K340").value = Lookup("Community_Based_Supervision_Provider_Num")(DataSheet.Range(hFind("Community-Based Agency", "Supervision Ordered #11", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("N340").value = Lookup("Residential_Supervision_Provider_Num")(DataSheet.Range(hFind("Residential Agency", "Supervision Ordered #11", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K342").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Supervision Ordered #11", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K344").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Supervision Ordered #11", "Supervision Programs", "AGGREGATES") & userRow).value)

    '#12
    PrintSheet.Range("D354").value = Lookup("Supervision_Program_Num")(DataSheet.Range(hFind("Supervision Ordered #12", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("G354").value = DataSheet.Range(hFind("Start Date", "Supervision Ordered #12", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G356").value = DataSheet.Range(hFind("End Date", "Supervision Ordered #12", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G358").value = DataSheet.Range(hFind("LOS", "Supervision Ordered #12", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("E360").value = DataSheet.Range(hFind("Supervision Description", "Supervision Ordered #12", "Supervision Programs", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Supervision Ordered #12", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("K354").value = Lookup("Community_Based_Supervision_Provider_Num")(DataSheet.Range(hFind("Community-Based Agency", "Supervision Ordered #12", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("N354").value = Lookup("Residential_Supervision_Provider_Num")(DataSheet.Range(hFind("Residential Agency", "Supervision Ordered #12", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K356").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Supervision Ordered #12", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K358").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Supervision Ordered #12", "Supervision Programs", "AGGREGATES") & userRow).value)

    '#13
    PrintSheet.Range("D368").value = Lookup("Supervision_Program_Num")(DataSheet.Range(hFind("Supervision Ordered #13", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("G368").value = DataSheet.Range(hFind("Start Date", "Supervision Ordered #13", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G370").value = DataSheet.Range(hFind("End Date", "Supervision Ordered #13", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G372").value = DataSheet.Range(hFind("LOS", "Supervision Ordered #13", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("E374").value = DataSheet.Range(hFind("Supervision Description", "Supervision Ordered #13", "Supervision Programs", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Supervision Ordered #13", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("K368").value = Lookup("Community_Based_Supervision_Provider_Num")(DataSheet.Range(hFind("Community-Based Agency", "Supervision Ordered #13", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("N368").value = Lookup("Residential_Supervision_Provider_Num")(DataSheet.Range(hFind("Residential Agency", "Supervision Ordered #13", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K370").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Supervision Ordered #13", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K372").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Supervision Ordered #13", "Supervision Programs", "AGGREGATES") & userRow).value)

    '#14
    PrintSheet.Range("D382").value = Lookup("Supervision_Program_Num")(DataSheet.Range(hFind("Supervision Ordered #14", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("G382").value = DataSheet.Range(hFind("Start Date", "Supervision Ordered #14", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G384").value = DataSheet.Range(hFind("End Date", "Supervision Ordered #14", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G386").value = DataSheet.Range(hFind("LOS", "Supervision Ordered #14", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("E388").value = DataSheet.Range(hFind("Supervision Description", "Supervision Ordered #14", "Supervision Programs", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Supervision Ordered #14", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("K382").value = Lookup("Community_Based_Supervision_Provider_Num")(DataSheet.Range(hFind("Community-Based Agency", "Supervision Ordered #14", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("N382").value = Lookup("Residential_Supervision_Provider_Num")(DataSheet.Range(hFind("Residential Agency", "Supervision Ordered #14", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K384").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Supervision Ordered #14", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K386").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Supervision Ordered #14", "Supervision Programs", "AGGREGATES") & userRow).value)

    '#15
    PrintSheet.Range("D396").value = Lookup("Supervision_Program_Num")(DataSheet.Range(hFind("Supervision Ordered #15", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("G396").value = DataSheet.Range(hFind("Start Date", "Supervision Ordered #15", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G398").value = DataSheet.Range(hFind("End Date", "Supervision Ordered #15", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G400").value = DataSheet.Range(hFind("LOS", "Supervision Ordered #15", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("E402").value = DataSheet.Range(hFind("Supervision Description", "Supervision Ordered #15", "Supervision Programs", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Supervision Ordered #15", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("K396").value = Lookup("Community_Based_Supervision_Provider_Num")(DataSheet.Range(hFind("Community-Based Agency", "Supervision Ordered #15", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("N396").value = Lookup("Residential_Supervision_Provider_Num")(DataSheet.Range(hFind("Residential Agency", "Supervision Ordered #15", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K398").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Supervision Ordered #15", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K400").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Supervision Ordered #15", "Supervision Programs", "AGGREGATES") & userRow).value)

   '#16
    PrintSheet.Range("D410").value = Lookup("Supervision_Program_Num")(DataSheet.Range(hFind("Supervision Ordered #16", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("G410").value = DataSheet.Range(hFind("Start Date", "Supervision Ordered #16", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G412").value = DataSheet.Range(hFind("End Date", "Supervision Ordered #16", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G414").value = DataSheet.Range(hFind("LOS", "Supervision Ordered #16", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("E416").value = DataSheet.Range(hFind("Supervision Description", "Supervision Ordered #16", "Supervision Programs", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Supervision Ordered #16", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("K410").value = Lookup("Community_Based_Supervision_Provider_Num")(DataSheet.Range(hFind("Community-Based Agency", "Supervision Ordered #16", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("N410").value = Lookup("Residential_Supervision_Provider_Num")(DataSheet.Range(hFind("Residential Agency", "Supervision Ordered #16", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K412").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Supervision Ordered #16", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K414").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Supervision Ordered #16", "Supervision Programs", "AGGREGATES") & userRow).value)

    '#17
    PrintSheet.Range("D424").value = Lookup("Supervision_Program_Num")(DataSheet.Range(hFind("Supervision Ordered #17", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("G424").value = DataSheet.Range(hFind("Start Date", "Supervision Ordered #17", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G426").value = DataSheet.Range(hFind("End Date", "Supervision Ordered #17", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G428").value = DataSheet.Range(hFind("LOS", "Supervision Ordered #17", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("E430").value = DataSheet.Range(hFind("Supervision Description", "Supervision Ordered #17", "Supervision Programs", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Supervision Ordered #17", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("K424").value = Lookup("Community_Based_Supervision_Provider_Num")(DataSheet.Range(hFind("Community-Based Agency", "Supervision Ordered #17", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("N424").value = Lookup("Residential_Supervision_Provider_Num")(DataSheet.Range(hFind("Residential Agency", "Supervision Ordered #17", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K426").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Supervision Ordered #17", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K428").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Supervision Ordered #17", "Supervision Programs", "AGGREGATES") & userRow).value)
    
    '#18
    PrintSheet.Range("D438").value = Lookup("Supervision_Program_Num")(DataSheet.Range(hFind("Supervision Ordered #18", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("G438").value = DataSheet.Range(hFind("Start Date", "Supervision Ordered #18", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G440").value = DataSheet.Range(hFind("End Date", "Supervision Ordered #18", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G442").value = DataSheet.Range(hFind("LOS", "Supervision Ordered #18", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("E444").value = DataSheet.Range(hFind("Supervision Description", "Supervision Ordered #18", "Supervision Programs", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Supervision Ordered #18", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("K438").value = Lookup("Community_Based_Supervision_Provider_Num")(DataSheet.Range(hFind("Community-Based Agency", "Supervision Ordered #18", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("N438").value = Lookup("Residential_Supervision_Provider_Num")(DataSheet.Range(hFind("Residential Agency", "Supervision Ordered #18", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K440").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Supervision Ordered #18", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K442").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Supervision Ordered #18", "Supervision Programs", "AGGREGATES") & userRow).value)
    
    '#19
    PrintSheet.Range("D452").value = Lookup("Supervision_Program_Num")(DataSheet.Range(hFind("Supervision Ordered #19", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("G452").value = DataSheet.Range(hFind("Start Date", "Supervision Ordered #19", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G454").value = DataSheet.Range(hFind("End Date", "Supervision Ordered #19", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G456").value = DataSheet.Range(hFind("LOS", "Supervision Ordered #19", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("E458").value = DataSheet.Range(hFind("Supervision Description", "Supervision Ordered #19", "Supervision Programs", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Supervision Ordered #19", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("K452").value = Lookup("Community_Based_Supervision_Provider_Num")(DataSheet.Range(hFind("Community-Based Agency", "Supervision Ordered #19", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("N452").value = Lookup("Residential_Supervision_Provider_Num")(DataSheet.Range(hFind("Residential Agency", "Supervision Ordered #19", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K454").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Supervision Ordered #19", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K456").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Supervision Ordered #19", "Supervision Programs", "AGGREGATES") & userRow).value)

    '#20
    PrintSheet.Range("D466").value = Lookup("Supervision_Program_Num")(DataSheet.Range(hFind("Supervision Ordered #20", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("G466").value = DataSheet.Range(hFind("Start Date", "Supervision Ordered #20", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G468").value = DataSheet.Range(hFind("End Date", "Supervision Ordered #20", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G470").value = DataSheet.Range(hFind("LOS", "Supervision Ordered #20", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("E472").value = DataSheet.Range(hFind("Supervision Description", "Supervision Ordered #20", "Supervision Programs", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Supervision Ordered #20", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("K466").value = Lookup("Community_Based_Supervision_Provider_Num")(DataSheet.Range(hFind("Community-Based Agency", "Supervision Ordered #20", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("N466").value = Lookup("Residential_Supervision_Provider_Num")(DataSheet.Range(hFind("Residential Agency", "Supervision Ordered #20", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K468").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Supervision Ordered #20", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K470").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Supervision Ordered #20", "Supervision Programs", "AGGREGATES") & userRow).value)
   
    
    Call YouthSearchPrint3

End Sub


Sub YouthSearchPrint3()


Call RefreshNamedRanges
    Call Generate_Dictionaries
    
    Dim PrintSheet As Worksheet
    Dim DataSheet As Worksheet
    Set PrintSheet = Worksheets("Youth Search")
    Set DataSheet = Worksheets("Entry")
    
    Dim userRow As Long
    userRow = PrintSheet.Range("J5").value


'Conditions HISTORY
    '#1
    PrintSheet.Range("S200").value = Lookup("Condition_Num")(DataSheet.Range(hFind("Condition Ordered #1", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("W200").value = DataSheet.Range(hFind("Start Date", "Condition Ordered #1", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("W202").value = DataSheet.Range(hFind("End Date", "Condition Ordered #1", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("W204").value = DataSheet.Range(hFind("LOS", "Condition Ordered #1", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("U206").value = DataSheet.Range(hFind("Condition Description", "Condition Ordered #1", "Conditions", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Condition Ordered #1", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("AA200").value = Lookup("Condition_Provider_Num")(DataSheet.Range(hFind("Condition Agency", "Condition Ordered #1", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA202").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Condition Ordered #1", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA204").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Condition Ordered #1", "Conditions", "AGGREGATES") & userRow).value)

    '#2
    PrintSheet.Range("S214").value = Lookup("Condition_Num")(DataSheet.Range(hFind("Condition Ordered #2", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("W214").value = DataSheet.Range(hFind("Start Date", "Condition Ordered #2", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("W216").value = DataSheet.Range(hFind("End Date", "Condition Ordered #2", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("W218").value = DataSheet.Range(hFind("LOS", "Condition Ordered #2", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("U220").value = DataSheet.Range(hFind("Condition Description", "Condition Ordered #2", "Conditions", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Condition Ordered #2", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("AA214").value = Lookup("Condition_Provider_Num")(DataSheet.Range(hFind("Condition Agency", "Condition Ordered #2", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA216").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Condition Ordered #2", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA218").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Condition Ordered #2", "Conditions", "AGGREGATES") & userRow).value)
    
    '#3
    PrintSheet.Range("S228").value = Lookup("Condition_Num")(DataSheet.Range(hFind("Condition Ordered #3", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("W228").value = DataSheet.Range(hFind("Start Date", "Condition Ordered #3", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("W230").value = DataSheet.Range(hFind("End Date", "Condition Ordered #3", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("W232").value = DataSheet.Range(hFind("LOS", "Condition Ordered #3", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("U234").value = DataSheet.Range(hFind("Condition Description", "Condition Ordered #3", "Conditions", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Condition Ordered #3", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("AA228").value = Lookup("Condition_Provider_Num")(DataSheet.Range(hFind("Condition Agency", "Condition Ordered #3", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA230").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Condition Ordered #3", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA232").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Condition Ordered #3", "Conditions", "AGGREGATES") & userRow).value)

    '#4
    PrintSheet.Range("S242").value = Lookup("Condition_Num")(DataSheet.Range(hFind("Condition Ordered #4", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("W242").value = DataSheet.Range(hFind("Start Date", "Condition Ordered #4", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("W244").value = DataSheet.Range(hFind("End Date", "Condition Ordered #4", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("W246").value = DataSheet.Range(hFind("LOS", "Condition Ordered #4", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("U248").value = DataSheet.Range(hFind("Condition Description", "Condition Ordered #4", "Conditions", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Condition Ordered #4", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("AA242").value = Lookup("Condition_Provider_Num")(DataSheet.Range(hFind("Condition Agency", "Condition Ordered #4", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA244").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Condition Ordered #4", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA246").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Condition Ordered #4", "Conditions", "AGGREGATES") & userRow).value)

    '#5
    PrintSheet.Range("S256").value = Lookup("Condition_Num")(DataSheet.Range(hFind("Condition Ordered #5", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("W256").value = DataSheet.Range(hFind("Start Date", "Condition Ordered #5", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("W258").value = DataSheet.Range(hFind("End Date", "Condition Ordered #5", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("W260").value = DataSheet.Range(hFind("LOS", "Condition Ordered #5", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("U262").value = DataSheet.Range(hFind("Condition Description", "Condition Ordered #5", "Conditions", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Condition Ordered #5", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("AA256").value = Lookup("Condition_Provider_Num")(DataSheet.Range(hFind("Condition Agency", "Condition Ordered #5", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA258").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Condition Ordered #5", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA260").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Condition Ordered #5", "Conditions", "AGGREGATES") & userRow).value)

    '#6
    PrintSheet.Range("S270").value = Lookup("Condition_Num")(DataSheet.Range(hFind("Condition Ordered #6", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("W270").value = DataSheet.Range(hFind("Start Date", "Condition Ordered #6", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("W272").value = DataSheet.Range(hFind("End Date", "Condition Ordered #6", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("W274").value = DataSheet.Range(hFind("LOS", "Condition Ordered #6", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("U276").value = DataSheet.Range(hFind("Condition Description", "Condition Ordered #6", "Conditions", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Condition Ordered #6", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("AA270").value = Lookup("Condition_Provider_Num")(DataSheet.Range(hFind("Condition Agency", "Condition Ordered #6", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA272").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Condition Ordered #6", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA274").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Condition Ordered #6", "Conditions", "AGGREGATES") & userRow).value)


    '#7
    PrintSheet.Range("S284").value = Lookup("Condition_Num")(DataSheet.Range(hFind("Condition Ordered #7", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("V284").value = DataSheet.Range(hFind("Start Date", "Condition Ordered #7", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V286").value = DataSheet.Range(hFind("End Date", "Condition Ordered #7", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V288").value = DataSheet.Range(hFind("LOS", "Condition Ordered #7", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("U290").value = DataSheet.Range(hFind("Condition Description", "Condition Ordered #7", "Conditions", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Condition Ordered #7", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("AA284").value = Lookup("Condition_Provider_Num")(DataSheet.Range(hFind("Condition Agency", "Condition Ordered #7", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA286").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Condition Ordered #7", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA288").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Condition Ordered #7", "Conditions", "AGGREGATES") & userRow).value)

    '#8
    PrintSheet.Range("S298").value = Lookup("Condition_Num")(DataSheet.Range(hFind("Condition Ordered #8", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("V298").value = DataSheet.Range(hFind("Start Date", "Condition Ordered #8", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V300").value = DataSheet.Range(hFind("End Date", "Condition Ordered #8", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V302").value = DataSheet.Range(hFind("LOS", "Condition Ordered #8", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("U304").value = DataSheet.Range(hFind("Condition Description", "Condition Ordered #8", "Conditions", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Condition Ordered #8", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("AA298").value = Lookup("Condition_Provider_Num")(DataSheet.Range(hFind("Condition Agency", "Condition Ordered #8", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA300").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Condition Ordered #8", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA302").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Condition Ordered #8", "Conditions", "AGGREGATES") & userRow).value)

    '#9
    PrintSheet.Range("S312").value = Lookup("Condition_Num")(DataSheet.Range(hFind("Condition Ordered #9", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("V312").value = DataSheet.Range(hFind("Start Date", "Condition Ordered #9", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V314").value = DataSheet.Range(hFind("End Date", "Condition Ordered #9", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V316").value = DataSheet.Range(hFind("LOS", "Condition Ordered #9", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("U318").value = DataSheet.Range(hFind("Condition Description", "Condition Ordered #9", "Conditions", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Condition Ordered #9", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("AA312").value = Lookup("Condition_Provider_Num")(DataSheet.Range(hFind("Condition Agency", "Condition Ordered #9", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA314").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Condition Ordered #9", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA316").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Condition Ordered #9", "Conditions", "AGGREGATES") & userRow).value)

    '#10
    PrintSheet.Range("S326").value = Lookup("Condition_Num")(DataSheet.Range(hFind("Condition Ordered #10", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("V326").value = DataSheet.Range(hFind("Start Date", "Condition Ordered #10", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V328").value = DataSheet.Range(hFind("End Date", "Condition Ordered #10", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V330").value = DataSheet.Range(hFind("LOS", "Condition Ordered #10", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("U332").value = DataSheet.Range(hFind("Condition Description", "Condition Ordered #10", "Conditions", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Condition Ordered #10", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("AA326").value = Lookup("Condition_Provider_Num")(DataSheet.Range(hFind("Condition Agency", "Condition Ordered #10", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA328").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Condition Ordered #10", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA330").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Condition Ordered #10", "Conditions", "AGGREGATES") & userRow).value)


    '#11
    PrintSheet.Range("S340").value = Lookup("Condition_Num")(DataSheet.Range(hFind("Condition Ordered #11", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("V340").value = DataSheet.Range(hFind("Start Date", "Condition Ordered #11", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V342").value = DataSheet.Range(hFind("End Date", "Condition Ordered #11", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V344").value = DataSheet.Range(hFind("LOS", "Condition Ordered #11", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("U346").value = DataSheet.Range(hFind("Condition Description", "Condition Ordered #11", "Conditions", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Condition Ordered #11", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("AA340").value = Lookup("Condition_Provider_Num")(DataSheet.Range(hFind("Condition Agency", "Condition Ordered #11", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA342").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Condition Ordered #11", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA344").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Condition Ordered #11", "Conditions", "AGGREGATES") & userRow).value)

    '#12
    PrintSheet.Range("S354").value = Lookup("Condition_Num")(DataSheet.Range(hFind("Condition Ordered #12", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("V354").value = DataSheet.Range(hFind("Start Date", "Condition Ordered #12", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V356").value = DataSheet.Range(hFind("End Date", "Condition Ordered #12", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V358").value = DataSheet.Range(hFind("LOS", "Condition Ordered #12", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("U360").value = DataSheet.Range(hFind("Condition Description", "Condition Ordered #12", "Conditions", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Condition Ordered #12", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("AA354").value = Lookup("Condition_Provider_Num")(DataSheet.Range(hFind("Condition Agency", "Condition Ordered #12", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA356").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Condition Ordered #12", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA358").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Condition Ordered #12", "Conditions", "AGGREGATES") & userRow).value)

    '#13
    PrintSheet.Range("S368").value = Lookup("Condition_Num")(DataSheet.Range(hFind("Condition Ordered #13", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("V368").value = DataSheet.Range(hFind("Start Date", "Condition Ordered #13", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V370").value = DataSheet.Range(hFind("End Date", "Condition Ordered #13", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V372").value = DataSheet.Range(hFind("LOS", "Condition Ordered #13", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("U374").value = DataSheet.Range(hFind("Condition Description", "Condition Ordered #13", "Conditions", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Condition Ordered #13", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("AA368").value = Lookup("Condition_Provider_Num")(DataSheet.Range(hFind("Condition Agency", "Condition Ordered #13", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA370").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Condition Ordered #13", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA372").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Condition Ordered #13", "Conditions", "AGGREGATES") & userRow).value)

    '#14
    PrintSheet.Range("S382").value = Lookup("Condition_Num")(DataSheet.Range(hFind("Condition Ordered #14", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("V382").value = DataSheet.Range(hFind("Start Date", "Condition Ordered #14", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V384").value = DataSheet.Range(hFind("End Date", "Condition Ordered #14", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V386").value = DataSheet.Range(hFind("LOS", "Condition Ordered #14", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("U388").value = DataSheet.Range(hFind("Condition Description", "Condition Ordered #14", "Conditions", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Condition Ordered #14", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("AA382").value = Lookup("Condition_Provider_Num")(DataSheet.Range(hFind("Condition Agency", "Condition Ordered #14", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA384").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Condition Ordered #14", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA386").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Condition Ordered #14", "Conditions", "AGGREGATES") & userRow).value)

    '#15
    PrintSheet.Range("S396").value = Lookup("Condition_Num")(DataSheet.Range(hFind("Condition Ordered #15", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("V396").value = DataSheet.Range(hFind("Start Date", "Condition Ordered #15", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V398").value = DataSheet.Range(hFind("End Date", "Condition Ordered #15", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V400").value = DataSheet.Range(hFind("LOS", "Condition Ordered #15", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("U402").value = DataSheet.Range(hFind("Condition Description", "Condition Ordered #15", "Conditions", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Condition Ordered #15", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("AA396").value = Lookup("Condition_Provider_Num")(DataSheet.Range(hFind("Condition Agency", "Condition Ordered #15", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA398").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Condition Ordered #15", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA400").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Condition Ordered #15", "Conditions", "AGGREGATES") & userRow).value)

    '#16
    PrintSheet.Range("S410").value = Lookup("Condition_Num")(DataSheet.Range(hFind("Condition Ordered #16", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("V410").value = DataSheet.Range(hFind("Start Date", "Condition Ordered #16", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V412").value = DataSheet.Range(hFind("End Date", "Condition Ordered #16", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V414").value = DataSheet.Range(hFind("LOS", "Condition Ordered #16", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("U416").value = DataSheet.Range(hFind("Condition Description", "Condition Ordered #16", "Conditions", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Condition Ordered #16", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("AA410").value = Lookup("Condition_Provider_Num")(DataSheet.Range(hFind("Condition Agency", "Condition Ordered #16", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA412").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Condition Ordered #16", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA414").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Condition Ordered #16", "Conditions", "AGGREGATES") & userRow).value)

    '#17
    PrintSheet.Range("S424").value = Lookup("Condition_Num")(DataSheet.Range(hFind("Condition Ordered #17", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("V424").value = DataSheet.Range(hFind("Start Date", "Condition Ordered #17", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V426").value = DataSheet.Range(hFind("End Date", "Condition Ordered #17", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V428").value = DataSheet.Range(hFind("LOS", "Condition Ordered #17", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("U430").value = DataSheet.Range(hFind("Condition Description", "Condition Ordered #17", "Conditions", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Condition Ordered #17", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("AA424").value = Lookup("Condition_Provider_Num")(DataSheet.Range(hFind("Condition Agency", "Condition Ordered #17", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA426").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Condition Ordered #17", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA428").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Condition Ordered #17", "Conditions", "AGGREGATES") & userRow).value)

    '#18
    PrintSheet.Range("S438").value = Lookup("Condition_Num")(DataSheet.Range(hFind("Condition Ordered #18", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("V438").value = DataSheet.Range(hFind("Start Date", "Condition Ordered #18", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V440").value = DataSheet.Range(hFind("End Date", "Condition Ordered #18", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V442").value = DataSheet.Range(hFind("LOS", "Condition Ordered #18", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("U444").value = DataSheet.Range(hFind("Condition Description", "Condition Ordered #18", "Conditions", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Condition Ordered #18", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("AA438").value = Lookup("Condition_Provider_Num")(DataSheet.Range(hFind("Condition Agency", "Condition Ordered #18", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA440").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Condition Ordered #18", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA442").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Condition Ordered #18", "Conditions", "AGGREGATES") & userRow).value)

    '#19
    PrintSheet.Range("S452").value = Lookup("Condition_Num")(DataSheet.Range(hFind("Condition Ordered #19", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("V452").value = DataSheet.Range(hFind("Start Date", "Condition Ordered #19", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V454").value = DataSheet.Range(hFind("End Date", "Condition Ordered #19", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V456").value = DataSheet.Range(hFind("LOS", "Condition Ordered #19", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("U458").value = DataSheet.Range(hFind("Condition Description", "Condition Ordered #19", "Conditions", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Condition Ordered #19", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("AA452").value = Lookup("Condition_Provider_Num")(DataSheet.Range(hFind("Condition Agency", "Condition Ordered #19", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA454").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Condition Ordered #19", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA456").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Condition Ordered #19", "Conditions", "AGGREGATES") & userRow).value)
    
    '#20
    PrintSheet.Range("S466").value = Lookup("Condition_Num")(DataSheet.Range(hFind("Condition Ordered #20", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("V466").value = DataSheet.Range(hFind("Start Date", "Condition Ordered #20", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V468").value = DataSheet.Range(hFind("End Date", "Condition Ordered #20", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V470").value = DataSheet.Range(hFind("LOS", "Condition Ordered #20", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("U472").value = DataSheet.Range(hFind("Condition Description", "Condition Ordered #20", "Conditions", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Condition Ordered #20", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("K466").value = Lookup("Condition_Provider_Num")(DataSheet.Range(hFind("Condition Agency", "Condition Ordered #20", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA468").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Condition Ordered #20", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA470").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Condition Ordered #20", "Conditions", "AGGREGATES") & userRow).value)

Call YouthSearchPrint4

End Sub

    
   Sub YouthSearchPrint4()


Call RefreshNamedRanges
    Call Generate_Dictionaries
    
    Dim PrintSheet As Worksheet
    Dim DataSheet As Worksheet
    Set PrintSheet = Worksheets("Youth Search")
    Set DataSheet = Worksheets("Entry")
    
    Dim userRow As Long
    userRow = PrintSheet.Range("J5").value
    
    
    
    'COURT LISTINGS HISTORY
    '#1
    PrintSheet.Range("K487").value = DataSheet.Range(hFind("Court Date #1", "LISTINGS") & userRow).value
    PrintSheet.Range("K489").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #1", "LISTINGS") & userRow).value)
    PrintSheet.Range("P487").value = DataSheet.Range(hFind("Notes", "Court Date #1", "LISTINGS") & userRow).value

    '#2
    PrintSheet.Range("K495").value = DataSheet.Range(hFind("Court Date #2", "LISTINGS") & userRow).value
    PrintSheet.Range("K497").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #2", "LISTINGS") & userRow).value)
    PrintSheet.Range("P495").value = DataSheet.Range(hFind("Notes", "Court Date #2", "LISTINGS") & userRow).value
    
    '#3
    PrintSheet.Range("K503").value = DataSheet.Range(hFind("Court Date #3", "LISTINGS") & userRow).value
    PrintSheet.Range("K505").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #3", "LISTINGS") & userRow).value)
    PrintSheet.Range("P503").value = DataSheet.Range(hFind("Notes", "Court Date #3", "LISTINGS") & userRow).value
    
    '#4
    PrintSheet.Range("K511").value = DataSheet.Range(hFind("Court Date #4", "LISTINGS") & userRow).value
    PrintSheet.Range("K513").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #4", "LISTINGS") & userRow).value)
    PrintSheet.Range("P511").value = DataSheet.Range(hFind("Notes", "Court Date #4", "LISTINGS") & userRow).value
    
    '#5
    PrintSheet.Range("K519").value = DataSheet.Range(hFind("Court Date #5", "LISTINGS") & userRow).value
    PrintSheet.Range("K521").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #5", "LISTINGS") & userRow).value)
    PrintSheet.Range("P519").value = DataSheet.Range(hFind("Notes", "Court Date #5", "LISTINGS") & userRow).value
    
    '#6
    PrintSheet.Range("K527").value = DataSheet.Range(hFind("Court Date #6", "LISTINGS") & userRow).value
    PrintSheet.Range("K529").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #6", "LISTINGS") & userRow).value)
    PrintSheet.Range("P527").value = DataSheet.Range(hFind("Notes", "Court Date #6", "LISTINGS") & userRow).value
    
    '#7
    PrintSheet.Range("K535").value = DataSheet.Range(hFind("Court Date #7", "LISTINGS") & userRow).value
    PrintSheet.Range("K537").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #7", "LISTINGS") & userRow).value)
    PrintSheet.Range("P535").value = DataSheet.Range(hFind("Notes", "Court Date #7", "LISTINGS") & userRow).value
    
    '#8
    PrintSheet.Range("K543").value = DataSheet.Range(hFind("Court Date #8", "LISTINGS") & userRow).value
    PrintSheet.Range("K545").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #8", "LISTINGS") & userRow).value)
    PrintSheet.Range("P543").value = DataSheet.Range(hFind("Notes", "Court Date #8", "LISTINGS") & userRow).value
    
    '#9
    PrintSheet.Range("K551").value = DataSheet.Range(hFind("Court Date #9", "LISTINGS") & userRow).value
    PrintSheet.Range("K553").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #9", "LISTINGS") & userRow).value)
    PrintSheet.Range("P551").value = DataSheet.Range(hFind("Notes", "Court Date #9", "LISTINGS") & userRow).value
    
    '#10
    PrintSheet.Range("K559").value = DataSheet.Range(hFind("Court Date #10", "LISTINGS") & userRow).value
    PrintSheet.Range("K561").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #10", "LISTINGS") & userRow).value)
    PrintSheet.Range("P559").value = DataSheet.Range(hFind("Notes", "Court Date #10", "LISTINGS") & userRow).value
    
    '#11
    PrintSheet.Range("K567").value = DataSheet.Range(hFind("Court Date #11", "LISTINGS") & userRow).value
    PrintSheet.Range("K569").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #11", "LISTINGS") & userRow).value)
    PrintSheet.Range("P567").value = DataSheet.Range(hFind("Notes", "Court Date #11", "LISTINGS") & userRow).value
    
    '#12
    PrintSheet.Range("K575").value = DataSheet.Range(hFind("Court Date #12", "LISTINGS") & userRow).value
    PrintSheet.Range("K577").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #12", "LISTINGS") & userRow).value)
    PrintSheet.Range("P575").value = DataSheet.Range(hFind("Notes", "Court Date #12", "LISTINGS") & userRow).value
    
    '#13
    PrintSheet.Range("K583").value = DataSheet.Range(hFind("Court Date #13", "LISTINGS") & userRow).value
    PrintSheet.Range("K585").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #13", "LISTINGS") & userRow).value)
    PrintSheet.Range("P583").value = DataSheet.Range(hFind("Notes", "Court Date #13", "LISTINGS") & userRow).value
    
    '#14
    PrintSheet.Range("K591").value = DataSheet.Range(hFind("Court Date #14", "LISTINGS") & userRow).value
    PrintSheet.Range("K593").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #14", "LISTINGS") & userRow).value)
    PrintSheet.Range("P591").value = DataSheet.Range(hFind("Notes", "Court Date #14", "LISTINGS") & userRow).value
    
    '#15
    PrintSheet.Range("K599").value = DataSheet.Range(hFind("Court Date #15", "LISTINGS") & userRow).value
    PrintSheet.Range("K601").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #15", "LISTINGS") & userRow).value)
    PrintSheet.Range("P599").value = DataSheet.Range(hFind("Notes", "Court Date #15", "LISTINGS") & userRow).value
    
    '#16
    PrintSheet.Range("K607").value = DataSheet.Range(hFind("Court Date #16", "LISTINGS") & userRow).value
    PrintSheet.Range("K609").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #16", "LISTINGS") & userRow).value)
    PrintSheet.Range("P607").value = DataSheet.Range(hFind("Notes", "Court Date #16", "LISTINGS") & userRow).value
    
    '#17
    PrintSheet.Range("K615").value = DataSheet.Range(hFind("Court Date #17", "LISTINGS") & userRow).value
    PrintSheet.Range("K617").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #17", "LISTINGS") & userRow).value)
    PrintSheet.Range("P615").value = DataSheet.Range(hFind("Notes", "Court Date #17", "LISTINGS") & userRow).value
    
    '#18
    PrintSheet.Range("K623").value = DataSheet.Range(hFind("Court Date #18", "LISTINGS") & userRow).value
    PrintSheet.Range("K625").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #18", "LISTINGS") & userRow).value)
    PrintSheet.Range("P623").value = DataSheet.Range(hFind("Notes", "Court Date #18", "LISTINGS") & userRow).value
    
    '#19
    PrintSheet.Range("K631").value = DataSheet.Range(hFind("Court Date #19", "LISTINGS") & userRow).value
    PrintSheet.Range("K633").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #19", "LISTINGS") & userRow).value)
    PrintSheet.Range("P631").value = DataSheet.Range(hFind("Notes", "Court Date #19", "LISTINGS") & userRow).value
    
    '#20
    PrintSheet.Range("K639").value = DataSheet.Range(hFind("Court Date #20", "LISTINGS") & userRow).value
    PrintSheet.Range("K641").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #20", "LISTINGS") & userRow).value)
    PrintSheet.Range("P639").value = DataSheet.Range(hFind("Notes", "Court Date #20", "LISTINGS") & userRow).value

    '#21
    PrintSheet.Range("K647").value = DataSheet.Range(hFind("Court Date #21", "LISTINGS") & userRow).value
    PrintSheet.Range("K649").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #21", "LISTINGS") & userRow).value)
    PrintSheet.Range("P647").value = DataSheet.Range(hFind("Notes", "Court Date #21", "LISTINGS") & userRow).value

    '#22
    PrintSheet.Range("K655").value = DataSheet.Range(hFind("Court Date #22", "LISTINGS") & userRow).value
    PrintSheet.Range("K657").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #22", "LISTINGS") & userRow).value)
    PrintSheet.Range("P655").value = DataSheet.Range(hFind("Notes", "Court Date #22", "LISTINGS") & userRow).value

    '#23
    PrintSheet.Range("K663").value = DataSheet.Range(hFind("Court Date #23", "LISTINGS") & userRow).value
    PrintSheet.Range("K665").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #23", "LISTINGS") & userRow).value)
    PrintSheet.Range("P663").value = DataSheet.Range(hFind("Notes", "Court Date #23", "LISTINGS") & userRow).value

    '#24
    PrintSheet.Range("K671").value = DataSheet.Range(hFind("Court Date #24", "LISTINGS") & userRow).value
    PrintSheet.Range("K673").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #24", "LISTINGS") & userRow).value)
    PrintSheet.Range("P671").value = DataSheet.Range(hFind("Notes", "Court Date #24", "LISTINGS") & userRow).value

    '#25
    PrintSheet.Range("K679").value = DataSheet.Range(hFind("Court Date #25", "LISTINGS") & userRow).value
    PrintSheet.Range("K681").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #25", "LISTINGS") & userRow).value)
    PrintSheet.Range("P679").value = DataSheet.Range(hFind("Notes", "Court Date #25", "LISTINGS") & userRow).value
    
    '#26
    PrintSheet.Range("K687").value = DataSheet.Range(hFind("Court Date #26", "LISTINGS") & userRow).value
    PrintSheet.Range("K689").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #26", "LISTINGS") & userRow).value)
    PrintSheet.Range("P687").value = DataSheet.Range(hFind("Notes", "Court Date #26", "LISTINGS") & userRow).value

    '#27
    PrintSheet.Range("K695").value = DataSheet.Range(hFind("Court Date #27", "LISTINGS") & userRow).value
    PrintSheet.Range("K697").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #27", "LISTINGS") & userRow).value)
    PrintSheet.Range("P695").value = DataSheet.Range(hFind("Notes", "Court Date #27", "LISTINGS") & userRow).value
    
    '#28
    PrintSheet.Range("K703").value = DataSheet.Range(hFind("Court Date #28", "LISTINGS") & userRow).value
    PrintSheet.Range("K705").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #28", "LISTINGS") & userRow).value)
    PrintSheet.Range("P703").value = DataSheet.Range(hFind("Notes", "Court Date #28", "LISTINGS") & userRow).value

    '#29
    PrintSheet.Range("K711").value = DataSheet.Range(hFind("Court Date #29", "LISTINGS") & userRow).value
    PrintSheet.Range("K713").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #29", "LISTINGS") & userRow).value)
    PrintSheet.Range("P711").value = DataSheet.Range(hFind("Notes", "Court Date #29", "LISTINGS") & userRow).value
    
    '#30
    PrintSheet.Range("K719").value = DataSheet.Range(hFind("Court Date #30", "LISTINGS") & userRow).value
    PrintSheet.Range("K721").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #30", "LISTINGS") & userRow).value)
    PrintSheet.Range("P719").value = DataSheet.Range(hFind("Notes", "Court Date #30", "LISTINGS") & userRow).value




    End Sub
