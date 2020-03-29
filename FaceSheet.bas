Attribute VB_Name = "FaceSheet"
Sub FaceSheetPrint0()

    Dim PrintSheet As Worksheet
    Set PrintSheet = Worksheets("Face Sheet")
    
    'Basic details
    PrintSheet.Range("A7").Select
    Selection.ClearContents
    PrintSheet.Range("L7").Select
    Selection.ClearContents
    PrintSheet.Range("C11").Select
    Selection.ClearContents
    PrintSheet.Range("H11").Select
    Selection.ClearContents
    PrintSheet.Range("L11").Select
    Selection.ClearContents
    PrintSheet.Range("R11").Select
    Selection.ClearContents
    PrintSheet.Range("V11").Select
    Selection.ClearContents
    PrintSheet.Range("L13").Select
    Selection.ClearContents
    PrintSheet.Range("R13").Select
    Selection.ClearContents
    PrintSheet.Range("V13").Select
    Selection.ClearContents
    
    'Current supervisions
    PrintSheet.Range("N17").Select
    Selection.ClearContents
    PrintSheet.Range("U17").Select
    Selection.ClearContents
    PrintSheet.Range("X17").Select
    Selection.ClearContents
    PrintSheet.Range("N19").Select
    Selection.ClearContents
    PrintSheet.Range("U19").Select
    Selection.ClearContents
    PrintSheet.Range("X19").Select
    Selection.ClearContents
    PrintSheet.Range("N21").Select
    Selection.ClearContents
    PrintSheet.Range("U21").Select
    Selection.ClearContents
    PrintSheet.Range("X21").Select
    Selection.ClearContents
    PrintSheet.Range("N23").Select
    Selection.ClearContents
    PrintSheet.Range("U23").Select
    Selection.ClearContents
    PrintSheet.Range("X23").Select
    Selection.ClearContents
    
    'Current conditions
    PrintSheet.Range("N32").Select
    Selection.ClearContents
    PrintSheet.Range("U32").Select
    Selection.ClearContents
    PrintSheet.Range("X32").Select
    Selection.ClearContents
    PrintSheet.Range("N34").Select
    Selection.ClearContents
    PrintSheet.Range("U34").Select
    Selection.ClearContents
    PrintSheet.Range("X34").Select
    Selection.ClearContents
    PrintSheet.Range("N36").Select
    Selection.ClearContents
    PrintSheet.Range("U36").Select
    Selection.ClearContents
    PrintSheet.Range("X36").Select
    Selection.ClearContents
    PrintSheet.Range("N38").Select
    Selection.ClearContents
    PrintSheet.Range("U38").Select
    Selection.ClearContents
    PrintSheet.Range("X38").Select
    Selection.ClearContents
    PrintSheet.Range("N40").Select
    Selection.ClearContents
    PrintSheet.Range("U40").Select
    Selection.ClearContents
    PrintSheet.Range("X40").Select
    Selection.ClearContents
    
    'Most recent listing
    PrintSheet.Range("C39").Select
    Selection.ClearContents
    PrintSheet.Range("I39").Select
    Selection.ClearContents
    PrintSheet.Range("I40").Select
    Selection.ClearContents
    PrintSheet.Range("B41").Select
    Selection.ClearContents
    'Update placeholder for DA once available
    
    'Most recent supervision
    PrintSheet.Range("C53").Select
    Selection.ClearContents
    PrintSheet.Range("C55").Select
    Selection.ClearContents
    PrintSheet.Range("C57").Select
    Selection.ClearContents
    PrintSheet.Range("C59").Select
    Selection.ClearContents
    PrintSheet.Range("H53").Select
    Selection.ClearContents
    PrintSheet.Range("H55").Select
    Selection.ClearContents
    PrintSheet.Range("H57").Select
    Selection.ClearContents
    PrintSheet.Range("J53").Select
    Selection.ClearContents
    PrintSheet.Range("B61").Select
    Selection.ClearContents
    PrintSheet.Range("H59").Select
    Selection.ClearContents
    
    'Most recent condition
    PrintSheet.Range("C67").Select
    Selection.ClearContents
    PrintSheet.Range("C69").Select
    Selection.ClearContents
    PrintSheet.Range("C71").Select
    Selection.ClearContents
    PrintSheet.Range("C73").Select
    Selection.ClearContents
    PrintSheet.Range("I67").Select
    Selection.ClearContents
    PrintSheet.Range("I69").Select
    Selection.ClearContents
    PrintSheet.Range("I71").Select
    Selection.ClearContents
    PrintSheet.Range("B75").Select
    Selection.ClearContents
    PrintSheet.Range("I73").Select
    Selection.ClearContents
    
    'Demographics
    PrintSheet.Range("M52").Select
    Selection.ClearContents
    PrintSheet.Range("M54").Select
    Selection.ClearContents
    PrintSheet.Range("S52").Select
    Selection.ClearContents
    PrintSheet.Range("S54").Select
    Selection.ClearContents
    PrintSheet.Range("S52").Select
    Selection.ClearContents
    
    'Arrest info
    PrintSheet.Range("M58").Select
    Selection.ClearContents
    PrintSheet.Range("N60").Select
    Selection.ClearContents
    PrintSheet.Range("N62").Select
    Selection.ClearContents
    PrintSheet.Range("N64").Select
    Selection.ClearContents
    PrintSheet.Range("N66").Select
    Selection.ClearContents
    PrintSheet.Range("Q58").Select
    Selection.ClearContents
    
    'Petition info
    PrintSheet.Range("L70").Select
    Selection.ClearContents
    PrintSheet.Range("L72").Select
    Selection.ClearContents
    PrintSheet.Range("S70").Select
    Selection.ClearContents
    PrintSheet.Range("S72").Select
    Selection.ClearContents
    PrintSheet.Range("X70").Select
    Selection.ClearContents
    PrintSheet.Range("L75").Select
    Selection.ClearContents
    PrintSheet.Range("L77").Select
    Selection.ClearContents
    PrintSheet.Range("S75").Select
    Selection.ClearContents
    PrintSheet.Range("S77").Select
    Selection.ClearContents
    PrintSheet.Range("X76").Select
    Selection.ClearContents
    
    'Recent Supervision History
    PrintSheet.Range("U80").Select
    Selection.ClearContents
    Dim i As Integer
    Dim j As Integer
    j = 81
    Do Until j > 172
        i = 1
        Do Until i > 7
            PrintSheet.Range("C" & j + i).Select
            Selection.ClearContents
            PrintSheet.Range("I" & j + i).Select
            Selection.ClearContents
            
            PrintSheet.Range("N" & j + i).Select
            Selection.ClearContents
            PrintSheet.Range("U" & j + i).Select
            Selection.ClearContents
            
            i = i + 2
        Loop
        PrintSheet.Range("K" & j + 1).Select
        Selection.ClearContents
        PrintSheet.Range("X" & j + 1).Select
        Selection.ClearContents
        PrintSheet.Range("B" & j + 9).Select
        Selection.ClearContents
        PrintSheet.Range("M" & j + 9).Select
        Selection.ClearContents
        
        j = j + 13
    Loop
    
    PrintSheet.Range("G5").Select
    
    Call FaceSheetPrint

End Sub

Sub FaceSheetPrint()

    Call RefreshNamedRanges
    Call Generate_Dictionaries
    
    Dim PrintSheet As Worksheet
    Dim DataSheet As Worksheet
    Set PrintSheet = Worksheets("Face Sheet")
    Set DataSheet = Worksheets("Entry")
    
    Dim userRow As Long
    userRow = PrintSheet.Range("F5").value
    
    'Basic details
    PrintSheet.Range("A7").value = DataSheet.Range(hFind("Last Name") & userRow).value & ", " & DataSheet.Range(hFind("First Name") & userRow).value
    PrintSheet.Range("L7").value = "Petition #: " & DataSheet.Range(hFind("Petition #1") & userRow).value
    PrintSheet.Range("C11").value = DataSheet.Range(hFind("Next Court Date") & userRow).value
    
    Dim activeStatus As String
    activeStatus = Lookup("Active_Num")(DataSheet.Range(hFind("Active or Discharged (in courtroom)?") & userRow).value)
    PrintSheet.Range("H11").value = activeStatus
    
    If StrComp(activeStatus, "Active") = 0 Then
        'LoS for petition
        Dim losPetition As Integer
    
        Dim petitionDate As String
        petitionDate = DataSheet.Range(hFind("Date Filed", "Petition") & userRow).value
        losPetition = DateDiff("d", petitionDate, VBA.format(Now(), "mm/dd/yyyy"))
        PrintSheet.Range("L11").value = losPetition & " days"
        
        PrintSheet.Range("L13").value = Lookup("Listing_Type_Num")(DataSheet.Range(hFind("Listing Type", "DEMOGRAPHICS") & userRow).value)
        
        Dim Courtroom As String
        Dim losCourtroom As Integer
        Dim courtroomOptions(1 To 6) As String
        courtroomOptions(1) = "4G"
        courtroomOptions(2) = "4E"
        courtroomOptions(3) = "6F"
        courtroomOptions(4) = "6H"
        courtroomOptions(5) = "3E"
        courtroomOptions(6) = "ADULT"
        Courtroom = findFirstValue(DataSheet, userRow, "4G", courtroomOptions, "Start Date", "End Date")
        PrintSheet.Range("R11") = Courtroom

        'if courtroom exists
        If Not StrComp(Courtroom, "") = 0 Then
            losCourtroom = DateDiff("d", DataSheet.Range(hFind("Start Date", Courtroom, "4G") & userRow).value, _
                VBA.format(Now(), "mm/dd/yyyy"))
            PrintSheet.Range("V11") = losCourtroom & " days"

        Else
            Dim SpecialtyCourtroom As String
            Dim losSpecialtyCourtroom As Integer
            Dim SpecialtyCourtroomOptions(1 To 3) As String
            SpecialtyCourtroomOptions(1) = "Crossover"
            SpecialtyCourtroomOptions(2) = "WRAP"
            SpecialtyCourtroomOptions(3) = "JTC"
            SpecialtyCourtroom = findFirstValue(DataSheet, userRow, "Crossover", SpecialtyCourtroomOptions, "Referral Date", "Date of Overall Discharge")
            PrintSheet.Range("R11") = SpecialtyCourtroom

            'if courtroom exists
            If Not StrComp(SpecialtyCourtroom, "") = 0 Then
                losSpecialtyCourtroom = DateDiff("d", DataSheet.Range(hFind("Referral Date", SpecialtyCourtroom, "Crossover") & userRow).value, _
                VBA.format(Now(), "mm/dd/yyyy"))
                PrintSheet.Range("V11") = losSpecialtyCourtroom & " days"

            End If
        End If

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
        PrintSheet.Range("R13") = legalStatus

        'if legal status exists
        If Not StrComp(legalStatus, "") = 0 Then
            losLegalStatus = DateDiff("d", DataSheet.Range(hFind("Start Date", legalStatus, "Aggregates") & userRow).value, _
            VBA.format(Now(), "mm/dd/yyyy"))
            PrintSheet.Range("V13") = losLegalStatus & " days"

        End If

        Dim SpecialtyLegalStatus As String
        Dim losSpecialtyLegalStatus As Integer
        Dim lostSpecialtyLegalStatus As Integer
        Dim SpecialtyLegalStatusOptions(1 To 3) As String
        SpecialtyLegalStatusOptions(1) = "Crossover"
        SpecialtyLegalStatusOptions(2) = "WRAP"
        SpecialtyLegalStatusOptions(3) = "JTC"

        'Find if youth has "Accepted Date" to courtroom and then return that he/she is on that legal status
        SpecialtyLegalStatus = findFirstValue(DataSheet, userRow, "Crossover", SpecialtyLegalStatusOptions, "Accepted Date", "Date of Overall Discharge")
        PrintSheet.Range("R13") = SpecialtyLegalStatus

        'If youth as been accepted to that courtroom, date-diff acceptance date to show LOS of specialty court legal status
        If Not StrComp(SpecialtyLegalStatus, "") = 0 Then
            losSpecialtyLegalStatus = DateDiff("d", DataSheet.Range(hFind("Accepted Date", SpecialtyLegalStatus) & userRow).value, _
            VBA.format(Now(), "mm/dd/yyyy"))
            PrintSheet.Range("V13") = losSpecialtyLegalStatus & " days"

        Else

            'If youth is in specialty courtroom and accepted date is still "0" (ie, has not yet been accepted),
            'print legal status as "Pretrial" - LOS Pretrial will get handled by standard LOS sweep)

            PrintSheet.Range("R13") = "Pretrial"
            
        End If
        
        'Current supervisions
        Dim i As Integer
        Dim bucketHead As String
        Dim printRow As Long
        Dim programType As String
        Dim providerName As String
        Dim thing1 As String, thing2 As String

        printRow = 17

        For i = 1 To 30
            bucketHead = hFind("Supervision Ordered #" & i, "AGGREGATES")

            thing1 = DataSheet.Range(headerFind("Start Date", bucketHead) & userRow).value
            thing2 = DataSheet.Range(headerFind("End Date", bucketHead) & userRow).value

            If isNotEmptyOrZero(DataSheet.Range(headerFind("Start Date", bucketHead) & userRow)) _
              And isEmptyOrZero(DataSheet.Range(headerFind("End Date", bucketHead) & userRow)) Then

                programType = Lookup("Supervision_Program_Num")(DataSheet.Range(bucketHead & userRow).value)

                If isResidential(programType) Then
                    providerName = Lookup("Residential_Supervision_Provider_Num")(DataSheet.Range(headerFind("Residential Agency", bucketHead) & userRow).value)
                Else
                    providerName = Lookup("Community_Based_Supervision_Provider_Num")(DataSheet.Range(headerFind("Community-Based Agency", bucketHead) & userRow).value)
                End If

                If printRow > 21 Then
                    MsgBox "The following Active Supervision will not be printed due to space constraints: " _
                        & vbNewLine & "Program: " & programType _
                        & vbNewLine & "Provider: " & providerName _
                        & vbNewLine & "Start Date: " & DataSheet.Range(headerFind("Start Date", bucketHead) & userRow).value
                Else
                    PrintSheet.Range("N" & printRow) = programType
                    PrintSheet.Range("U" & printRow) = providerName
                    PrintSheet.Range("X" & printRow) = DateDiff("d", DataSheet.Range(headerFind("Start Date", bucketHead) & userRow).value, Date) & " days"
                End If
                
                printRow = printRow + 2
                
            End If

        Next i
        
        Dim j As Integer
        j = 17
        Do Until j > 21
            If IsEmpty(PrintSheet.Range("N" & j)) Then
                PrintSheet.Range("N" & j).value = "None"
                PrintSheet.Range("U" & j).value = "N/A"
                PrintSheet.Range("X" & j).value = "N/A"
            End If
    
            j = j + 2
        Loop
        
        'Current conditions
        Dim k As Integer
        Dim bucketHead1 As String
        'Dim printRow As Long
        Dim ConditionType As String
        Dim conditionproviderName As String
        'Dim thing2 As String, thing2 As String

        printRow = 32

        For k = 1 To 20
            bucketHead1 = hFind("Condition Ordered #" & k, "AGGREGATES")

            thing1 = DataSheet.Range(headerFind("Start Date", bucketHead1) & userRow).value
            thing2 = DataSheet.Range(headerFind("End Date", bucketHead1) & userRow).value

            If isNotEmptyOrZero(DataSheet.Range(headerFind("Start Date", bucketHead1) & userRow)) _
              And isEmptyOrZero(DataSheet.Range(headerFind("End Date", bucketHead1) & userRow)) Then

                ConditionType = Lookup("Condition_Num")(DataSheet.Range(bucketHead1 & userRow).value)
                conditionproviderName = Lookup("Condition_Provider_Num")(DataSheet.Range(headerFind("Condition Agency", bucketHead1) & userRow).value)
            
            If printRow > 40 Then
                    MsgBox "The following Active Condition will not be printed due to space constraints: " _
                        & vbNewLine & "Program: " & ConditionType _
                        & vbNewLine & "Provider: " & conditionproviderName _
                        & vbNewLine & "Start Date: " & DataSheet.Range(headerFind("Start Date", bucketHead1) & userRow).value
               
                Else
                    PrintSheet.Range("N" & printRow) = ConditionType
                    PrintSheet.Range("U" & printRow) = conditionproviderName
                    PrintSheet.Range("X" & printRow) = DateDiff("d", DataSheet.Range(headerFind("Start Date", bucketHead1) & userRow).value, Date) & " days"
                End If
                
                printRow = printRow + 2
                
           End If

        Next k
        
        Dim l As Integer
        l = 32
        Do Until l > 40
            If IsEmpty(PrintSheet.Range("N" & l)) Then
                PrintSheet.Range("N" & l).value = "None"
                PrintSheet.Range("U" & l).value = "N/A"
                PrintSheet.Range("X" & l).value = "N/A"
            End If
    
            l = l + 2
        Loop
        
        End If
        
        'Conditions
        'Dim conditionsColumns() As String
        'conditionsColumns = findAllValues(DataSheet, userRow, "Aggregates", "Condition Ordered", "Start Date", "End Date")
        'Dim conditionArrLength As Integer

        'If (Not conditionsColumns) = -1 Then
            'conditionArrLength = 0
        'Else
            'conditionArrLength = UBound(conditionsColumns) - LBound(conditionsColumns) + 1
        'End If
        'Dim conditionI As Integer
        'Dim conditionStart As String

        'For conditionI = 1 To conditionArrLength
            'PrintSheet.Range("N" & 30 + 2 * conditionI) = Lookup("Condition_Num")(DataSheet.Range(hFind(conditionsColumns(conditionI - 1), "Aggregates") & userRow).value)
            'PrintSheet.Range("U" & 30 + 2 * conditionI) = Lookup("Condition_Provider_Num")(DataSheet.Range(hFind("Condition Agency", conditionsColumns(conditionI - 1), "Aggregates") & userRow).value)
            
            'conditionStart = DataSheet.Range(hFind("Start Date", conditionsColumns(conditionI - 1), "Aggregates") & userRow).value
            'PrintSheet.Range("X" & 30 + 2 * conditionI) = DateDiff("d", conditionStart, VBA.format(Now(), "mm/dd/yyyy")) & " days"
            'If conditionI = 6 Then Exit For
        'Next conditionI
        
        'Dim k As Integer
        'k = 32
        'Do Until k > 40
            'If IsEmpty(PrintSheet.Range("N" & k)) Then
                'PrintSheet.Range("N" & k).value = "None"
                'PrintSheet.Range("U" & k).value = "N/A"
                'PrintSheet.Range("X" & k).value = "N/A"
            'End If
            
            'k = k + 2
        'Loop
    
    'End If
    
    'Most recent listing
    If Not IsEmpty(DataSheet.Range(hFind("Court Date #1", "LISTINGS") & userRow).value) Then
        Dim lastListing As Integer
        lastListing = 1
        Do Until IsEmpty(DataSheet.Range(hFind("Court Date #" & lastListing, "LISTINGS") & userRow).value)
            lastListing = lastListing + 1
        Loop
        
        lastListing = lastListing - 1
        
        PrintSheet.Range("C39").value = DataSheet.Range(hFind("Court Date #" & lastListing, "LISTINGS") & userRow).value
        PrintSheet.Range("I39").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #" & lastListing, "LISTINGS") & userRow).value)
        PrintSheet.Range("C40").value = Lookup("DA_Last_Name_Num")(DataSheet.Range(hFind("DA", "Court Date #" & lastListing, "LISTINGS") & userRow).value)
        PrintSheet.Range("I40").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status", "Court Date #" & lastListing, "LISTINGS") & userRow).value)
        PrintSheet.Range("B41").value = DataSheet.Range(hFind("Notes", "Court Date #" & lastListing, "LISTINGS") & userRow).value
    End If
    'Update placeholder for DA once available
    
    'Most recent supervision
    If Not IsEmpty(DataSheet.Range(hFind("Supervision Ordered #1", "Supervision Programs", "AGGREGATES") & userRow).value) Then
    
        Dim lastSup As Integer
        lastSup = 1
        Do Until IsEmpty(DataSheet.Range(hFind("Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value)
            lastSup = lastSup + 1
            If lastSup = 31 Then
                Exit Do
            End If
        Loop
        
        lastSup = lastSup - 1
        
        PrintSheet.Range("C53").value = Lookup("Supervision_Program_Num")(DataSheet.Range(hFind("Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value)
        PrintSheet.Range("C55").value = DataSheet.Range(hFind("Start Date", "Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value
        PrintSheet.Range("C57").value = DataSheet.Range(hFind("End Date", "Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value
        PrintSheet.Range("C59").value = DataSheet.Range(hFind("LOS", "Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value
        PrintSheet.Range("B61").value = DataSheet.Range(hFind("Supervision Description", "Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value
        PrintSheet.Range("H53").value = Lookup("Community_Based_Supervision_Provider_Num")(DataSheet.Range(hFind("Community-Based Agency", "Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value)
        PrintSheet.Range("J53").value = Lookup("Residential_Supervision_Provider_Num")(DataSheet.Range(hFind("Residential Agency", "Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value)
        PrintSheet.Range("H55").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value)
        PrintSheet.Range("H57").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value)
        PrintSheet.Range("H59").value = Lookup("DA_Last_Name_Num")(DataSheet.Range(hFind("DA", "Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value)
    End If
    
    'Most recent condition
    If Not IsEmpty(DataSheet.Range(hFind("Condition Ordered #1", "Conditions", "AGGREGATES") & userRow).value) Then
    
        Dim lastCon As Integer
        lastCon = 1
        Do Until IsEmpty(DataSheet.Range(hFind("Condition Ordered #" & lastCon, "Conditions", "AGGREGATES") & userRow).value)
            lastCon = lastCon + 1
            If lastCon = 21 Then
                Exit Do
            End If
        Loop
        
        lastCon = lastCon - 1
        
        PrintSheet.Range("C67").value = Lookup("Condition_Num")(DataSheet.Range(hFind("Condition Ordered #" & lastCon, "Conditions", "AGGREGATES") & userRow).value)
        PrintSheet.Range("C69").value = DataSheet.Range(hFind("Start Date", "Condition Ordered #" & lastCon, "Conditions", "AGGREGATES") & userRow).value
        PrintSheet.Range("C71").value = DataSheet.Range(hFind("End Date", "Condition Ordered #" & lastCon, "Conditions", "AGGREGATES") & userRow).value
        PrintSheet.Range("C73").value = DataSheet.Range(hFind("LOS", "Condition Ordered #" & lastCon, "Conditions", "AGGREGATES") & userRow).value
        PrintSheet.Range("B75").value = DataSheet.Range(hFind("Condition Description", "Condition Ordered #" & lastCon, "Conditions", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Condition Ordered #" & lastCon, "Conditions", "AGGREGATES") & userRow).value
        PrintSheet.Range("I67").value = Lookup("Condition_Provider_Num")(DataSheet.Range(hFind("Condition Agency", "Condition Ordered #" & lastCon, "Conditions", "AGGREGATES") & userRow).value)
        PrintSheet.Range("I69").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Condition Ordered #" & lastCon, "Conditions", "AGGREGATES") & userRow).value)
        PrintSheet.Range("I71").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Condition Ordered #" & lastCon, "Conditions", "AGGREGATES") & userRow).value)
        PrintSheet.Range("I73").value = Lookup("DA_Last_Name_Num")(DataSheet.Range(hFind("DA", "Condition Ordered #" & lastCon, "Conditions", "AGGREGATES") & userRow).value)
    End If
    
    'Demographics
    PrintSheet.Range("M52").value = DataSheet.Range(hFind("DOB") & userRow).value
    PrintSheet.Range("M54").value = DataSheet.Range(hFind("School") & userRow).value
    PrintSheet.Range("W54").value = DataSheet.Range(hFind("Grade") & userRow).value
    PrintSheet.Range("S52").value = ageAtTime(VBA.format(Now(), "mm/dd/yyyy"), userRow)
    PrintSheet.Range("W52").value = DataSheet.Range(hFind("Age @ Intake") & userRow).value
    
    'Arrest Info
    PrintSheet.Range("M58").value = DataSheet.Range(hFind("DC #") & userRow).value
    PrintSheet.Range("N60").value = DataSheet.Range(hFind("Incident District") & userRow).value
    PrintSheet.Range("N62").value = DataSheet.Range(hFind("Arrest Date") & userRow).value
    PrintSheet.Range("N64").value = Lookup("Generic_YNOU_Num")(DataSheet.Range(hFind("Active in System at Time of Arrest?", "Petition") & userRow).value)
    PrintSheet.Range("N66").value = DataSheet.Range(hFind("# of Prior Arrests") & userRow).value
    PrintSheet.Range("P58").value = DataSheet.Range(hFind("General Notes from Intake") & userRow).value
    
    'Petition Info
    PrintSheet.Range("L70").value = DataSheet.Range(hFind("Petition #1") & userRow).value
    PrintSheet.Range("S70").value = DataSheet.Range(hFind("Lead Charge Name", "Petition #1") & userRow).value
    PrintSheet.Range("X70").value = Lookup("Charge_Grade_Specific_Num")(DataSheet.Range(hFind("Charge Grade (specific) #1", "Petition #1") & userRow).value)
    PrintSheet.Range("L72").value = DataSheet.Range(hFind("Date Filed", "Petition #1") & userRow).value
    
    PrintSheet.Range("L75").value = DataSheet.Range(hFind("Petition #2") & userRow).value
    PrintSheet.Range("S75").value = DataSheet.Range(hFind("Lead Charge Name", "Petition #2") & userRow).value
    PrintSheet.Range("X75").value = Lookup("Charge_Grade_Specific_Num")(DataSheet.Range(hFind("Charge Grade (specific) #1", "Petition #2") & userRow).value)
    PrintSheet.Range("L77").value = DataSheet.Range(hFind("Date Filed", "Petition #2") & userRow).value
    
    
    'Recent Listing History
    
    If Not IsEmpty(DataSheet.Range(hFind("Court Date #1", "LISTINGS") & userRow).value) Then
        PrintSheet.Range("G80").value = lastListing
        
        Dim sheetRow As Integer
        sheetRow = 82
        
        lastListing = lastListing - 1
        
        
        
        
        'Dim lastListing As Integer
        'lastListing = 1
        'Do Until IsEmpty(DataSheet.Range(hFind("Court Date #" & lastListing, "LISTINGS") & userRow).value)
            'lastListing = lastListing + 1
        'Loop
        
        'lastListing = lastListing - 1
        
        Do Until lastListing = 0 Or sheetRow > 279
            If sheetRow < 174 Then
            
        PrintSheet.Cells(sheetRow, 3).value = DataSheet.Range(hFind("Court Date #" & lastListing, "LISTINGS") & userRow).value
        PrintSheet.Cells(sheetRow + 2, 3).value = Lookup("DA_Last_Name_Num")(DataSheet.Range(hFind("DA", "Court Date #" & lastListing, "LISTINGS") & userRow).value)
        PrintSheet.Cells(sheetRow + 4, 2).value = DataSheet.Range(hFind("Notes", "Court Date #" & lastListing, "LISTINGS") & userRow).value
        PrintSheet.Cells(sheetRow, 9).value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #" & lastListing, "LISTINGS") & userRow).value)
        PrintSheet.Cells(sheetRow + 2, 9).value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status", "Court Date #" & lastListing, "LISTINGS") & userRow).value)
        Else
        'PrintSheet.Cells(sheetRow - 104, 14).value = Lookup("Supervision_Program_Num")(DataSheet.Range(hFind("Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value)
        'PrintSheet.Cells(sheetRow - 102, 14).value = DataSheet.Range(hFind("Start Date", "Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value
        'PrintSheet.Cells(sheetRow - 100, 14).value = DataSheet.Range(hFind("End Date", "Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value
        'PrintSheet.Cells(sheetRow - 98, 14).value = DataSheet.Range(hFind("LOS", "Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value
        'PrintSheet.Cells(sheetRow - 96, 13).value = DataSheet.Range(hFind("Supervision Description", "Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value
        'PrintSheet.Cells(sheetRow - 104, 21).value = Lookup("Community_Based_Supervision_Provider_Num")(DataSheet.Range(hFind("Community-Based Agency", "Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value)
        'PrintSheet.Cells(sheetRow - 104, 24).value = Lookup("Residential_Supervision_Provider_Num")(DataSheet.Range(hFind("Residential Agency", "Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value)
        'PrintSheet.Cells(sheetRow - 102, 21).value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value)
        'PrintSheet.Cells(sheetRow - 100, 21).value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value)
        'PrintSheet.Cells(sheetRow - 98, 21).value = Lookup("DA_Last_Name_Num")(DataSheet.Range(hFind("DA", "Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value)
    End If
    
    sheetRow = sheetRow + 13
    
    lastListing = lastListing - 1
            If lastListing = 0 Then
                Exit Do
            End If
     Loop
    
    End If
    'Recent Supervision History
    If Not IsEmpty(DataSheet.Range(hFind("Supervision Ordered #1", "Supervision Programs", "AGGREGATES") & userRow).value) Then
        PrintSheet.Range("U80").value = lastSup
        
    Dim sheetRow1 As Integer
        sheetRow1 = 82
        
        lastSup = lastSup - 1
        
        Do Until lastSup = 0 Or sheetRow1 > 279
            If sheetRow1 < 174 Then
                PrintSheet.Cells(sheetRow1, 14).value = Lookup("Supervision_Program_Num")(DataSheet.Range(hFind("Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value)
                PrintSheet.Cells(sheetRow1 + 2, 14).value = DataSheet.Range(hFind("Start Date", "Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value
                PrintSheet.Cells(sheetRow1 + 4, 14).value = DataSheet.Range(hFind("End Date", "Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value
                PrintSheet.Cells(sheetRow1 + 6, 14).value = DataSheet.Range(hFind("LOS", "Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value
                PrintSheet.Cells(sheetRow1 + 8, 13).value = DataSheet.Range(hFind("Supervision Description", "Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value
                PrintSheet.Cells(sheetRow1, 21).value = Lookup("Community_Based_Supervision_Provider_Num")(DataSheet.Range(hFind("Community-Based Agency", "Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value)
                PrintSheet.Cells(sheetRow1, 24).value = Lookup("Residential_Supervision_Provider_Num")(DataSheet.Range(hFind("Residential Agency", "Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value)
                PrintSheet.Cells(sheetRow1 + 2, 21).value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value)
                PrintSheet.Cells(sheetRow1 + 4, 21).value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value)
                PrintSheet.Cells(sheetRow + 6, 21).value = Lookup("DA_Last_Name_Num")(DataSheet.Range(hFind("DA", "Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value)
    Else
    
    End If
    
        
        
    'Recent Supervision History
    'If Not IsEmpty(DataSheet.Range(hFind("Supervision Ordered #1", "Supervision Programs", "AGGREGATES") & userRow).value) Then
        'PrintSheet.Range("U80").value = lastSup
        
        'Dim sheetRow1 As Integer
        'sheetRow1 = 82
        
        'lastSup = lastSup - 1
        
        'Do Until lastSup = 0 Or sheetRow1 > 279
            'If sheetRow1 < 174 Then
                'PrintSheet.Cells(sheetRow, 14).value = Lookup("Supervision_Program_Num")(DataSheet.Range(hFind("Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value)
                'PrintSheet.Cells(sheetRow + 2, 3).value = DataSheet.Range(hFind("Start Date", "Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value
                'PrintSheet.Cells(sheetRow + 4, 3).value = DataSheet.Range(hFind("End Date", "Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value
                'PrintSheet.Cells(sheetRow + 6, 3).value = DataSheet.Range(hFind("LOS", "Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value
                'PrintSheet.Cells(sheetRow + 8, 2).value = DataSheet.Range(hFind("Supervision Description", "Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value
                'PrintSheet.Cells(sheetRow, 9).value = Lookup("Community_Based_Supervision_Provider_Num")(DataSheet.Range(hFind("Community-Based Agency", "Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value)
                'PrintSheet.Cells(sheetRow, 11).value = Lookup("Residential_Supervision_Provider_Num")(DataSheet.Range(hFind("Residential Agency", "Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value)
                'PrintSheet.Cells(sheetRow + 2, 9).value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value)
                'PrintSheet.Cells(sheetRow + 4, 9).value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value)
                'PrintSheet.Cells(sheetRow + 6, 9).value = Lookup("DA_Last_Name_Num")(DataSheet.Range(hFind("DA", "Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value)
            'Else
                'PrintSheet.Cells(sheetRow - 104, 14).value = Lookup("Supervision_Program_Num")(DataSheet.Range(hFind("Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value)
                'PrintSheet.Cells(sheetRow - 102, 14).value = DataSheet.Range(hFind("Start Date", "Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value
                'PrintSheet.Cells(sheetRow - 100, 14).value = DataSheet.Range(hFind("End Date", "Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value
                'PrintSheet.Cells(sheetRow - 98, 14).value = DataSheet.Range(hFind("LOS", "Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value
                'PrintSheet.Cells(sheetRow - 96, 13).value = DataSheet.Range(hFind("Supervision Description", "Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value
                'PrintSheet.Cells(sheetRow - 104, 21).value = Lookup("Community_Based_Supervision_Provider_Num")(DataSheet.Range(hFind("Community-Based Agency", "Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value)
                'PrintSheet.Cells(sheetRow - 104, 24).value = Lookup("Residential_Supervision_Provider_Num")(DataSheet.Range(hFind("Residential Agency", "Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value)
                'PrintSheet.Cells(sheetRow - 102, 21).value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value)
                'PrintSheet.Cells(sheetRow - 100, 21).value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value)
                'PrintSheet.Cells(sheetRow - 98, 21).value = Lookup("DA_Last_Name_Num")(DataSheet.Range(hFind("DA", "Supervision Ordered #" & lastSup, "Supervision Programs", "AGGREGATES") & userRow).value)
            'End If
            
            sheetRow1 = sheetRow1 + 13
            
            lastSup = lastSup - 1
            If lastSup = 0 Then
                Exit Do
           End If
        Loop
    
    Else
        PrintSheet.Range("N79").value = "0"
    End If

End Sub

Sub PrintFaceSheet()
    
    Dim FaceSheet As Worksheet
    Dim DataSheet As Worksheet
    Dim EntrySheet As Worksheet
    Set FaceSheet = Worksheets("Face Sheet")
    Set DataSheet = Worksheets("Run Sheet")
    Set EntrySheet = Worksheets("Entry")

    Dim x As Integer
    Dim r As Range
    
    Application.ScreenUpdating = False
    
    'Find start of Petition #1 column on run sheet
    Set r = DataSheet.Cells.Find("Petition #1")
    DataSheet.Activate
    r.Select
    r.Activate
    
    'set number of rows of data
    NumRows = DataSheet.Range(r, ActiveCell.End(xlDown)).Rows.count - 1
    
    For x = 1 To NumRows
        'Get Petition #1 for current kid
        DataSheet.Activate
        r.Select
        r.Activate
        ActiveCell.Offset(x, 0).Select
        
        'Load data into Face Sheet
        Dim rfind As Range
        Set rfind = EntrySheet.Cells.Find("Petition #1")
        FaceSheet.Range("F5").value = EntrySheet.Cells.Find(ActiveCell.value, After:=rfind, SearchOrder:=xlColumns).row
        FaceSheet.Activate
        Run ("FaceSheetPrint0")
        
        FaceSheet.PrintOut
        
    Next x
    
    Application.ScreenUpdating = True

End Sub





