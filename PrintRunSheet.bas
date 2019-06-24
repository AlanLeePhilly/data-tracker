Attribute VB_Name = "PrintRunSheet"
Sub PrintRunSheet()

    Call RefreshNamedRanges
    Call Generate_Dictionaries

    Dim InputSheet As Worksheet
    Dim PrintSheet As Worksheet
    Dim DataSheet As Worksheet
    Set InputSheet = Worksheets("Youth Search")
    Set PrintSheet = Worksheets("Run Sheet")
    Set DataSheet = Worksheets("Entry")

    Dim d As String
    d = InputSheet.Range("F5").value
    If StrComp(d, "") = 0 Then
        Dim a As String
        Exit Sub
    End If

    PrintSheet.Cells(1, 1) = "Printing run sheet for court date: " & d
    PrintSheet.Cells(3, 1) = "Last Name"
    PrintSheet.Cells(3, 2) = "First Name"
    PrintSheet.Cells(3, 3) = "Courtroom"
    PrintSheet.Cells(3, 4) = "LoS Courtroom"
    PrintSheet.Cells(3, 5) = "Listing Type"
    PrintSheet.Cells(3, 6) = "DoB"
    PrintSheet.Cells(3, 7) = "Age"
    PrintSheet.Cells(3, 8) = "Petition #1"
    PrintSheet.Cells(3, 9) = "Petition #2"
    PrintSheet.Cells(3, 10) = "Legal Status"
    PrintSheet.Cells(3, 11) = "LoS Legal Status"
    PrintSheet.Cells(3, 12) = "Supervision"
    PrintSheet.Cells(3, 13) = "LoS Supervision"

    Dim currCount As Long
    currCount = 3

    'Set up column lookups
    Dim lastNameCol As String
    lastNameCol = headerFind("Last Name")
    Dim firstNameCol As String
    firstNameCol = headerFind("First Name")
    Dim listingTypeCol As String
    listingTypeCol = headerFind("Listing Type")
    Dim dobCol As String
    dobCol = headerFind("DOB")
    Dim petition1Col As String
    petition1Col = headerFind("Petition #1")
    Dim petition2Col As String
    petition2Col = headerFind("Petition #2")

    Dim col As Long
    Dim lastr As Long
    col = DataSheet.Cells(3, hFind("Next Court Date")).Column
    lastr = DataSheet.Cells.Find(What:="*", After:=DataSheet.Cells(1, col + 1), LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False).row
    
    For i = 3 To lastr
        If DataSheet.Cells(i, hFind("Next Court Date")).value = d Then
            currCount = currCount + 1
            PrintSheet.Cells(currCount, 1).value = DataSheet.Cells(i, lastNameCol).value
            PrintSheet.Cells(currCount, 2).value = DataSheet.Cells(i, firstNameCol).value

            'Looks like Listing_Type is not set up properly as a mapping somehow. It is not recognized
            'MsgBox Lookup("Listing_Type")(DataSheet.Range(listingTypeCol & i).value)
            'PrintSheet.Cells(currCount, 5).value = Lookup("Listing_Type")(DataSheet.Cells(i, listingTypeCol).value)


            PrintSheet.Cells(currCount, 6).value = DataSheet.Cells(i, dobCol).value
            PrintSheet.Cells(currCount, 8).value = DataSheet.Cells(i, petition1Col).value
            PrintSheet.Cells(currCount, 9).value = DataSheet.Cells(i, petition2Col).value


            'Courtroom
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
            'Courtroom = findFirstValue(DataSheet, userRow, "4G", courtroomOptions, "Start Date", "End Date")
            'PrintSheet.Range("G48") = Courtroom

            'Courtroom LoS
            'If Not StrComp(Courtroom, "") = 0 Then
            'losCourtroom = DateDiff("d", DataSheet.Range(hFind("Start Date", Courtroom, "4G") & userRow).value, _
            'VBA.Format(Now(), "mm/dd/yyyy"))
            'PrintSheet.Range("J48") = losCourtroom & " days"
            'End If
        End If
    Next i

End Sub
