Attribute VB_Name = "Module07_Testing"
Sub Test_Print()

    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With

    Worksheets("Entry").Activate
    
    Worksheets("TestOutput").Range("A2:X2000").ClearContents
    Dim userRow As Long
    Dim i As Long
    Dim j As Long
    Dim dicName As String
    Dim value As Long
    Call RefreshNamedRanges
    Call Generate_Dictionaries

    'marker for what ROW of the test output we're writing to
    j = 2


    k = 1

    'Grab the row number that was written in G1 of the "TestOutput" sheet
    userRow = Worksheets("TestOutput").Range("G1").value
    Call AggAggSupervisionsAndConditions(userRow)

    With Worksheets("Entry")

        'i represents column index, so we're starting at column "B" and moving right
        For i = 2 To alphaToNum(headerFind("END"))

            'if the userRow has content at the column we're talking about
            If Not IsEmpty(Worksheets("Entry").Range(numToAlpha(i) & userRow)) Then

                'write the column letter in row 2, column A of the "TestOutput" row we're currently at
                Worksheets("TestOutput").Cells(j, 1).value = numToAlpha(i) 'column letter

                'write the header name in row 2, column A of the "TestOutput" row we're currently at
                Worksheets("TestOutput").Cells(j, 2).value = Worksheets("Entry").Cells(2, i).value 'header name


                'if there's a Lookup List value in ROW1 of the column we're talking about
                If isNotEmptyOrZero(Worksheets("Entry").Range(numToAlpha(i) & 1)) Then

                    'make the name of the lookup dictionary by appending "_Num" to the value in ROW1
                    dicName = Worksheets("Entry").Range(numToAlpha(i) & 1).value & "_Num"

                    'pull the raw value (enum number) from the cell
                    value = Worksheets("Entry").Cells(userRow, i).value

                    'set color to easily identify these translation printouts
                    Worksheets("TestOutput").Range("A" & j & ":E" & j).Interior.Color = RGB(240, 240, 240)

                    'print the translated value to column "D"
                    Worksheets("TestOutput").Cells(j, 3).value = Lookup(dicName)(value)

                    'print the raw value from the datasheet to column "C"
                    Worksheets("TestOutput").Cells(j, 4).value = Worksheets("Entry").Cells(userRow, i).value
                    Worksheets("TestOutput").Cells(j, 4).NumberFormat = "General"

                    'print the lookup dictionary name
                    Worksheets("TestOutput").Cells(j, 5).value = Worksheets("Entry").Cells(1, i).value

                    'if there's not a lookup list value in ROW1 (meaning that this will be a direct value entry
                Else

                    'print the value from the datasheet to column "C"
                    Worksheets("TestOutput").Cells(j, 3).value = Worksheets("Entry").Cells(userRow, i).value

                    'make sure this row is not colored (resetting in case it was colored on last render
                    Worksheets("TestOutput").Cells(j, 3).EntireRow.Interior.Color = xlNone
                End If

                'move on to next printrow
                j = j + 1

                'if userRow does not have content at this column
            Else

                'look at the header name for this column
                Select Case Worksheets("Entry").Range(numToAlpha(i) & 2).value

                        'if it's the name of one of our big section headers
                    Case "DEMOGRAPHICS", "PETITION", "DRAI", "INTAKE CONFERENCE", "DETENTION", "DETENTION (VOP)", "DIVERSION", "4G", "4E", "6F", "6H", "3E", "Crossover", "WRAP", "JTC", "ADULT", "AGGREGATES", "LISTINGS", "JUVENILE PETITION", "ADULT PETITION"

                        'if the prior row was also one of the big section headers
                        If Worksheets("TestOutput").Cells(j - 1, 2).Interior.Color = RGB(200, 200, 200) Then

                            'print "... Empty Section" and make the row uncolored
                            Worksheets("TestOutput").Cells(j, 1).value = "..."
                            Worksheets("TestOutput").Cells(j, 2).value = "Empty Section"
                            Worksheets("TestOutput").Cells(j, 3).EntireRow.Interior.Color = xlNone

                            'move on to next printrow
                            j = j + 1
                        End If

                        'print column letter to column "A"
                        Worksheets("TestOutput").Cells(j, 1).value = numToAlpha(i) 'column letter

                        'print header name to column "B"
                        Worksheets("TestOutput").Cells(j, 2).value = Worksheets("Entry").Cells(2, i).value 'header name

                        'print color the row darkly
                        Worksheets("TestOutput").Range("A" & j & ":E" & j).Interior.Color = RGB(200, 200, 200)

                        'move on to next printrow
                        j = j + 1
                End Select
            End If

            'move on to next column on datasheet
        Next i
    End With
    
     With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
    End With
    Worksheets("TestOutput").Activate
End Sub

Sub Test_Write_Client_1()

    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With

    Dim i As Long
    Dim col As String



    Worksheets("Entry").Range("C2:ARJ2").ClearContents

    For i = 2 To 1000
        If Not IsEmpty(Worksheets("TestUsers").Cells(i, 1)) Then
            col = Worksheets("TestUsers").Cells(i, 1).value
            Worksheets("Entry").Range(col & 2).value = Worksheets("TestUsers").Cells(i, 3).value
        End If
    Next i

    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
    End With
End Sub

Sub Test_Compare_Client_1()
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With
    Call RefreshNamedRanges
    Call Generate_Dictionaries
    Dim col As String
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim switch As Long
    Dim rng As Range
    Dim TestCopy As Worksheet
    Dim DataSheet As Worksheet
    Set TestCopy = Worksheets("TestCopy")
    Set DataSheet = Worksheets("Entry")
    Dim hold As String
    Dim hold2 As Long
    Dim userRow As Long
    Dim dataVal As String

    Worksheets("TestCopy").Range("H2:L200").ClearContents

    userRow = Worksheets("TestCopy").Range("G1").value

    k = 2
    switch = 0
    For i = 2 To alphaToNum(headerFind("END"))
        'if the cell has content
        If Not (Trim(DataSheet.Cells(userRow, i).value) = "") Then

            'grab column name
            col = numToAlpha(i)

            switch = 0

            For j = 2 To 200
                If TestCopy.Cells(j, 1).value = col Then
                    switch = 1
                    If Not IsEmpty(DataSheet.Cells(1, i).value) Then
                        dataVal = Lookup(DataSheet.Cells(1, i).value & "_Num")(DataSheet.Cells(userRow, i).value)
                    Else
                        dataVal = DataSheet.Cells(userRow, i).value
                    End If

                    If Not TestCopy.Cells(j, 3).value = dataVal Then
                        TestCopy.Cells(k, 9).value = col
                        TestCopy.Cells(k, 10).value = TestCopy.Cells(j, 2).value
                        TestCopy.Cells(k, 11).value = TestCopy.Cells(j, 3).value
                        TestCopy.Cells(k, 12).value = DataSheet.Cells(userRow, i).value
                        If Not IsEmpty(DataSheet.Cells(1, i).value) Then
                            TestCopy.Cells(k, 13).value = Lookup(DataSheet.Cells(1, i).value & "_Num")(DataSheet.Cells(userRow, i).value)
                            TestCopy.Cells(k, 14).value = DataSheet.Cells(1, i).value
                        End If
                        TestCopy.Cells(k, 11).EntireRow.Interior.Color = xlNone
                        k = k + 1
                    End If
                Else

                    If IsEmpty(TestCopy.Cells(j, 1)) And switch = 0 Then
                        switch = 1
                        TestCopy.Cells(k, 9).value = col
                        TestCopy.Cells(k, 10).value = DataSheet.Cells(2, i).value
                        TestCopy.Cells(k, 11).value = "(blank)"
                        TestCopy.Cells(k, 12).value = DataSheet.Cells(userRow, i).value
                        If Not IsEmpty(DataSheet.Cells(1, i).value) Then
                            hold = DataSheet.Cells(1, i).value & "_Num"
                            hold2 = DataSheet.Cells(userRow, i).value
                            TestCopy.Cells(k, 13).value = Lookup(DataSheet.Cells(1, i).value & "_Num")(DataSheet.Cells(userRow, i).value)
                            TestCopy.Cells(k, 14).value = DataSheet.Cells(1, i).value
                        End If
                        TestCopy.Cells(k, 11).EntireRow.Interior.Color = xlNone
                        k = k + 1
                    End If
                End If
            Next j
        End If
        Dim athing As String
        athing = DataSheet.Range(numToAlpha(i) & 2).value
        Select Case DataSheet.Range(numToAlpha(i) & 2).value
            Case "DEMOGRAPHICS", "PETITION", "DRAI", "DETENTION", "DETENTION (VOP)", "DIVERSION", "4G", "4E", "6F", "6H", "3E", "Crossover", "WRAP", "JTC", "ADULT", "AGGREGATES"
                If TestCopy.Cells(k - 1, 9).Interior.Color = RGB(200, 200, 200) Then
                    TestCopy.Cells(k, 9).value = "..."
                    TestCopy.Cells(k, 10).value = "No Change"
                    TestCopy.Cells(k, 11).EntireRow.Interior.Color = xlNone
                    k = k + 1
                End If

                TestCopy.Cells(k, 9).value = numToAlpha(i) 'column letter
                TestCopy.Cells(k, 10).value = Worksheets("Entry").Cells(2, i).value 'header name
                TestCopy.Range("I" & k & ":L" & k).Interior.Color = RGB(200, 200, 200)
                k = k + 1
        End Select

    Next i


    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
    End With
End Sub

Sub Test_Copy()
    Sheets("TestOutput").Cells.Copy Destination:=Sheets("TestCopy").Range("A1")

    Dim btn As Shape
    Dim totalTable As Range

    Set totalTable = Sheets("TestCopy").Range("G2")

    For Each btn In ActiveSheet.Shapes
        If Not Intersect(btn.TopLeftCell, totalTable) Is Nothing Then btn.Delete
    Next btn


End Sub

Sub Full_Print()

    Worksheets("TestOutput").Range("A2:X2000").ClearContents
    Dim userRow As Long
    Dim i As Long
    Dim j As Long
    Dim dicName As String
    Dim value As Long
    Call RefreshNamedRanges
    Call Generate_Dictionaries

    'marker for what ROW of the test output we're writing to
    j = 2


    k = 1

    'Grab the row number that was written in G1 of the "TestOutput" sheet
    userRow = Worksheets("TestFull").Range("G1").value


    With Worksheets("Entry")

        'i represents column index, so we're starting at column "B" and moving right
        For i = 2 To 200

            'write the column letter in row 2, column A of the "TestOutput" row we're currently at
            Worksheets("TestFull").Cells(j, 1).value = numToAlpha(i) 'column letter

            'write the header name in row 2, column A of the "TestOutput" row we're currently at
            Worksheets("TestFull").Cells(j, 2).value = Worksheets("Entry").Cells(2, i).value 'header name


            'if there's a Lookup List value in ROW1 of the column we're talking about
            If isNotEmptyOrZero(Worksheets("Entry").Range(numToAlpha(i) & 1)) Then

                'make the name of the lookup dictionary by appending "_Num" to the value in ROW1
                dicName = Worksheets("Entry").Range(numToAlpha(i) & 1).value & "_Num"

                'pull the raw value (enum number) from the cell
                value = Worksheets("Entry").Cells(userRow, i).value

                'set color to easily identify these translation printouts
                Worksheets("TestFull").Range("A" & j & ":E" & j).Interior.Color = RGB(240, 240, 240)

                'print the translated value to column "D"
                Worksheets("TestFull").Cells(j, 4).value = Lookup(dicName)(value)

                'print the raw value from the datasheet to column "C"
                Worksheets("TestFull").Cells(j, 3).value = Worksheets("Entry").Cells(userRow, i).value
                Worksheets("TestFull").Cells(j, 3).NumberFormat = "General"
                'print the lookup dictionary name
                Worksheets("TestFull").Cells(j, 5).value = Worksheets("Entry").Cells(1, i).value

                'if there's not a lookup list value in ROW1 (meaning that this will be a direct value entry
            Else

                'print the value from the datasheet to column "C"
                Worksheets("TestFull").Cells(j, 3).value = Worksheets("Entry").Cells(userRow, i).value

                'make sure this row is not colored (resetting in case it was colored on last render
                Worksheets("TestFull").Cells(j, 3).EntireRow.Interior.Color = xlNone
            End If



            'look at the header name for this column
            Select Case Worksheets("Entry").Range(numToAlpha(i) & 2).value

                    'if it's the name of one of our big section headers
                Case "DEMOGRAPHICS", "PETITION", "DRAI", "INTAKE CONFERENCE", "DETENTION", "DETENTION (VOP)", "DIVERSION", "4G", "4E", "6F", "6H", "3E", "Crossover", "WRAP", "JTC", "ADULT", "AGGREGATES", "LISTINGS"

                    'if the prior row was also one of the big section headers
                    If Worksheets("TestFull").Cells(j - 1, 2).Interior.Color = RGB(200, 200, 200) Then

                        'print "... Empty Section" and make the row uncolored
                        Worksheets("TestFull").Cells(j, 1).value = "..."
                        Worksheets("TestFull").Cells(j, 2).value = "Empty Section"
                        Worksheets("TestFull").Cells(j, 3).EntireRow.Interior.Color = xlNone

                        'move on to next printrow
                        j = j + 1
                    End If

                    'print column letter to column "A"
                    Worksheets("TestFull").Cells(j, 1).value = numToAlpha(i) 'column letter

                    'print header name to column "B"
                    Worksheets("TestFull").Cells(j, 2).value = Worksheets("Entry").Cells(2, i).value 'header name

                    'print color the row darkly
                    Worksheets("TestFull").Range("A" & j & ":E" & j).Interior.Color = RGB(200, 200, 200)


            End Select

            'move on to next printrow
            j = j + 1


            'move on to next column on datasheet
        Next i
    End With
End Sub
