Attribute VB_Name = "Helpers_Mini"


Sub append(ByRef rng As Range, val As String, Optional dateMark As String = "1/1/1900")
    Dim payload As String

    If Not dateMark = "1/1/1900" Then
        payload = dateMark & " - " & val & ";"
    Else
        payload = val & ";"
    End If

    If IsEmpty(rng) Then
        rng.value = payload
    Else
        rng.value = rng & vbNewLine & payload
    End If
End Sub

Sub prepend(ByRef rng As Range, val As String, Optional dateMark As String = "1/1/1900")
    Dim payload As String

    If Not dateMark = "1/1/1900" Then
        payload = dateMark & " - " & val & ";"
    Else
        payload = val & ";"
    End If

    If IsEmpty(rng) Then
        rng.value = payload
    Else
        rng.value = payload & vbNewLine & rng
    End If
End Sub

Sub trimmer()
    Dim count As Long
    With Sheets("Entry")
        For count = 1 To 15000
            If Not Trim(.Range(numToAlpha(count) & "2").value) = .Range(numToAlpha(count) & "2").value Then
                Debug.Print numToAlpha(count)
                Debug.Print .Range(numToAlpha(count) & "2").value
                Debug.Print Trim(.Range(numToAlpha(count) & "2").value)

                '.Range(numToAlpha(count) & "2").value = Trim(.Range(numToAlpha(count) & "2").value)
            End If
        Next count
    End With
End Sub

Sub trimmer2()
    Dim countX As Long
    Dim countY As Long

    With Sheets("MasterList")
        For countX = 1 To 200
            For countY = 1 To 200
                If Not Trim(.Range(numToAlpha(countX) & countY).value) = .Range(numToAlpha(countX) & countY).value _
                    And Not IsNumeric(.Range(numToAlpha(countX) & countY).value) Then

                    Debug.Print numToAlpha(countX) + CStr(countY)
                    Debug.Print .Range(numToAlpha(countX) & countY).value
                    Debug.Print Trim(.Range(numToAlpha(countX) & countY).value)

                    '.Range(numToAlpha(countX) & countY).value = Trim(.Range(numToAlpha(countX) & countY).value)
                End If
            Next countY
        Next countX
    End With
End Sub

Function isNotEmptyOrZero(ByRef rng As Range) As Boolean
    If Not IsEmpty(rng) And rng.value > 0 Then
        isNotEmptyOrZero = True
    Else
        isNotEmptyOrZero = False
    End If
End Function

Function isEmptyOrZero(ByRef rng As Range) As Boolean
    If IsEmpty(rng) Or rng.value = 0 Then
        isEmptyOrZero = True
    Else
        isEmptyOrZero = False
    End If
End Function
Function isResidential(ByVal supType As String) As Boolean
    If Range("Community_Based_Supervision").Find(supType, , Excel.xlValues) Is Nothing Then
        isResidential = True
    Else
        isResidential = False
    End If
End Function


Sub toggleSelect( _
    ByRef btn As Control, _
    Optional ByRef targLabel As Control, _
    Optional ByVal selVal As String, _
    Optional ByVal unselVal As String = "" _
)
    'If btn.BackColor = selectedColor Then
    'btn.BackColor = unselectedColor
    'If Not targLabel Is Nothing Then
    'targLabel.Caption = unselVal
    'End If
    'Else
    'If btn.BackColor = unselectedColor Then
    btn.BackColor = selectedColor
    If Not targLabel Is Nothing Then
        targLabel.Caption = selVal
    End If
    'End If
    'End If
End Sub

Function findFirstValue(ds As Worksheet, rowNum As Long, header As String, sectionHeads() As String, startColName As String, endColName As String) As String
    'Used to go through several columns in a section, and returns first one that does not have end value
    Dim arrLength As Integer
    arrLength = UBound(sectionHeads) - LBound(sectionHeads) + 1

    Dim i As Integer
    Dim startFound As String
    Dim endFound As String
    For i = 1 To arrLength
        startFound = ds.Cells(rowNum, hFind(startColName, sectionHeads(i), header)).value
        endFound = ds.Cells(rowNum, hFind(endColName, sectionHeads(i), header)).value

        'If start found, but no end found, return sectionHeads of this index
        If Not StrComp(startFound, "") = 0 Then
            If StrComp(endFound, "") = 0 Then
                findFirstValue = sectionHeads(i)
                Exit Function
            End If
        End If
    Next i
End Function

Function findFirstValue2(ds As Worksheet, rowNum As Long, header As String, sectionHeads() As String, searchVariable As String, startColName As String, endColName As String) As String
    'Used to go through several columns in a section, and returns first one that does not have end value
    Dim arrLength As Integer
    arrLength = UBound(sectionHeads) - LBound(sectionHeads) + 1

    Dim i As Integer
    Dim startFound As String
    Dim endFound As String
    For i = 1 To arrLength
        searchVariableFound = ds.Cells(rowNum, hFind(searchVariableName, sectionHeads(i), header)).value
        startFound = ds.Cells(rowNum, hFind(startColName, sectionHeads(i), header)).value
        endFound = ds.Cells(rowNum, hFind(endColName, sectionHeads(i), header)).value

        'If search variable found, start found, but no end found, return searchvariable of this index
        If Not StrComp(searchVariableFound, "") = 0 Then
            If Not StrComp(startFound, "") = 0 Then
                If StrComp(endFound, "") = 0 Then
                    findFirstValue2 = sectionHeads(i)
                    Exit Function
                End If
            End If
        End If
    Next i
End Function


Function findAllValues(ds As Worksheet, rowNum As Long, header As String, sectionHead As String, startColName As String, endColName As String) As String()
    'Used to go through several columns in a section, and returns first one that does not have end value
    Dim returnArr() As String
    Dim returnValues As Integer
    returnValues = 0

    Dim i As Integer
    Dim startFound As String
    Dim endFound As String
    For i = 1 To 6
        startFound = ds.Cells(rowNum, hFind(startColName, sectionHead & " #" & i, header)).value
        endFound = ds.Cells(rowNum, hFind(endColName, sectionHead & " #" & i, header)).value

        'If start found, but no end found, return sectionHeads of this index
        If Not StrComp(startFound, "") = 0 Then
            If StrComp(endFound, "") = 0 Then
                ReDim Preserve returnArr(returnValues)
                returnArr(returnValues) = sectionHead & " #" & i
                returnValues = returnValues + 1
            End If
        End If
    Next i

    findAllValues = returnArr
End Function
Function findAllValues2(ds As Worksheet, rowNum As Long, header As String, sectionHead As String, searchVariable As String, startColName As String, endColName As String) As String()
    'Used to go through several columns in a section, and returns first one that does not have end value
    Dim returnArr() As String
    Dim returnValues As Integer
    returnValues = 0

    Dim i As Integer
    Dim startFound As String
    Dim endFound As String
    For i = 1 To 7
        startFound = ds.Cells(rowNum, hFind(startColName, sectionHead & " #" & i, header)).value
        endFound = ds.Cells(rowNum, hFind(endColName, sectionHead & " #" & i, header)).value

        'If start found, but no end found, return sectionHeads of this index
        If Not StrComp(startFound, "") = 0 Then
            If StrComp(endFound, "") = 0 Then
                ReDim Preserve returnArr(returnValues)
                returnArr(returnValues) = sectionHead & " #" & i
                returnValues = returnValues + 1
            End If
        End If
    Next i

    findAllValues2 = returnArr
End Function



Function ageAtTime(eventDate As String, rowNum As Long) As Double
    Dim DOB As String
    DOB = Worksheets("Entry").Range(headerFind("DOB") & rowNum).value
    ageAtTime = Round((DateDiff("d", DOB, eventDate) / 365), 2)
End Function

Function calcLOS(ByVal event1 As String, ByVal event2 As String) As Long
    calcLOS = DateDiff("d", event1, event2)
End Function

Function timeDiff(time1 As Double, time2 As Double) As Double
    If time1 < time2 Then
        timeDiff = time2 - time1
    Else
        timeDiff = 1 + time2 - time1
    End If
End Function

Sub flagNo(ByRef rng As Range)
    'WILL NOT OVERWRITE A "YES" or "NO"
    If isEmptyOrZero(rng) Then
        rng.value = Lookup("Generic_YN_Name")("No")
    End If
End Sub

Sub flagYes(ByRef rng As Range)
    rng.value = Lookup("Generic_YN_Name")("Yes")
End Sub

Sub addNotes(Courtroom As String, dateOf As String, userRow As Long, Notes As String, Optional legalStatus As String = "")
    Dim bucketHead As String
    For i = 1 To 100
        If IsEmpty(Range(hFind("Court Date #" & i, "LISTINGS") & userRow)) Then
            bucketHead = hFind("Court Date #" & i, "LISTINGS")

            Range(bucketHead & userRow).value = dateOf
            Range(headerFind("Courtroom", bucketHead) & userRow).value = Lookup("Courtroom_Name")(Courtroom)
            Range(headerFind("Legal Status", bucketHead) & userRow).value = Lookup("Legal_Status_Name")(legalStatus)
            Range(headerFind("DA", bucketHead) & userRow).value = Lookup("DA_Last_Name_Name")(DA)
            Range(headerFind("Notes", bucketHead) & userRow).value = Notes

            i = 100
        End If
    Next i
End Sub

Sub addPetitionsToBox(ByRef MyBox As Object)
    Dim i As Integer
    Dim bucketHead As String
    MyBox.Clear
    Worksheets("Entry").Activate

    For i = 1 To 5
        If Range(hFind("Petition Filed?", "Petition #" & i, "PETITION") & updateRow).value = Lookup("Generic_YNOU_Name")("Yes") Then
            bucketHead = hFind("Petition #" & i, "PETITION")
            With MyBox
                .ColumnCount = 6
                .ColumnWidths = "50;50;30;50;65;50"
                ' 0 Petition Number
                ' 1 Date Filed
                ' 2 Charge Grade
                ' 3 Charge Group
                ' 4 Charge Code
                ' 5 Charge Name

                .AddItem Range(bucketHead & updateRow).value
                .List(MyBox.ListCount - 1, 0) = Range(bucketHead & updateRow).value
                .List(MyBox.ListCount - 1, 1) = Range(headerFind("Date Filed", bucketHead) & updateRow).value
                .List(MyBox.ListCount - 1, 2) = Lookup("Charge_Grade_Specific_Num")(Range(headerFind("Charge Grade (specific) #1", bucketHead) & updateRow).value)
                .List(MyBox.ListCount - 1, 3) = Lookup("Charge_Num")(Range(headerFind("Charge Category #1", bucketHead) & updateRow).value)
                .List(MyBox.ListCount - 1, 4) = Range(headerFind("Lead Charge Code", bucketHead) & updateRow).value
                .List(MyBox.ListCount - 1, 5) = Range(headerFind("Lead Charge Name", bucketHead) & updateRow).value

            End With
        End If
    Next i

End Sub

Sub UnloadAll()
    Dim objLoop As Object
    Dim nameOf As String


    Dim i As Long
    For i = VBA.UserForms.count - 1 To 0 Step -1
        nameOf = VBA.UserForms(i).name
        Unload VBA.UserForms(i)
    Next i
End Sub

Public Function NatureFromDetailed(detailed As String) As String
    Select Case detailed
        Case _
            "Judgment of Acquittal", _
            "Petition Closed - Positive Comp. Terms", _
            "Petition Withdrawn", _
            "Petition Diverted and Withdrawn"

                NatureFromDetailed = "Positive"
                
        Case _
            "Rearested & Held (adult)", _
            "Bench Warrant", _
            "Transfer to New Del. Room - Negative", _
            "Aged Out", _
            "Certified Adult (original petition)", _
            "Death", _
            "Admin. D/C - Reasonable Efforts"
            
                NatureFromDetailed = "Negative"
                
        Case _
            "Transfer to Dependent", _
            "Acceptance to Room Not Granted", _
            "Transfer to Other County", _
            "Transfer to New Del. Room - Neutral", _
            "Not Fit to Stand Trial", _
            "Other"
            
                NatureFromDetailed = "Neutral"
                
        Case "N/A", "Unknown"
        
            NatureFromDetailed = detailed
            
    End Select
End Function

Public Function statusHasAgg(legalStatus As String) As Boolean
    Select Case legalStatus
        Case "Pretrial", "Consent Decree", "Probation", "Interim Probation", "Aftercare Probation"
            statusHasAgg = True
        Case Else
            statusHasAgg = False
    End Select
End Function
