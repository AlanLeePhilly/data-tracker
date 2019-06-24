Attribute VB_Name = "Module03_Generate_Lookups"
Sub RefreshNamedRanges()
    With Worksheets("MasterList")

        Dim lastRow As Long
        Dim Title As String
        Dim rng As Range
        Dim i As Long

        For i = 1 To 10000
            If Not IsEmpty(.Cells(1, i).value) Then
                Title = .Cells(1, i).value
                lastRow = .Cells(.Rows.count, i).End(xlUp).row

                If RangeExists(Title) Then
                    ThisWorkbook.Names(Title).Delete
                End If

                If IsEmpty(.Cells(1, i + 1).value) Then
                    Set rng = .Range(numToAlpha(i) + "2:" + numToAlpha(i + 1) + CStr(lastRow))
                Else
                    Set rng = .Range(numToAlpha(i) + "2:" + numToAlpha(i) + CStr(lastRow))
                End If

                ThisWorkbook.Names.Add name:=Title, RefersTo:=rng
            End If
        Next i
    End With

    With Worksheets("CrimeCodes")
        Title = "Crime_Code"
        lastRow = .Cells(.Rows.count, 1).End(xlUp).row

        If RangeExists(Title) Then
            ThisWorkbook.Names(Title).Delete
        End If

        Set rng = .Range("A2:" + "E" + CStr(lastRow))

        ThisWorkbook.Names.Add name:=Title, RefersTo:=rng

    End With
End Sub

Public Function RangeExists(r As String) As Boolean
    Dim Test As Range
    On Error Resume Next
    Set Test = ThisWorkbook.Names(r)
    RangeExists = err.Number = 0
End Function

Sub SetRangesToDictionaries()
    Call Generate_Dictionaries

End Sub

Public Function Generate_Dictionaries()

    With Worksheets("MasterList")
        Dim Title As String
        Dim rng As Range
        Dim i As Integer
        Dim k As Variant
        Dim j As Variant

        Set Lookup = New Scripting.Dictionary
        For i = 1 To 10000
            If Not IsEmpty(.Cells(1, i).value) Then
                Title = .Cells(1, i).value
                Set Lookup(Title + "_Name") = New Scripting.Dictionary
                Set Lookup(Title + "_Num") = New Scripting.Dictionary
                Call Make_Name_Dictionary(.Range(Title), Lookup(Title + "_Name"))
                Call Make_Num_Dictionary(.Range(Title), Lookup(Title + "_Num"))
            End If
        Next i

        Set Lookup("Crime_Code_Name") = New Scripting.Dictionary
        Set Lookup("Crime_Code_Num") = New Scripting.Dictionary
        Call Make_Name_Dictionary(Range("Crime_Code"), Lookup("Crime_Code_Name"))
        Call Make_Num_Dictionary(Range("Crime_Code"), Lookup("Crime_Code_Num"))

        'For Each k In Lookup.Keys
        'Debug.Print "Table: " & k
        'If k = "JTC_Outcome_Name" Then

        'For Each j In Lookup(k).Keys
        'Debug.Print "Key: " & j & " Val: " & Lookup(k)(j)
        'Next j
        'End If
        'Next k
    End With
End Function



Public Function Make_Name_Dictionary(ByRef Named As Variant, ByRef Dict As Scripting.Dictionary)
    For Each row In Named.Rows
        Dict(Named.Cells(row.row - 1, 1).value) = Named.Cells(row.row - 1, 2).value
    Next row
End Function

Public Function Make_Num_Dictionary(ByRef Named As Variant, ByRef Dict As Scripting.Dictionary)
    For Each row In Named.Rows
        Dict(Named.Cells(row.row - 1, 2).value) = Named.Cells(row.row - 1, 1).value
    Next row
End Function

Public Function CodeConcat(ByVal Title As String, ByVal section As String, ByVal subsection As String) As String
    If Not subsection = "" Then
        CodeConcat = Title & " - " & section & " - " & subsection
    Else
        CodeConcat = Title & " - " & section
    End If
End Function
