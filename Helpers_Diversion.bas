Attribute VB_Name = "Helpers_Diversion"
Public Function getContractTerm(row As Long, Num As Long) As Integer
    getContractTerm = Range(headerFind("Contract Term #" & Num, headerFind("DIVERSION")) & row)
    Dim i As Integer

    For i = 1 To 5
        If Range(headerFind("Which Term Was Updated #" & i, headerFind("DIVERSION")) & row) = Num Then
            getContractTerm = Range(headerFind("New Term", headerFind("Which Term Was Updated #" & i, headerFind("DIVERSION"))) & row)
        End If
    Next i
End Function

Public Function getContractTermProvider(row As Long, Num As Long) As Integer
    getContractTermProvider _
        = Range( _
            headerFind("Contract Term #" & Num & " Provider", _
            headerFind("DIVERSION")) & row)
    Dim i As Integer

    For i = 1 To 5
        If Range(headerFind("Which Term Was Updated #" & i, headerFind("DIVERSION")) & row) = Num Then
            getContractTermProvider _
                = Range( _
                    headerFind("New Term Provider", _
                    headerFind("Which Term Was Updated #" & i, _
                    headerFind("DIVERSION"))) & row)
        End If
    Next i
End Function

Public Function getContractTermDate(row As Long, Num As Long) As String
    getContractTermDate = Range(headerFind("Date of Contract", headerFind("DIVERSION")) & row)
    Dim i As Integer

    For i = 1 To 5
        If Range(headerFind("Which Term Was Updated #" & i, headerFind("DIVERSION")) & row) = Num Then
            getContractTermDate _
                = Range(headerFind("Date of Update", _
                    headerFind("Which Term Was Updated #" & i, _
                    headerFind("DIVERSION"))) & row)
        End If
    Next i
End Function

Public Function getOpenTermEdit(row As Long) As Integer
    getOpenTermEdit = 0
    Dim i As Integer

    For i = 1 To 5
        If isEmptyOrZero(Range(headerFind("Which Term Was Updated #" & i, headerFind("DIVERSION")) & row)) Then
            getOpenTermEdit = i
            Exit Function
        End If
    Next i

End Function
