Attribute VB_Name = "Module05_HeaderFind"
Public Function headerFind(ByVal Query As String, Optional ByVal Start As String = "A") As String
    Dim Result As Variant
    If Start = "" Then
        Start = "A"
    End If

    Set Result = Worksheets("Entry").Range(Start & "2:XFD2").Find(Query, LookAt:=xlWhole)


    If Result Is Nothing Then
        err.Raise 9999, "headerFind", _
        "Error: Tried to find column header " & Chr(34) & Query & Chr(34) _
        & " starting at column " & Start & " but could not find it."
        Exit Function
    Else
        headerFind = numToAlpha(Result.Column)
        Exit Function
    End If

End Function

Public Function numToAlpha(ByVal ColNum As Long) As String
    numToAlpha = Split(Cells(, ColNum).Address, "$")(1)
End Function

Public Function alphaToNum(ColAlpha As String) As Long
    alphaToNum = Range(ColAlpha & 1).Column
End Function

Function hFind(ParamArray myArgs() As Variant) As String
    ' PASS THIS FUNCTION JUST A LIST OF STRING ARGUMENTS IN DESCENDING SPECIFICITY
    ' EXAMPLE: Call hFind("Start Date", "Pretrial", "4E")

    Dim Result As Variant
    Dim Start As String
    Dim length As Integer
    Dim i As Integer

    length = UBound(myArgs) - LBound(myArgs) + 1
    For i = UBound(myArgs) To LBound(myArgs) Step -1
        If length = 1 Then
            hFind = headerFind(myArgs(i))
            Debug.Print hFind
            Exit Function
        End If
        Start = headerFind(myArgs(i), Start)
    Next i
    hFind = Start
    Debug.Print Start

End Function

