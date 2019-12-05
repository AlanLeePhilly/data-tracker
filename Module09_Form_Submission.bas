Attribute VB_Name = "Module09_Form_Submission"
Public Sub CLEAR_DATA_DANGEROUS()
    Worksheets("Entry").Range("C3:SRZ500").ClearContents
End Sub

Sub formSubmitStart(Optional userRow As Long = -1)
    Worksheets("Entry").Activate
    If userRow >= 0 Then
        Call cacheRow(userRow)
    End If

    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With
End Sub


Sub formSubmitEnd()
    Call Save_Countdown
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
    End With
    Worksheets("User Entry").Activate
End Sub
