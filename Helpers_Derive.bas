Attribute VB_Name = "Helpers_Derive"
Function calcTimeGroup(ByVal hrs As Integer, ByVal per As String) As Long
    If per = "AM" Then
        Select Case hrs
            Case 12, 1, 2
                calcTimeGroup = Lookup("Time_Block_Name")("12:00AM - 2:59AM")
            Case 3, 4, 5, 6
                calcTimeGroup = Lookup("Time_Block_Name")("3:00AM - 6:59AM")
            Case 7, 8, 9, 10, 11
                calcTimeGroup = Lookup("Time_Block_Name")("7:00AM - 2:59PM")
        End Select
    End If

    If per = "PM" Then
        Select Case hrs
            Case 12, 1, 2
                calcTimeGroup = Lookup("Time_Block_Name")("7:00AM - 2:59PM")
            Case 3, 4, 5, 6, 7
                calcTimeGroup = Lookup("Time_Block_Name")("3:00PM - 7:59PM")
            Case 8, 9, 10, 11
                calcTimeGroup = Lookup("Time_Block_Name")("8:00PM - 11:59PM")
        End Select
    End If

    If Not IsNumeric(calcTimeGroup) Then
        err.Raise 9998, "calcTimeGroup"
        Exit Function
    End If
End Function

Function calcChargeBroad(ByVal chargeSpecific As String) As Long
    Select Case chargeSpecific
        Case "F1", "F2", "F3", "F"
            calcChargeBroad = Lookup("Charge_Grade_Broad_Name")("Felony")
        Case "M1", "M2", "M3", "M"
            calcChargeBroad = Lookup("Charge_Grade_Broad_Name")("Misdemeanor")
        Case "Summary"
            calcChargeBroad = Lookup("Charge_Grade_Broad_Name")("Summary")
        Case "Other"
            calcChargeBroad = Lookup("Charge_Grade_Broad_Name")("Other")
        Case "Unknown"
            calcChargeBroad = Lookup("Charge_Grade_Broad_Name")("Unknown")
    End Select
    If Not IsNumeric(calcChargeBroad) Then
        err.Raise 9998, "calcChargeBroad"
        Exit Function
    End If
End Function

Function commonwealthCat(ByVal reason As String) As Long
    Select Case Lookup("Detailed_Reason_for_Commonwealth_Continuance_Name")(reason)
        Case 0
            commonwealthCat = 0
        Case 1, 2, 3
            commonwealthCat = 1
        Case 4, 5
            commonwealthCat = 2
        Case Else
            commonwealthCat = 99
    End Select
End Function


