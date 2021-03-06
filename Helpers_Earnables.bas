Attribute VB_Name = "Helpers_Earnables"

Sub fetchFiledRecord(ByVal userRow As Long)
    Dim i As Integer
    Dim sectionHead As String
    Dim restHead As String
    Dim costHead As String
    Dim commHead As String
    Dim bucketHead As String
    
    Worksheets("Entry").Activate
    Log_Payment.History.Clear
    
    'FETCH RESTITUTION
    restHead = hFind("Restitution", "AGGREGATES")
    For i = 1 To NUM_RESTITUTION_FILED_BUCKETS
        If isNotEmptyOrZero(Range(headerFind("Amount Filed #" & i, restHead) & userRow)) Then

            bucketHead = headerFind("Amount Filed #" & i, restHead)
            
            With Log_Payment.History
                .ColumnCount = 5
                .ColumnWidths = "50;50;50;50;50;"
                ' 0 Filing Type
                ' 1 Date
                ' 2 Amount Filed
                ' 3 Amount Paid
                ' 4 Remaining Balance
                .AddItem "Restitution"
                .List(.ListCount - 1, 0) = "Restitution"
                .List(.ListCount - 1, 1) = Range(headerFind("Date", bucketHead) & userRow).value
                .List(.ListCount - 1, 2) = Range(bucketHead & userRow).value
            End With
        End If
    Next i
    

    For i = 1 To NUM_RESTITUTION_PAID_BUCKETS
        If isNotEmptyOrZero(Range(headerFind("Amount Paid #" & i, restHead) & userRow)) Then
            bucketHead = headerFind("Amount Paid #" & i, restHead)
            
            With Log_Payment.History
                .ColumnCount = 5
                .ColumnWidths = "50;50;50;50;50;"
                ' 0 Filing Type
                ' 1 Date
                ' 2 Amount Filed
                ' 3 Amount Paid
                ' 4 Remaining Balance
                .AddItem "Restitution"
                .List(.ListCount - 1, 0) = "Restitution"
                .List(.ListCount - 1, 1) = Range(headerFind("Date", bucketHead) & userRow).value
                .List(.ListCount - 1, 3) = Range(bucketHead & userRow).value
            End With
        End If
    Next i
    Log_Payment.Remaining_Restitution.Caption = Range(headerFind("Total Amount Remaining", restHead) & userRow).value


    'FETCH COURT COSTS
    costHead = hFind("Court Costs", "AGGREGATES")
    
    For i = 1 To NUM_COURT_COST_FILED_BUCKETS
        If isNotEmptyOrZero(Range(headerFind("Amount Filed #" & i, costHead) & userRow)) Then

            bucketHead = headerFind("Amount Filed #" & i, costHead)
            
            With Log_Payment.History
                .ColumnCount = 5
                .ColumnWidths = "50;50;50;50;50;"
                ' 0 Filing Type
                ' 1 Date
                ' 2 Amount Filed
                ' 3 Amount Paid
                ' 4 Remaining Balance
                .AddItem "Court Cost"
                .List(.ListCount - 1, 0) = "Court Cost"
                .List(.ListCount - 1, 1) = Range(headerFind("Date", bucketHead) & userRow).value
                .List(.ListCount - 1, 2) = Range(bucketHead & userRow).value
            End With
        End If
    Next i
    

    For i = 1 To NUM_COURT_COST_PAID_BUCKETS
        If isNotEmptyOrZero(Range(headerFind("Amount Paid #" & i, costHead) & userRow)) Then
            bucketHead = headerFind("Amount Paid #" & i, costHead)
            
            With Log_Payment.History
                .ColumnCount = 5
                .ColumnWidths = "50;50;50;50;50;"
                ' 0 Filing Type
                ' 1 Date
                ' 2 Amount Filed
                ' 3 Amount Paid
                ' 4 Remaining Balance
                .AddItem "Court Cost"
                .List(.ListCount - 1, 0) = "Court Cost"
                .List(.ListCount - 1, 1) = Range(headerFind("Date", bucketHead) & userRow).value
                .List(.ListCount - 1, 3) = Range(bucketHead & userRow).value
            End With
        End If
    Next i
    Log_Payment.Remaining_Court_Cost.Caption = Range(headerFind("Total Amount Remaining", costHead) & userRow).value


    'FETCH COMM SERV
    commHead = hFind("Comm. Service", "AGGREGATES")
    
    For i = 1 To NUM_COMM_SERVICE_FILED_BUCKETS
        If isNotEmptyOrZero(Range(headerFind("Amount Filed #" & i, commHead) & userRow)) Then

            bucketHead = headerFind("Amount Filed #" & i, commHead)
            
            With Log_Payment.History
                .ColumnCount = 5
                .ColumnWidths = "50;50;50;50;50;"
                ' 0 Filing Type
                ' 1 Date
                ' 2 Amount Filed
                ' 3 Amount Paid
                ' 4 Remaining Balance
                .AddItem "Comm. Serv"
                .List(.ListCount - 1, 0) = "Comm. Serv"
                .List(.ListCount - 1, 1) = Range(headerFind("Date", bucketHead) & userRow).value
                .List(.ListCount - 1, 2) = Range(bucketHead & userRow).value
            End With
        End If
    Next i
    
    For i = 1 To NUM_COMM_SERVICE_EARNED_BUCKETS
        If isNotEmptyOrZero(Range(headerFind("Amount Earned #" & i, commHead) & userRow)) Then
            bucketHead = headerFind("Amount Earned #" & i, commHead)
            
            With Log_Payment.History
                .ColumnCount = 5
                .ColumnWidths = "50;50;50;50;50;"
                ' 0 Filing Type
                ' 1 Date
                ' 2 Amount Filed
                ' 3 Amount Paid
                ' 4 Remaining Balance
                .AddItem "Comm. Serv"
                .List(.ListCount - 1, 0) = "Comm. Serv"
                .List(.ListCount - 1, 1) = Range(headerFind("Date", bucketHead) & userRow).value
                .List(.ListCount - 1, 3) = Range(bucketHead & userRow).value
            End With
        End If
    Next i
    Log_Payment.Remaining_Comm_Serv.Caption = Range(headerFind("Total Amount Remaining", commHead) & userRow).value

End Sub

Sub startRestitution( _
    ByVal Amount As String, _
    ByVal Courtroom As String, _
    ByVal DA As String, _
    ByVal DateOf As String, _
    ByVal userRow As Long _
)
    
    Dim sectionHead As String
    sectionHead = hFind("Restitution", "AGGREGATES")
    
    If isNotEmptyOrZero(Range(headerFind("Amount Filed #1", sectionHead) & userRow)) Then
        MsgBox "Warning: Restitution is already started for this client. Did not update to avoid overwrtiting"
        Exit Sub
    End If
    
    Range(headerFind("Did Youth Have Restitution?", sectionHead) & userRow).value = 1 'Yes
    

    Range(headerFind("Amount Filed #1", sectionHead) & userRow).value = Amount
    Range(headerFind("Date", sectionHead) & userRow).value = DateOf
    Range(headerFind("Courtroom", sectionHead) & userRow).value = Lookup("Courtroom_Name")(Courtroom)
    Range(headerFind("DA", sectionHead) & userRow).value = Lookup("DA_Last_Name_Name")(DA)
    Range(headerFind("Restitution Status", sectionHead) & userRow).value = 3 ' "Active and Unpaid"
    
    Call autoCalcRestitution(userRow)
    
End Sub

Sub updateRestitution( _
    ByVal Courtroom As String, _
    ByVal DA As String, _
    ByVal userRow As Long, _
    Optional ByVal DateOf As String = "", _
    Optional ByVal amountFiled As String = "", _
    Optional ByVal amountPaid As String = "" _
)
    Dim i As Integer
    Dim sectionHead As String
    Dim bucketHead As String
    sectionHead = hFind("Restitution", "AGGREGATES")
    
    Range(headerFind("Did Youth Have Restitution?", sectionHead) & userRow).value = 1 'Yes
    
    If Not amountFiled = "" Then
        For i = 1 To NUM_RESTITUTION_FILED_BUCKETS
            If isEmptyOrZero(Range(headerFind("Amount Filed #" & i, sectionHead) & userRow)) Then
                bucketHead = headerFind("Amount Filed #" & i, sectionHead)
                
                Range(bucketHead & userRow).value = amountFiled
                Range(headerFind("Date", bucketHead) & userRow).value = DateOf
                Range(headerFind("Courtroom", bucketHead) & userRow).value = Lookup("Courtroom_Name")(Courtroom)
                Range(headerFind("DA", bucketHead) & userRow).value = Lookup("DA_Last_Name_Name")(DA)
                
                Exit For
                ' TODO Warning for exceeding buckets
            End If
        Next i
    End If

    
    If Not amountPaid = "" Then
        For i = 1 To NUM_RESTITUTION_PAID_BUCKETS
            If isEmptyOrZero(Range(headerFind("Amount Paid #" & i, sectionHead) & userRow)) Then
                bucketHead = headerFind("Amount Paid #" & i, sectionHead)
                
                Range(bucketHead & userRow).value = amountPaid
                Range(headerFind("Date", bucketHead) & userRow).value = DateOf
                Range(headerFind("Courtroom", bucketHead) & userRow).value = Lookup("Courtroom_Name")(Courtroom)
                Range(headerFind("DA", bucketHead) & userRow).value = Lookup("DA_Last_Name_Name")(DA)
                
                Exit For
                ' TODO Warning for exceeding buckets
            End If
        Next i
    End If
    
    Call autoCalcRestitution(userRow)
    
End Sub

Sub autoCalcRestitution(ByVal userRow As Long)
    Dim i As Integer
    Dim sectionHead As String
    Dim bucketHead As String
    Dim dateOfLastPayment As String
    Dim dateOfFirstFiling As String
    Dim dateOfLastFiling As String
    sectionHead = hFind("Restitution", "AGGREGATES")
    
    
    'Calc Total Amount Filed
    Dim totalAmountFiled As Double
    totalAmountFiled = 0
    
    For i = 1 To NUM_RESTITUTION_FILED_BUCKETS
        If isNotEmptyOrZero(Range(headerFind("Amount Filed #" & i, sectionHead) & userRow)) Then
            bucketHead = headerFind("Amount Filed #" & i, sectionHead)
            totalAmountFiled = totalAmountFiled + CDbl(Range(bucketHead & userRow).value)
            dateOfLastFiling = Range(headerFind("Date", bucketHead) & userRow).value
            If i = 1 Then
                dateOfFirstFiling = Range(headerFind("Date", bucketHead) & userRow).value
                Range(headerFind("LOS Arrest to First Payment", sectionHead) & userRow).value _
                    = calcLOS(Range(hFind("Arrest Date") & userRow).value, dateOfFirstFiling)
            End If
        End If
    Next i
    
    Range(headerFind("Total Amount Filed", sectionHead) & userRow).value = totalAmountFiled
    
    
    
    'Calc Total Amount Remaining
    Dim totalAmountPaid As Double
    totalAmountPaid = 0
    
    For i = 1 To NUM_RESTITUTION_PAID_BUCKETS
        If isNotEmptyOrZero(Range(headerFind("Amount Paid #" & i, sectionHead) & userRow)) Then
            bucketHead = headerFind("Amount Paid #" & i, sectionHead)
            totalAmountPaid = totalAmountPaid + CDbl(Range(bucketHead & userRow).value)
            dateOfLastPayment = Range(headerFind("Date", bucketHead) & userRow).value
            
            If i = 1 Then
                Range(headerFind("LOS Arrest to First Payment", bucketHead) & userRow).value _
                = calcLOS( _
                    Range(hFind("Arrest Date") & userRow).value, _
                    Range(headerFind("Date", bucketHead) & userRow).value)
            End If
        End If
    Next i
    
    Range(headerFind("Total Amount Paid", sectionHead) & userRow).value = totalAmountPaid
    
    Range(headerFind("Total Amount Remaining", sectionHead) & userRow).value = totalAmountFiled - totalAmountPaid
    
    Range(headerFind("LOS to File", sectionHead) & userRow).value _
            = calcLOS(Range(headerFind("Arrest Date") & userRow).value, dateOfLastFiling)
    
    'Calc Restitution Status
    Dim restitutionStatus As String
    
    If totalAmountPaid >= totalAmountFiled Then
        restitutionStatus = 1 ' Paid in Full
        Range(headerFind("Date Paid in Full", sectionHead) & userRow).value = dateOfLastPayment
        Range(headerFind("LOS to Pay in Full", sectionHead) & userRow).value = calcLOS(dateOfFirstFiling, dateOfLastPayment)
        Range(headerFind("LOS to Pay in Full (from arrest)", sectionHead) & userRow).value = calcLOS(Range(headerFind("Arrest Date") & userRow).value, dateOfLastPayment)
    Else
        restitutionStatus = 3 ' Active and Unpaid
        Range(headerFind("Date Paid in Full", sectionHead) & userRow).value = ""
        Range(headerFind("LOS to Pay in Full", sectionHead) & userRow).value = ""
        Range(headerFind("LOS to Pay in Full (from arrest)", sectionHead) & userRow).value = ""
    
    End If
        
            
     Range(headerFind("Restitution Status", sectionHead) & userRow).value = restitutionStatus
     
     Call autoCalcCostsAndRest(userRow)

End Sub

Sub startCourtCosts( _
    ByVal Amount As String, _
    ByVal Courtroom As String, _
    ByVal DA As String, _
    ByVal DateOf As String, _
    ByVal userRow As Long _
)
    
    Dim sectionHead As String
    sectionHead = hFind("Court Costs", "AGGREGATES")
    
    If isNotEmptyOrZero(Range(headerFind("Amount Filed #1", sectionHead) & userRow)) Then
        MsgBox "Warning: Court Cost is already started for this client. Did not update to avoid overwrtiting"
        Exit Sub
    End If
    
    Range(headerFind("Did Youth Have Court Costs?", sectionHead) & userRow).value = 1 'Yes
    

    Range(headerFind("Amount Filed #1", sectionHead) & userRow).value = Amount
    Range(headerFind("Date", sectionHead) & userRow).value = DateOf
    Range(headerFind("Courtroom", sectionHead) & userRow).value = Lookup("Courtroom_Name")(Courtroom)
    Range(headerFind("DA", sectionHead) & userRow).value = Lookup("DA_Last_Name_Name")(DA)
    Range(headerFind("Court Cost Status", sectionHead) & userRow).value = 3 ' "Active and Unpaid"
    
    Call autoCalcCourtCost(userRow)
    
End Sub

Sub updateCourtCost( _
    ByVal Courtroom As String, _
    ByVal DA As String, _
    ByVal userRow As Long, _
    Optional ByVal DateOf As String = "", _
    Optional ByVal amountFiled As String = "", _
    Optional ByVal amountPaid As String = "" _
)
    Dim i As Integer
    Dim sectionHead As String
    Dim bucketHead As String
    sectionHead = hFind("Court Costs", "AGGREGATES")
    
    Range(headerFind("Did Youth Have Court Costs?", sectionHead) & userRow).value = 1 'Yes
    
    If Not amountFiled = "" Then
        For i = 1 To NUM_COURT_COST_FILED_BUCKETS
            If isEmptyOrZero(Range(headerFind("Amount Filed #" & i, sectionHead) & userRow)) Then
                bucketHead = headerFind("Amount Filed #" & i, sectionHead)
                
                Range(bucketHead & userRow).value = amountFiled
                Range(headerFind("Date", bucketHead) & userRow).value = DateOf
                Range(headerFind("Courtroom", bucketHead) & userRow).value = Lookup("Courtroom_Name")(Courtroom)
                Range(headerFind("DA", bucketHead) & userRow).value = Lookup("DA_Last_Name_Name")(DA)
                
                Exit For
                ' TODO Warning for exceeding buckets
            End If
        Next i
    End If

    
    If Not amountPaid = "" Then
        For i = 1 To NUM_COURT_COST_PAID_BUCKETS
            If isEmptyOrZero(Range(headerFind("Amount Paid #" & i, sectionHead) & userRow)) Then
                bucketHead = headerFind("Amount Paid #" & i, sectionHead)
                
                Range(bucketHead & userRow).value = amountPaid
                Range(headerFind("Date", bucketHead) & userRow).value = DateOf
                Range(headerFind("Courtroom", bucketHead) & userRow).value = Lookup("Courtroom_Name")(Courtroom)
                Range(headerFind("DA", bucketHead) & userRow).value = Lookup("DA_Last_Name_Name")(DA)
                
                Exit For
                ' TODO Warning for exceeding buckets
            End If
        Next i
    End If
    
    Call autoCalcCourtCost(userRow)
    
End Sub

Sub autoCalcCourtCost(ByVal userRow As Long)
    Dim i As Integer
    Dim sectionHead As String
    Dim bucketHead As String
    Dim dateOfLastPayment As String
    Dim dateOfFirstFiling As String
    Dim dateOfLastFiling As String
    sectionHead = hFind("Court Costs", "AGGREGATES")
    
    
    'Calc Total Amount Filed
    Dim totalAmountFiled As Double
    totalAmountFiled = 0
    
    For i = 1 To NUM_COURT_COST_FILED_BUCKETS
        If isNotEmptyOrZero(Range(headerFind("Amount Filed #" & i, sectionHead) & userRow)) Then
            bucketHead = headerFind("Amount Filed #" & i, sectionHead)
            totalAmountFiled = totalAmountFiled + CDbl(Range(bucketHead & userRow).value)
            dateOfLastFiling = Range(headerFind("Date", bucketHead) & userRow).value
            If i = 1 Then
                dateOfFirstFiling = Range(headerFind("Date", bucketHead) & userRow).value
                Range(headerFind("LOS Arrest to First Payment", sectionHead) & userRow).value _
                    = calcLOS(Range(hFind("Arrest Date") & userRow).value, dateOfFirstFiling)
            End If
        End If
    Next i
    
    Range(headerFind("Total Amount Filed", sectionHead) & userRow).value = totalAmountFiled
    
    
    'Calc Total Amount Remaining
    Dim totalAmountPaid As Double
    totalAmountPaid = 0
    
    For i = 1 To NUM_COURT_COST_PAID_BUCKETS
        If isNotEmptyOrZero(Range(headerFind("Amount Paid #" & i, sectionHead) & userRow)) Then
            bucketHead = headerFind("Amount Paid #" & i, sectionHead)
            totalAmountPaid = totalAmountPaid + CDbl(Range(bucketHead & userRow).value)
            dateOfLastPayment = Range(headerFind("Date", bucketHead) & userRow).value
            
            If i = 1 Then
                Range(headerFind("LOS Arrest to First Payment", bucketHead) & userRow).value _
                = calcLOS( _
                    Range(hFind("Arrest Date") & userRow).value, _
                    Range(headerFind("Date", bucketHead) & userRow).value)
            End If
        End If
    Next i
    
    Range(headerFind("Total Amount Paid", sectionHead) & userRow).value = totalAmountPaid
    
    Range(headerFind("Total Amount Remaining", sectionHead) & userRow).value = totalAmountFiled - totalAmountPaid
    
    Range(headerFind("LOS to File", sectionHead) & userRow).value _
            = calcLOS(Range(headerFind("Arrest Date") & userRow).value, dateOfLastFiling)
    'Calc Court Cost Status
    
    If totalAmountPaid >= totalAmountFiled Then
        Range(headerFind("Court Cost Status", sectionHead) & userRow).value = 1 ' Paid in Full
        Range(headerFind("Date Paid in Full", sectionHead) & userRow).value = dateOfLastPayment
        Range(headerFind("LOS to Pay in Full", sectionHead) & userRow).value = calcLOS(dateOfFirstFiling, dateOfLastPayment)
        Range(headerFind("LOS to Pay in Full (from arrest)", sectionHead) & userRow).value = calcLOS(Range(headerFind("Arrest Date") & userRow).value, dateOfLastPayment)
    Else
        Range(headerFind("Court Cost Status", sectionHead) & userRow).value = 3 ' Active and Unpaid
        Range(headerFind("Date Paid in Full", sectionHead) & userRow).value = ""
        Range(headerFind("LOS to Pay in Full", sectionHead) & userRow).value = ""
        Range(headerFind("LOS to Pay in Full (from arrest)", sectionHead) & userRow).value = ""
    End If
     
     Call autoCalcCostsAndRest(userRow)
End Sub

Sub autoCalcCostsAndRest(ByVal userRow As Long)
    Dim costHead As String
    Dim restHead As String
    Dim aggHead As String
    Dim aggOwed As Double
    Dim aggPaid As Double
    
    
    costHead = hFind("Court Costs", "AGGREGATES")
    restHead = hFind("Restitution", "AGGREGATES")
    aggHead = hFind("Costs & Restitution", "AGGREGATES")
    
    aggOwed = CDbl(Range(headerFind("Total Amount Filed", costHead) & userRow).value) _
            + CDbl(Range(headerFind("Total Amount Filed", restHead) & userRow).value)
            
    aggPaid = CDbl(Range(headerFind("Total Amount Paid", costHead) & userRow).value) _
            + CDbl(Range(headerFind("Total Amount Paid", restHead) & userRow).value)
    
    Range(headerFind("Aggregate Owed", aggHead) & userRow).value = aggOwed
    Range(headerFind("Aggregate Paid", aggHead) & userRow).value = aggPaid
    Range(headerFind("Aggregate Remaining", aggHead) & userRow).value = aggOwed - aggPaid
    
    If Range(headerFind("LOS to File", costHead) & userRow).value _
        > Range(headerFind("LOS to File", restHead) & userRow).value Then
        
        Range(headerFind("LOS to File", aggHead) & userRow).value _
        = Range(headerFind("LOS to File", costHead) & userRow).value
    Else
        Range(headerFind("LOS to File", aggHead) & userRow).value _
        = Range(headerFind("LOS to File", restHead) & userRow).value
    End If
        
    
    If aggPaid >= aggOwed Then
        Dim dateOfFirstRest As String
        Dim dateOfFirstCost As String
        Dim dateOfFirstFiling As String
        Dim dateOfLastRest As String
        Dim dateOfLastCost As String
        Dim dateOfLastPayment As String
        Dim dateOfLastFiling As String
        
        dateOfFirstRest = CDate(Range(headerFind("Date", restHead) & userRow).value)
        dateOfFirstCost = CDate(Range(headerFind("Date", costHead) & userRow).value)
        
        Select Case True
            Case dateOfFirstRest = "12:00:00 AM"
                dateOfFirstFiling = dateOfFirstCost
            Case dateOfFirstCost = "12:00:00 AM"
                dateOfFirstFiling = dateOfFirstRest
            Case dateOfFirstRest > dateOfFirstCost
                dateOfFirstFiling = dateOfFirstCost
            Case dateOfFirstRest < dateOfFirstCost
                dateOfFirstFiling = dateOfFirstRest
        End Select

        dateOfLastRest = CDate(Range(headerFind("Date Paid in Full", restHead) & userRow).value)
        dateOfLastCost = CDate(Range(headerFind("Date Paid in Full", costHead) & userRow).value)
        
        Select Case True
            Case dateOfLastRest = "12:00:00 AM"
                dateOfLastPayment = dateOfLastCost
            Case dateOfLastCost = "12:00:00 AM"
                dateOfLastPayment = dateOfLastRest
            Case dateOfLastRest > dateOfLastCost
                dateOfLastPayment = dateOfLastRest
            Case dateOfLastRest < dateOfLastCost
                dateOfLastPayment = dateOfLastCost
        End Select

 
        Range(headerFind("Total Cost Status", aggHead) & userRow).value = 1 ' Paid in Full
        Range(headerFind("LOS to Pay in Full", aggHead) & userRow).value _
            = calcLOS(dateOfFirstFiling, dateOfLastPayment)
        Range(headerFind("LOS to Pay in Full (from arrest)", aggHead) & userRow).value _
            = calcLOS(Range(headerFind("Arrest Date") & userRow).value, dateOfLastPayment)
    Else
        Range(headerFind("Total Cost Status", aggHead) & userRow).value = 3 ' Active and Unpaid
        Range(headerFind("LOS to Pay in Full", aggHead) & userRow).value = ""
        Range(headerFind("LOS to Pay in Full (from arrest)", aggHead) & userRow).value = ""
    End If
End Sub


Sub startCommService( _
    ByVal Amount As String, _
    ByVal Courtroom As String, _
    ByVal DA As String, _
    ByVal DateOf As String, _
    ByVal userRow As Long _
)
        
    Dim sectionHead As String
    sectionHead = hFind("COMM. SERVICE", "AGGREGATES")
    
    If isNotEmptyOrZero(Range(headerFind("Amount Filed #1", sectionHead) & userRow)) Then
        MsgBox "Warning: Community Service is already started for this client. Did not update to avoid overwrtiting"
        Exit Sub
    End If
    
    Range(headerFind("Amount Filed #1", sectionHead) & userRow).value = Amount
    Range(headerFind("Did Youth Have Comm. Service?", sectionHead) & userRow).value = 1 'Yes
    Range(headerFind("Date", sectionHead) & userRow).value = DateOf
    Range(headerFind("Courtroom", sectionHead) & userRow).value = Lookup("Courtroom_Name")(Courtroom)
    Range(headerFind("DA", sectionHead) & userRow).value = Lookup("DA_Last_Name_Name")(DA)
    
    
    Call autoCalcCommServ(userRow)

End Sub

Sub updateCommService( _
    ByVal Courtroom As String, _
    ByVal DA As String, _
    ByVal userRow As Long, _
    Optional ByVal DateOf As String = "", _
    Optional ByVal amountFiled As String = "", _
    Optional ByVal amountEarned As String = "" _
)
    Dim i As Integer
    Dim sectionHead As String
    Dim bucketHead As String
    sectionHead = hFind("Comm. Service", "AGGREGATES")
    
    Range(headerFind("Did Youth Have Comm. Service?", sectionHead) & userRow).value = 1 'Yes
    
    If Not amountFiled = "" Then
        For i = 1 To NUM_COMM_SERVICE_FILED_BUCKETS
            If isEmptyOrZero(Range(headerFind("Amount Filed #" & i, sectionHead) & userRow)) Then
                bucketHead = headerFind("Amount Filed #" & i, sectionHead)
                
                Range(bucketHead & userRow).value = amountFiled
                Range(headerFind("Date", bucketHead) & userRow).value = DateOf
                Range(headerFind("Courtroom", bucketHead) & userRow).value = Lookup("Courtroom_Name")(Courtroom)
                Range(headerFind("DA", bucketHead) & userRow).value = Lookup("DA_Last_Name_Name")(DA)
                
                Exit For
                ' TODO Warning for exceeding buckets
            End If
        Next i
    End If

    
    If Not amountEarned = "" Then
        For i = 1 To NUM_COMM_SERVICE_EARNED_BUCKETS
            If isEmptyOrZero(Range(headerFind("Amount Earned #" & i, sectionHead) & userRow)) Then
                bucketHead = headerFind("Amount Earned #" & i, sectionHead)
                
                Range(bucketHead & userRow).value = amountEarned
                Range(headerFind("Date", bucketHead) & userRow).value = DateOf
                Range(headerFind("Courtroom", bucketHead) & userRow).value = Lookup("Courtroom_Name")(Courtroom)
                Range(headerFind("DA", bucketHead) & userRow).value = Lookup("DA_Last_Name_Name")(DA)
                
                Exit For
                ' TODO Warning for exceeding buckets
            End If
        Next i
    End If
    
    Call autoCalcCommServ(userRow)
    
End Sub

Sub autoCalcCommServ(ByVal userRow As Long)
    Dim i As Integer
    Dim sectionHead As String
    Dim bucketHead As String
    Dim dateOfLastEarning As String
    Dim dateOfFirstFiling As String
    Dim dateOfLastFiling As String
    sectionHead = hFind("Comm. Service", "AGGREGATES")
    
    'Calc Total Amount Filed
    Dim totalAmountFiled As Double
    totalAmountFiled = 0
    
    For i = 1 To NUM_COMM_SERVICE_FILED_BUCKETS
        If isNotEmptyOrZero(Range(headerFind("Amount Filed #" & i, sectionHead) & userRow)) Then
            bucketHead = headerFind("Amount Filed #" & i, sectionHead)
            totalAmountFiled = totalAmountFiled + CDbl(Range(bucketHead & userRow).value)
            dateOfLastFiling = Range(headerFind("Date", bucketHead) & userRow).value
            If i = 1 Then
                dateOfFirstFiling = Range(headerFind("Date", bucketHead) & userRow).value
            End If
        End If
    Next i
    
    Range(headerFind("Total Amount Filed", sectionHead) & userRow).value = totalAmountFiled
    
    
    'Calc Total Amount Remaining
    Dim totalAmountEarned As Double
    totalAmountEarned = 0
    
    For i = 1 To NUM_COMM_SERVICE_EARNED_BUCKETS
        If isNotEmptyOrZero(Range(headerFind("Amount Earned #" & i, sectionHead) & userRow)) Then
            bucketHead = headerFind("Amount Earned #" & i, sectionHead)
            totalAmountEarned = totalAmountEarned + CDbl(Range(bucketHead & userRow).value)
            dateOfLastPayment = Range(headerFind("Date", bucketHead) & userRow).value
        End If
    Next i
    
    Range(headerFind("Total Amount Earned", sectionHead) & userRow).value = totalAmountEarned
    
    Range(headerFind("Total Amount Remaining", sectionHead) & userRow).value = totalAmountFiled - totalAmountEarned
    
    If totalAmountEarned >= totalAmountFiled Then
        Range(headerFind("Comm. Service Status", sectionHead) & userRow).value = 1 ' Paid in Full
        Range(headerFind("Date Earned in Full", sectionHead) & userRow).value = dateOfLastEarning
        Range(headerFind("LOS to Earn in Full", sectionHead) & userRow).value = calcLOS(dateOfFirstFiling, dateOfLastPayment)
    Else
        Range(headerFind("Comm. Service Status", sectionHead) & userRow).value = 3 ' Active and Unpaid
        Range(headerFind("Date Earned in Full", sectionHead) & userRow).value = ""
        Range(headerFind("LOS to Earn in Full", sectionHead) & userRow).value = ""
    End If
End Sub

