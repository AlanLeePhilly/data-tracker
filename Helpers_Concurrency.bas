Attribute VB_Name = "Helpers_Concurrency"
Sub CheckForConcurrency(userRow As Long, dateOf As String)
    Worksheets("Entry").Activate
    Dim lastRow As Long
    Dim i As Long, hasAtLeastOne As Boolean
    
    hasAtLeastOne = False
    lastRow = Range("C" & Rows.count).End(xlUp).row
    
    For i = 3 To lastRow
        If Range(hFind("PID #") & i).value = Range(hFind("PID #") & userRow).value _
            And Range(hFind("Active Courtroom") & i).value = Range(hFind("Active Courtroom") & userRow).value _
            And Not i = userRow _
            And Range(hFind("Active or Discharged (in courtroom)?") & i).value = 1 Then
            
            hasAtLeastOne = True
            Call addUserToBox(i, dateOf)
        End If
    Next i
    
    If hasAtLeastOne Then
        Concurrency.Show
    End If
End Sub

Sub addUserToBox(userRow As Long, dateOf As String)
    Worksheets("Entry").Activate
    
    Dim hasUpdateForToday As Boolean
    hasUpdateForToday = False
    Dim i As Integer
    
    For i = 1 To 100
        If Range(hFind("Court Date #" & i, "LISTINGS") & userRow).value = dateOf Then
            hasUpdateForToday = True
        End If
    Next i

    With Concurrency.RowBox
        .ColumnCount = 6
                    .ColumnWidths = "0;20;40;55;100;0;"
                    .AddItem userRow
                    
                        If hasUpdateForToday Then
                            .List(.ListCount - 1, 1) = "*"
                        End If
                        .List(.ListCount - 1, 2) = Range(hFind("DC #") & userRow).value
                        .List(.ListCount - 1, 3) = Range(hFind("Arrest Date (current petition)") & userRow).value
                        .List(.ListCount - 1, 4) = Range(hFind("Lead Charge Name") & userRow).value
                        .List(.ListCount - 1, 5) = dateOf
     End With
End Sub
