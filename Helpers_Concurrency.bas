Attribute VB_Name = "Helpers_Concurrency"
Sub CheckForConcurrency(userRow As Long, DateOf As String)
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
            Call addUserToBox(i, DateOf)
        End If
    Next i

    If hasAtLeastOne Then
        Concurrency.Show
    End If
End Sub

Function isActiveInCourtroom(userRow As Long, Courtroom As String, DateOf As String) As String
    Worksheets("Entry").Activate
    Dim lastRow As Long
    Dim i As Long, hasAtLeastOne As Boolean
    Dim currentPID As String
    Dim courtroomHead As String
    
    isActiveInCourtroom = 2 'No
    lastRow = Range("C" & Rows.count).End(xlUp).row
    currentPID = Range(hFind("PID #") & userRow).value
    courtroomHead = getCourtroomHead(Courtroom)
    
    If isReferralCourtroom(Courtroom) Then
        For i = 3 To lastRow
            If Range(hFind("PID #") & i).value = currentPID _
            And Not i = userRow Then
                If IsDate(Range(headerFind("Referral Date", courtroomHead) & i).value) Then
                    If Range(headerFind("Referral Date", courtroomHead) & i).value < DateOf Then
                        If Not IsDate(Range(headerFind("End Date", courtroomHead) & i).value) _
                        Or Range(headerFind("End Date", courtroomHead) & i).value > DateOf Then
                            isActiveInCourtroom = 1 'Yes
                        Else
                            MsgBox "'Active in Courtroom' debug: Found same client at row " & i & " but End Date in " & Courtroom & " is before " & DateOf
                        End If
                    Else
                        MsgBox "'Active in Courtroom' debug: Found same client at row " & i & " but Referral Date in " & Courtroom & " is not before " & DateOf
                    End If
                Else
                    MsgBox "'Active in Courtroom' debug: Found same client at row " & i & " but no Referral Date in " & Courtroom
                End If
            End If
        Next i
    Else
        For i = 3 To lastRow
            If Range(hFind("PID #") & i).value = currentPID _
            And Not i = userRow Then
                If IsDate(Range(headerFind("Start Date", courtroomHead) & i).value) Then
                    If Range(headerFind("Start Date", courtroomHead) & i).value < DateOf Then
                        If Not IsDate(Range(headerFind("End Date", courtroomHead) & i).value) _
                        Or Range(headerFind("End Date", courtroomHead) & i).value > DateOf Then
                            isActiveInCourtroom = 1 'Yes
                            MsgBox "'Active in Courtroom' logged: Found same client at row " & i & " in courtroom " & Courtroom & " on " & DateOf
                        Else
                            MsgBox "'Active in Courtroom' debug: Found same client at row " & i & " but End Date in " & Courtroom & " is before " & DateOf
                        End If
                    Else
                        MsgBox "'Active in Courtroom' debug: Found same client at row " & i & " but Start Date in " & Courtroom & " is not before " & DateOf
                    End If
                Else
                    MsgBox "'Active in Courtroom' debug: Found same client at row " & i & " but no Start Date in " & Courtroom
                End If
            End If
        Next i
    End If
End Function

Sub addUserToBox(userRow As Long, DateOf As String)
    Worksheets("Entry").Activate

    Dim hasUpdateForToday As Boolean
    hasUpdateForToday = False
    Dim i As Integer

    For i = 1 To 100
        If Range(hFind("Court Date #" & i, "LISTINGS") & userRow).value = DateOf Then
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
        .List(.ListCount - 1, 5) = DateOf
    End With
End Sub
