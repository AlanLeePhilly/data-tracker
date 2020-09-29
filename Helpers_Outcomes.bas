Attribute VB_Name = "Helpers_Outcomes"
Sub totalOutcome( _
    ByVal clientRow As Long, _
    ByVal DateOf As String, _
    ByVal Courtroom As String, _
    ByVal DA As String, _
    ByVal legalStatus As String, _
    ByVal Nature As String, _
    ByVal detailed As String, _
    Optional ByVal Notes As String = "" _
)
    Dim outcomeHead As String
    outcomeHead = hFind("Petition Outcomes", "AGGREGATES")
    
    Range(headerFind("Notes on Outcome", outcomeHead) & clientRow).value = Notes
    Range(headerFind("Date of Overall Discharge", outcomeHead) & clientRow).value = DateOf
    Range(headerFind("Courtroom of Discharge", outcomeHead) & clientRow).value = Lookup("Courtroom_Name")(Courtroom)

    Range(headerFind("DA", outcomeHead) & clientRow).value = Lookup("DA_Last_Name_Name")(DA)
    Range(headerFind("Legal Status of Discharge", outcomeHead) & clientRow).value _
        = Lookup("Legal_Status_Name")(legalStatus)
    Range(headerFind("Active or Discharged", outcomeHead) & clientRow).value _
        = Lookup("Active_Name")("Discharged")

    Range(headerFind("Nature of Petition Outcome", outcomeHead) & clientRow).value _
        = Lookup("Nature_of_Discharge_Name")(Nature)
    Range(headerFind("Detailed Petition Outcome", outcomeHead) & clientRow).value _
        = Lookup("Detailed_Petition_Outcome_Name")(detailed)
        
    If detailed = "Judgment of Acquittal" Or detailed = "Petition Withdrawn" Then
        Range(headerFind("Acquittal or Supervision Discharge?", outcomeHead) & clientRow).value _
        = Lookup("Acquittal_or_Supervision_Discharge_Name")("Acquittal")
    Else
        Range(headerFind("Acquittal or Supervision Discharge?", outcomeHead) & clientRow).value _
        = Lookup("Acquittal_or_Supervision_Discharge_Name")("Completion of Terms")
    End If
    
    
    Range(headerFind("Acquittal or Supervision Discharge?", outcomeHead) & clientRow).value _
        = Lookup("Acquittal_or_Supervision_Discharge_Name")("Completion of Terms")

    Range(headerFind("Total LOS in System (from petition)", outcomeHead) & clientRow).value _
        = calcLOS(Range(hFind("Date Filed", "Petition #1", "JUVENILE PETITION") & clientRow).value, DateOf)
    Range(headerFind("Total LOS From Arrest", outcomeHead) & clientRow).value _
        = calcLOS(Range(hFind("Arrest Date") & clientRow).value, DateOf)
    
    If Range(hFind("Restitution Status", "AGGREGATES") & clientRow).value = 1 Then 'paid in full
        Range(hFind("LOS from Pay in Full to Discharge", "Restitution Status", "AGGREGATES") & clientRow).value _
            = calcLOS(Range(hFind("Date Paid in Full", "Restitution Status", "AGGREGATES") & clientRow).value, DateOf)
    End If
    If Range(hFind("Court Cost Status", "AGGREGATES") & clientRow).value = 1 Then 'paid in full
        Range(hFind("LOS from Pay in Full to Discharge", "Court Cost Status", "AGGREGATES") & clientRow).value _
            = calcLOS(Range(hFind("Date Paid in Full", "Court Cost Status", "AGGREGATES") & clientRow).value, DateOf)
    End If

        
    Select Case Courtroom
        Case "4G", "4E", "6F", "6H", "3E"
            outcomeHead = hFind("OUTCOMES", Courtroom)
            
            Range(headerFind("Notes on Outcome", outcomeHead) & clientRow).value = Notes
            Range(headerFind("Date of Overall Discharge", outcomeHead) & clientRow).value = DateOf
            Range(headerFind("Courtroom of Discharge", outcomeHead) & clientRow).value = Lookup("Courtroom_Name")(Courtroom)
            Range(headerFind("Legal Status of Discharge", outcomeHead) & clientRow).value _
                = Lookup("Legal_Status_Name")(legalStatus)
                
            Range(headerFind("DA", outcomeHead) & clientRow).value = Lookup("DA_Last_Name_Name")(DA)
            Range(headerFind("Active or Discharged", outcomeHead) & clientRow).value _
                = Lookup("Active_Name")("Discharged")
        
            Range(headerFind("Nature of Courtroom Outcome", outcomeHead) & clientRow).value _
                = Lookup("Nature_of_Discharge_Name")(Nature)
            Range(headerFind("Detailed Courtroom Outcome", outcomeHead) & clientRow).value _
                = Lookup("Detailed_Petition_Outcome_Name")(detailed)
            Range(headerFind("Acquittal or Supervision Discharge?", outcomeHead) & clientRow).value _
                = Lookup("Acquittal_or_Supervision_Discharge_Name")("Completion of Terms")
        
            Range(headerFind("Total LOS in " & Courtroom, outcomeHead) & clientRow).value _
                = calcLOS(Range(hFind("Start Date", Courtroom) & clientRow).value, DateOf)
            Range(headerFind("Total LOS From Arrest", outcomeHead) & clientRow).value _
                = calcLOS(Range(hFind("Arrest Date") & clientRow).value, DateOf)
                
            Range(hFind("End Date", Courtroom) & clientRow).value = DateOf
            Range(hFind("LOS", Courtroom) & clientRow).value _
                = calcLOS(Range(hFind("Start Date", Courtroom) & clientRow).value, DateOf)
        
    End Select
    

    'TODO Confirm FoH values
    Range(headerFind("Next Court Date") & clientRow).Clear
    Range(headerFind("Listing Type") & clientRow).value = 0 'N/A
    Range(headerFind("Petition D/C Date") & clientRow).value = DateOf
    Range(headerFind("Active or Discharged (in courtroom)?") & clientRow).value = 2 'Discharged
    Range(headerFind("Legal Status") & clientRow).value = 0 'N/A
    Range(headerFind("Active Courtroom") & clientRow).value = 0 'N/A
    Range(headerFind("Active Supervision") & clientRow).value = 0 'N/A
    Range(headerFind("Active Supervision Provider") & clientRow).value = 0 'N/A
    Range(headerFind("IOP Provider") & clientRow).value = 0 'N/A
    Range(headerFind("LOS (discharged)") & clientRow).value _
        = calcLOS(Range(hFind("Arrest Date") & clientRow).value, DateOf)
End Sub

Sub CourtroomOutcome( _
    ByVal clientRow As Long, _
    ByVal DateOf As String, _
    ByVal Courtroom As String, _
    ByVal legalStatus As String, _
    ByVal DA As String, _
    ByVal Nature As String, _
    ByVal detailed As String, _
    Optional ByVal Notes As String = "" _
)
    Dim outcomeHead As String

    Select Case Courtroom
        Case "4G", "4E", "6F", "6H", "3E"
            outcomeHead = hFind("OUTCOMES", Courtroom)
            Range(headerFind("Notes on Outcome", outcomeHead) & clientRow).value = Notes
            Range(headerFind("Date of Overall Discharge", outcomeHead) & clientRow).value = DateOf

            Range(headerFind("Legal Status of Discharge", outcomeHead) & clientRow).value _
                = Lookup("Legal_Status_Name")(legalStatus)
            Range(headerFind("DA", outcomeHead) & clientRow).value = Lookup("DA_Last_Name_Name")(DA)

            Range(headerFind("Active or Discharged", outcomeHead) & clientRow).value _
                = Lookup("Active_Name")("Discharged")
            Range(headerFind("Nature of Courtroom Outcome", outcomeHead) & clientRow).value _
                = Lookup("Nature_of_Discharge_Name")(Nature)
            Range(headerFind("Detailed Courtroom Outcome", outcomeHead) & clientRow).value _
                = Lookup("Detailed_Petition_Outcome_Name")(detailed)
            Range(headerFind("Acquittal or Supervision Discharge?", outcomeHead) & clientRow).value _
                = Lookup("Acquittal_or_Supervision_Discharge_Name")("Completion of Terms")
            Range(headerFind("Total LOS in " & Courtroom, outcomeHead) & clientRow).value _
                = calcLOS(Range(hFind("Date Filed", "Petition #1") & clientRow).value, DateOf)
            Range(headerFind("Total LOS From Arrest", outcomeHead) & clientRow).value _
                = calcLOS(Range(hFind("Arrest Date") & clientRow).value, DateOf)
        Case Else
            MsgBox "Whoops! That Courtroom's outcomes have not been programmed yet"
    End Select


End Sub

