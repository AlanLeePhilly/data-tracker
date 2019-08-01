Attribute VB_Name = "Helpers_Outcomes"
Sub totalOutcome( _
    ByVal clientRow As Long, _
    ByVal dateOf As String, _
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
    Range(headerFind("Date of Overall Discharge", outcomeHead) & clientRow).value = dateOf
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
    Range(headerFind("Acquittal or Supervision Discharge?", outcomeHead) & clientRow).value _
        = Lookup("Acquittal_or_Supervision_Discharge_Name")("Completion of Terms")

    Range(headerFind("Total LOS in System (from petition)", outcomeHead) & clientRow).value _
        = calcLOS(Range(hFind("Date Filed", "Petition #1") & clientRow).value, dateOf)
    Range(headerFind("Total LOS From Arrest", outcomeHead) & clientRow).value _
        = calcLOS(Range(hFind("Arrest Date") & clientRow).value, dateOf)

    'TODO Confirm FoH values
    Range(headerFind("Next Court Date") & clientRow).Clear
    Range(headerFind("Listing Type") & clientRow).value = 0 'N/A
    Range(headerFind("Petition D/C Date") & clientRow).value = dateOf
    Range(headerFind("Active or Discharged (in courtroom)?") & clientRow).value = 2 'Discharged
    Range(headerFind("Legal Status") & clientRow).value = 0 'N/A
    Range(headerFind("Active Courtroom") & clientRow).value = 0 'N/A
    Range(headerFind("Active Supervision") & clientRow).value = 0 'N/A
    Range(headerFind("Active Supervision Provider") & clientRow).value = 0 'N/A
    Range(headerFind("IOP Provider") & clientRow).value = 0 'N/A
    Range(headerFind("LOS (discharged)") & clientRow).value _
        = calcLOS(Range(hFind("Arrest Date") & clientRow).value, dateOf)
End Sub

Sub CourtroomOutcome( _
    ByVal clientRow As Long, _
    ByVal dateOf As String, _
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
            Range(headerFind("Date of Overall Discharge", outcomeHead) & clientRow).value = dateOf

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
                = calcLOS(Range(hFind("Date Filed", "Petition #1") & clientRow).value, dateOf)
            Range(headerFind("Total LOS From Arrest", outcomeHead) & clientRow).value _
                = calcLOS(Range(hFind("Arrest Date") & clientRow).value, dateOf)
        Case Else
            MsgBox "Whoops! That Courtroom's outcomes have not been programmed yet"
    End Select


End Sub

