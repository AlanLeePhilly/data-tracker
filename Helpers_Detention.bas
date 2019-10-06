Attribute VB_Name = "Helpers_Detention"
Sub initialDetention( _
    ByVal clientRow As Long, _
    ByVal DateOf As String, _
    ByVal DA As String, _
    ByVal DA_Action As String, _
    ByVal DA_actionAccepted As String, _
    ByVal decision As String, _
    ByVal Facility As String, _
    ByVal Re1 As String, _
    ByVal Re2 As String, _
    ByVal Re3 As String, _
    ByVal Re4 As String, _
    ByVal Re5 As String)

    Dim bucketHead As String

    bucketHead = hFind("DETENTION")

    Call flagYes(Range(headerFind("Did Youth Have Initial Detention Hearing?", bucketHead) & clientRow))
    Range(headerFind("Date of Initial Detention Hearing", bucketHead) & clientRow).value = DateOf
    Range(headerFind("Type of Detention Hearing", bucketHead) & clientRow).value _
        = Lookup("Type_of_Detention_Hearing_Name")("Initial")
    Range(headerFind("DA", bucketHead) & clientRow).value _
        = Lookup("DA_Last_Name_Name")(DA)
    Range(headerFind("DA Action", bucketHead) & clientRow).value _
        = Lookup("DA_Action_Name")(DA_Action)
    Range(headerFind("DA Action Accepted?", bucketHead) & clientRow).value _
        = Lookup("Generic_YNOU_Name")(DA_actionAccepted)
    Range(headerFind("Detention Decision", bucketHead) & clientRow).value _
        = Lookup("Detention_Decision_Name")(decision)
    Range(headerFind("Detention Facility", bucketHead) & clientRow).value _
        = Lookup("Detention_Facility_Name")(Facility)

    Range(headerFind("Reason #1 for Detention Commit", bucketHead) & clientRow).value _
        = Lookup("Detention_Hearing_Reason_Name")(Re1)
    Range(headerFind("Reason #2 for Detention Commit", bucketHead) & clientRow).value _
        = Lookup("Detention_Hearing_Reason_Name")(Re2)
    Range(headerFind("Reason #3 for Detention Commit", bucketHead) & clientRow).value _
        = Lookup("Detention_Hearing_Reason_Name")(Re3)
    Range(headerFind("Reason #4 for Detention Commit", bucketHead) & clientRow).value _
        = Lookup("Detention_Hearing_Reason_Name")(Re4)
    Range(headerFind("Reason #5 for Detention Commit", bucketHead) & clientRow).value _
        = Lookup("Detention_Hearing_Reason_Name")(Re5)
End Sub
