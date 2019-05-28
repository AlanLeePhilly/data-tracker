Attribute VB_Name = "Helpers_Forms"
Sub RearrestIntake(userRow As Long, rearrestNum As Long)
    Worksheets("Entry").Activate

    '''Demographics
    Load NewClientForm

    With NewClientForm

        .FirstName.value = Range(headerFind("First Name") & userRow).value
        .LastName.value = Range(headerFind("Last Name") & userRow).value
        .DateOfBirth.value = Range(headerFind("DOB") & userRow).value
        .Race.value = Lookup("Race_Num")(Range(headerFind("Race") & userRow).value)
        .Sex.value = Lookup("Sex_Num")(Range(headerFind("Sex") & userRow).value)
        .Latino.value = Lookup("Latino_Num")(Range(headerFind("Latino/Not Latino") & userRow).value)

        '''Community

        .GuardianFirstName.value = Range(headerFind("Guardian First") & userRow).value
        .GuardianLastName.value = Range(headerFind("Guardian Last") & userRow).value

        .Address.value = Range(headerFind("Address") & userRow).value
        .Zipcode.value = Range(headerFind("Zipcode") & userRow).value


        .PhoneNumber.value = Range(headerFind("Phone #") & userRow).value
        .School.value = Range(headerFind("School") & userRow).value
        .Grade.value = Range(headerFind("Grade") & userRow).value

        '''Incident and Arrest
        Dim bucketHead As String

        bucketHead = hFind("Arrest Date #" & rearrestNum, "REARRESTS", "AGGREGATES")

        .IncidentDate.value = Range(headerFind("Incident Date", bucketHead) & userRow).value
        .TimeOfIncident_H.value = getHour(Range(headerFind("Time of Incident", bucketHead) & userRow).value)
        .TimeOfIncident_M.value = getMinute(Range(headerFind("Time of Incident", bucketHead) & userRow).value)
        .TimeOfIncident_P.value = getPeriod(Range(headerFind("Time of Incident", bucketHead) & userRow).value)
        .IncidentDistrict.value = Range(headerFind("Incident District", bucketHead) & userRow).value
        .IncidentAddress.value = Range(headerFind("Incident Address", bucketHead) & userRow).value
        .IncidentZipcode.value = Range(headerFind("Incident Zipcode", bucketHead) & userRow).value

        .ArrestDate.value = Range(bucketHead & userRow).value
        .TimeOfArrest_H.value = getHour(Range(headerFind("Time of Arrest", bucketHead) & userRow).value)
        .TimeOfArrest_M.value = getMinute(Range(headerFind("Time of Arrest", bucketHead) & userRow).value)
        .TimeOfArrest_P.value = getPeriod(Range(headerFind("Time of Arrest", bucketHead) & userRow).value)

        .TimeReferredToDA_H.value = getHour(Range(headerFind("Time of Referral to DA", bucketHead) & userRow).value)
        .TimeReferredToDA_M.value = getMinute(Range(headerFind("Time of Referral to DA", bucketHead) & userRow).value)
        .TimeReferredToDA_P.value = getPeriod(Range(headerFind("Time of Referral to DA", bucketHead) & userRow).value)
        .ArrestingDistrict.value = Range(headerFind("Arresting District", bucketHead) & userRow).value

        .ActiveAtArrest.value = "Yes"
        .NumOfPriorArrests.value = Lookup("Num_Prior_Arrests_Num")(Range(headerFind("# of Prior Arrests") & userRow).value)
        .DCNum.value = Range(headerFind("DC #", bucketHead) & userRow).value
        .PIDNum.value = Range(headerFind("PID #", bucketHead) & userRow).value
        .SIDNum.value = Range(headerFind("SID #", bucketHead) & userRow).value

        .Officer1.value = Range(headerFind("Officer #1", bucketHead) & userRow).value
        .Officer2.value = Range(headerFind("Officer #2", bucketHead) & userRow).value
        .Officer3.value = Range(headerFind("Officer #3", bucketHead) & userRow).value
        .Officer4.value = Range(headerFind("Officer #4", bucketHead) & userRow).value
        .Officer4.value = Range(headerFind("Officer #5", bucketHead) & userRow).value

        .VictimFirstName.value = Range(headerFind("Victim First Name", bucketHead) & userRow).value
        .VictimLastName.value = Range(headerFind("Victim Last Name", bucketHead) & userRow).value


        .DA.value = Lookup("DA_Last_Name_Num")(Range(headerFind("DA", bucketHead) & userRow).value)

        .GeneralNotes.value = Range(headerFind("General Notes from Intake", bucketHead) & userRow).value

        Dim i As Integer
        Dim j As Integer
        Dim subBucketHead As String

        For i = 1 To 5
            If isNotEmptyOrZero(Range(headerFind("Petition #" & i, bucketHead) & userRow)) Then
                subBucketHead = headerFind("Petition #" & i, bucketHead)

                With .PetitionBox
                    .ColumnCount = 7
                    .ColumnWidths = "50;50;30;50;65;50;0"
                    ' 0 Date Filed
                    ' 1 Petition Number
                    ' 2 Charge Grade
                    ' 3 Charge Group
                    ' 4 Charge Code
                    ' 5 Charge Name
                    ' 6 Was Petition from other county?
                    .AddItem Range(headerFind("Date Filed", subBucketHead) & userRow).value
                    .List(.ListCount - 1, 0) = Range(headerFind("Date Filed", subBucketHead) & userRow).value
                    .List(.ListCount - 1, 1) = Range(subBucketHead & userRow).value
                    .List(.ListCount - 1, 2) = Lookup("Charge_Grade_Specific_Num")(Range(headerFind("Charge Grade (specific) #1", subBucketHead) & userRow).value)
                    .List(.ListCount - 1, 3) = Lookup("Charge_Num")(Range(headerFind("Charge Category #1", subBucketHead) & userRow).value)
                    .List(.ListCount - 1, 4) = Range(headerFind("Lead Charge Code", subBucketHead) & userRow).value
                    .List(.ListCount - 1, 5) = Range(headerFind("Lead Charge Name", subBucketHead) & userRow).value
                    .List(.ListCount - 1, 6) = Lookup("Generic_YNOU_Num")(Range(headerFind("Was Petition Transferred from Other County?", subBucketHead) & userRow).value)
                End With

                For j = 2 To 5
                    If isNotEmptyOrZero(Range(headerFind("Charge Code #" & j, subBucketHead) & userRow)) Then
                        With NewClientForm.ChargeBox
                            .ColumnCount = 5
                            .ColumnWidths = "50;50;30;50;65;"
                            ' 0 Petition Number
                            ' 1 Charge Grade
                            ' 2 Charge Group (specific)
                            ' 3 Charge Code
                            ' 4 Charge Name
                            .AddItem Range(subBucketHead & userRow).value
                            .List(.ListCount - 1, 0) = Range(subBucketHead & userRow).value
                            .List(.ListCount - 1, 1) = Lookup("Charge_Grade_Specific_Num")(Range(headerFind("Charge Grade (specific) #" & j, subBucketHead) & userRow).value)
                            .List(.ListCount - 1, 2) = Lookup("Charge_Num")(Range(headerFind("Charge Category #" & j, subBucketHead) & userRow).value)
                            .List(.ListCount - 1, 3) = Range(headerFind("Charge Code #" & j, subBucketHead) & userRow).value
                            .List(.ListCount - 1, 4) = Range(headerFind("Charge Name #" & j, subBucketHead) & userRow).value
                        End With
                    End If
                Next j
            End If
        Next i
    End With

    NewClientForm.Show
End Sub
