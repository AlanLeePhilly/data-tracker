VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Adult_Reslate_Juvenile_Petition 
   Caption         =   "UserForm1"
   ClientHeight    =   7575
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18990
   OleObjectBlob   =   "Adult_Reslate_Juvenile_Petition.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Adult_Reslate_Juvenile_Petition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public NCF_userRow As Long
Public NCF_rearrestNum As Long

Private Sub ConfOutcome_Change()
    Select Case ConfOutcome.value
        Case "Release for Diversion"
            DiversionProgram.value = "Yes"
        Case Else
            DiversionProgram.value = "No"
    End Select
End Sub

Private Sub DRAI_Score_Change()
    If IsNumeric(DRAI_Score.value) Then
        Select Case DRAI_Score.value
            Case Is < 10
                DRAI_Rec.value = "Release"
            Case Is < 15
                DRAI_Rec.value = "Release w/ Supervision"
            Case Is >= 15
                DRAI_Rec.value = "Hold"
            Case Else
                DRAI_Rec.value = "Unknown"
        End Select
    End If
End Sub
Private Sub DRAI_Action_Change()
    Select Case DRAI_Action.value
        Case "Follow - Hold", "Override - Hold"
            InitialHearingLocation.value = "PJJSC"
            DetentionFacility.Enabled = True
            DetentionFacilityLabel.Enabled = True

            Supv1.Enabled = False
            Supv1Pro.Enabled = False
            Supv1Re1.Enabled = False
            Supv1Re2.Enabled = False
            Supv1Re3.Enabled = False

            Supv2.Enabled = False
            Supv2Pro.Enabled = False
            Supv2Re1.Enabled = False
            Supv2Re2.Enabled = False
            Supv2Re3.Enabled = False

            Cond1.Enabled = False
            Cond1Pro.Enabled = False
            Cond2.Enabled = False
            Cond2Pro.Enabled = False
            Cond3.Enabled = False
            Cond3Pro.Enabled = False

        Case Else
            Supv1.Enabled = True
            Supv1Pro.Enabled = True
            Supv1Re1.Enabled = True
            Supv1Re2.Enabled = True
            Supv1Re3.Enabled = True

            Supv2.Enabled = True
            Supv2Pro.Enabled = True
            Supv2Re1.Enabled = True
            Supv2Re2.Enabled = True
            Supv2Re3.Enabled = True

            Cond1.Enabled = True
            Cond1Pro.Enabled = True
            Cond2.Enabled = True
            Cond2Pro.Enabled = True
            Cond3.Enabled = True
            Cond3Pro.Enabled = True

    End Select
End Sub



Private Sub InConfDate_Enter()
    InConfDate.value = CalendarForm.GetDate(RangeOfYears:=5)
End Sub
Private Sub InConfDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Set ctl = Me.InConfDate
    Call DateValidation(ctl, Cancel)
End Sub
Private Sub CallInDate_Enter()
    CallInDate.value = CalendarForm.GetDate(RangeOfYears:=5)
End Sub
Private Sub CallInDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Set ctl = Me.CallInDate
    Call DateValidation(ctl, Cancel)
End Sub



Private Sub InitialHearingLocation_Change()
    If InitialHearingLocation.value = "Intake Conf." Then
        MsgBox "Not a valid value for this prompt"
        InitialHearingLocation.value = "N/A"
        Exit Sub
    End If
End Sub


Private Sub UserForm_Initialize()
    Me.ScrollTop = 0
    DetentionFacilityLabel.Enabled = False
    DetentionFacility.Enabled = False
End Sub

Private Sub AddPetition_Click()
    If PetitionBox.ListCount < 5 Then
        Load Modal_NewClient_Add_Petition
        Modal_NewClient_Add_Petition.headline.Caption = "Reslate"
        Modal_NewClient_Add_Petition.Show
    Else
        MsgBox "Maximum of five petitions for reslate"
    End If
End Sub


Private Sub DeletePetition_Click()
    Dim petitionNum As String
    Dim i As Integer
    Dim listIndex As Integer

    If PetitionBox.listIndex = -1 Then
        Exit Sub
    End If

    petitionNum = PetitionBox.List(PetitionBox.listIndex, 1)

    MsgBox "Removing Petition #" & petitionNum
    listIndex = ChargeBox.ListCount - 1
    For i = listIndex To 0 Step -1
        If ChargeBox.List(i, 0) = petitionNum Then
            ChargeBox.RemoveItem (i)
        End If
    Next i
    PetitionBox.RemoveItem (PetitionBox.listIndex)
End Sub


Private Sub DiversionProgramReferralDate_Enter()
    DiversionProgramReferralDate.value = CalendarForm.GetDate(RangeOfYears:=5)
End Sub
Private Sub DiversionProgramReferralDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Set ctl = Me.DiversionProgramReferralDate
    Call DateValidation(ctl, Cancel)
End Sub

Private Sub Cancel_Click()
    Call Clear_Click
    Adult_Reslate_Juvenile_Petition.Hide
End Sub

Private Sub Clear_Click()
    Dim ctl As Control ' Removed MSForms.

    For Each ctl In Me.Controls
        Select Case TypeName(ctl)
            Case "TextBox"
                ctl.value = ""
            Case "CheckBox", "ToggleButton" ' Removed OptionButton
                ctl.value = False
            Case "OptionGroup" ' Add OptionGroup
                ctl = Null
            Case "OptionButton" ' Add OptionButton
                ' Do not reset an optionbutton if it is part of an OptionGroup
                If TypeName(ctl.Parent) <> "OptionGroup" Then ctl.value = False
            Case "ComboBox", "ListBox"
                ctl.listIndex = -1
        End Select
    Next ctl
    Call UserForm_Initialize

End Sub

Private Sub DiversionProgram_Change()
    Select Case DiversionProgram.value
        Case "No"

            DiversionProgramReferralDateLabel.Enabled = False
            DiversionProgramReferralDate.Enabled = False
            DiversionProgramReferralDate.value = ""

            ReferralSource.Enabled = False
            ReferralSource.value = "N/A"
            ReferralSourceLabel.Enabled = False

            NameOfProgramLabel.Enabled = False
            NameOfProgram.Enabled = False
            NameOfProgram.value = "N/A"

            YAPDistrictLabel.Enabled = False
            YAPDistrict.Enabled = False
            YAPDistrict.value = ""

            NoDiversionReason1.Enabled = True
            NoDiversionReason2.Enabled = True
            NoDiversionReason3.Enabled = True
        Case Else

            DiversionProgramReferralDateLabel.Enabled = True
            DiversionProgramReferralDate.Enabled = True
            DiversionProgramReferralDate.value = InitialHearingDate.value

            ReferralSource.Enabled = True
            ReferralSourceLabel.Enabled = True

            NameOfProgramLabel.Enabled = True
            NameOfProgram.Enabled = True

            YAPDistrictLabel.Enabled = True
            YAPDistrict.Enabled = True

            NoDiversionReason1.Enabled = False
            NoDiversionReason2.Enabled = False
            NoDiversionReason3.Enabled = False
            NoDiversionReason1.value = "N/A"
            NoDiversionReason2.value = "N/A"
            NoDiversionReason3.value = "N/A"
    End Select
End Sub

Private Sub NameOfProgram_Change()
    If NameOfProgram = "YAP" Then
        YAPDistrictLabel.Enabled = True
        YAPDistrict.Enabled = True
    Else
        YAPDistrictLabel.Enabled = False
        YAPDistrict.Enabled = False
        YAPDistrict.value = ""
    End If
End Sub

Private Sub Submit_Click()

    '''''''''''''
    'Validations'
    '''''''''''''

    If InConfDate.value = "" And InConfRecord.value = "Yes" Then
        MsgBox "Intake Date Required if record available"
        Exit Sub
    End If

    If ConfOutcome.value = "N/A" And InConfRecord.value = "Yes" Then
        MsgBox "Conference Outcome Required if record available"
        Exit Sub
    End If

    If CallInDate.value = "" And CallInRecord.value = "Yes" Then
        MsgBox "Call-in Date required if record available"
        Exit Sub
    End If

    If PetitionBox.ListCount = 0 Then
        MsgBox "Petition required"
        Exit Sub
    End If
    
    If GunCase.value = "" Then
        MsgBox "'Gun Case?' required"
        Exit Sub
    End If
    
    If GunInvolved.value = "" Then
        MsgBox "'Gun Involved?' required"
        Exit Sub
    End If
    
    If DirectFiled.value = "" Then
        MsgBox "'Direct Filed?' required"
        Exit Sub
    End If

    If DRAI_Action.value = "Follow - Hold" Or DRAI_Action.value = "Override - Hold" Then
        If DetentionFacility.value = "N/A" Then
            MsgBox "Detention facility required for call-in hold"
            Exit Sub
        End If
    End If
    
    If DiversionProgram.value = "No" And NoDiversionReason1.value = "N/A" Then
        MsgBox "Reason Not Diverted Required"
        Exit Sub
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''
    '''''MOVE TO CLIENTUPDATE ADULT_SUBMIT'''''
    '''''''''''''''''''''''''''''''''''''''''''

    Range(headerFind("Initial Hearing Location") & emptyRow).value _
            = Lookup("Courtroom_Name")(InitialHearingLocation.value)

    Range(headerFind("Active or Discharged (in courtroom)?") & emptyRow).value _
            = Lookup("Active_Name")("Active")

    Range(headerFind("Active Courtroom") & emptyRow).value _
            = Lookup("Courtroom_Name")(InitialHearingLocation.value)
    Range(hFind("Active B/W?") & emptyRow).value = Lookup("Generic_YNOU_Name")("No")
    'direct entry from textbox
    Range(headerFind("DOB") & emptyRow).value _
            = DateOfBirth.value
    Range(headerFind("Age @ Intake") & emptyRow).value _
            = ageAtTime(InitialHearingDate.value, emptyRow)
    Select Case ageAtTime(InitialHearingDate, emptyRow)
        Case Is < 12
            Range(headerFind("Age Group") & emptyRow).value _
                    = Lookup("Age_Group_Name")("<=11")
        Case Is < 15
            Range(headerFind("Age Group") & emptyRow).value _
                    = Lookup("Age_Group_Name")("12-14")
        Case Is < 18
            Range(headerFind("Age Group") & emptyRow).value _
                    = Lookup("Age_Group_Name")("15-17")
        Case Else
            Range(headerFind("Age Group") & emptyRow).value _
                    = Lookup("Age_Group_Name")("18+")
    End Select

    Range(headerFind("Sex") & emptyRow) = Lookup("Sex_Name")(Sex.value)
    Range(headerFind("Race") & emptyRow) = Lookup("Race_Name")(Race.value)
    Range(headerFind("Latino/Not Latino") & emptyRow) = Lookup("Latino_Name")(Latino.value)
    Range(headerFind("Guardian First") & emptyRow).value = GuardianFirstName.value
    Range(headerFind("Guardian Last") & emptyRow).value = GuardianLastName.value
    Range(headerFind("Address") & emptyRow).value = Address.value
    Range(headerFind("Zipcode") & emptyRow).value = Zipcode.value
    Range(headerFind("Phone #") & emptyRow).value = PhoneNumber.value
    Range(headerFind("School") & emptyRow).value = School.value
    Range(headerFind("Grade") & emptyRow).value = Grade.value

    'Finding address lat and lon
    Dim coords As Variant
    'Probably want to try-catch the whole block after this line
    coords = FindLatLon(Address.value, Zipcode.value)
    Range(headerFind("Latitude") & emptyRow).value = coords(1)
    Range(headerFind("Longitude") & emptyRow).value = coords(2)
    Range(headerFind("Zipcode") & emptyRow).value = coords(3)

    If Not StrComp(Zipcode.value, coords(3)) = 0 Then
        MsgBox ("ALERT: The zipcode entered and zipcode found by geolocating services are different. Please check the new zipcode entered to make sure it is correct.")
    End If
    
    
    ''''''''''''
    ''PETITION''
    ''''''''''''

    Dim petitionHead As String
    Dim count As Integer
    
    For count = 1 To 2
        If Not DirectFiled.value = "Yes" Then
            count = 2
        End If
    
        Select Case count
            Case 1
                petitionHead = hFind("ADULT PETITION")
            Case 2
                petitionHead = hFind("PETITION")
        End Select

        Range(headerFind("Initial Court Date", petitionHead) & emptyRow).value _
                = InitialHearingDate.value
        Range(headerFind("Initial Hearing Location", petitionHead) & emptyRow).value _
                = Lookup("Courtroom_Name")(InitialHearingLocation.value)
    
        If IsNumeric(NumOfPriorArrests.value) And Not NumOfPriorArrests.value = "10+" Then
            Range(headerFind("# of Prior Arrests") & emptyRow).value _
                = Lookup("Num_Prior_Arrests_Name")(CInt(NumOfPriorArrests.value))
        Else
            Range(headerFind("# of Prior Arrests") & emptyRow).value _
                = Lookup("Num_Prior_Arrests_Name")(NumOfPriorArrests.value)
        End If
    
        Range(headerFind("Active in System at Time of Arrest?") & emptyRow) = _
                Lookup("Generic_YNOU_Name")(ActiveAtArrest.value)
    
        Range(headerFind("Arrest Date", petitionHead) & emptyRow).value _
                = ArrestDate.value
        Range(headerFind("Day of Arrest", petitionHead) & emptyRow).value _
                = Weekday(ArrestDate.value, vbMonday) * 2 - 1
        Range(headerFind("Time of Arrest", petitionHead) & emptyRow).value _
                = TimeOfArrest_H.value & ":" & TimeOfArrest_M.value & " " & TimeOfArrest_P.value
        Range(headerFind("Time Category of Arrest", petitionHead) & emptyRow).value _
                = calcTimeGroup(TimeOfArrest_H.value, TimeOfArrest_P.value)
        Range(headerFind("Arresting District", petitionHead) & emptyRow).value _
                = ArrestingDistrict.value
        Range(headerFind("Time of Referral to DA") & emptyRow).value _
                = TimeReferredToDA_H.value & ":" & TimeReferredToDA_M.value & " " & TimeReferredToDA_P.value
    
    
        Range(headerFind("DC #", petitionHead) & emptyRow).value = DCNum.value
        Range(headerFind("PID #") & emptyRow).value = PIDNum.value
        Range(headerFind("DC-PID #") & emptyRow).value = DCNum.value & "-" & PIDNum.value
        Range(headerFind("SID #") & emptyRow).value = SIDNum.value
    
        Range(headerFind("Officer #1") & emptyRow).value = Officer1.value
        Range(headerFind("Officer #2") & emptyRow).value = Officer2.value
        Range(headerFind("Officer #3") & emptyRow).value = Officer3.value
        Range(headerFind("Officer #4") & emptyRow).value = Officer4.value
        Range(headerFind("Officer #5") & emptyRow).value = Officer5.value
    
        'confirm if needs to happen twice
        Range(headerFind("Victim First Name") & emptyRow) _
                = VictimFirstName.value
        Range(headerFind("Victim Last Name") & emptyRow) _
                = VictimLastName.value
    
        Range(headerFind("Incident Date", petitionHead) & emptyRow).value _
                = IncidentDate.value
        Range(headerFind("Day of Incident", petitionHead) & emptyRow).value _
                = Weekday(IncidentDate.value, vbMonday) * 2 - 1
        Range(headerFind("Time of Incident", petitionHead) & emptyRow).value _
                = TimeOfIncident_H.value & ":" & TimeOfIncident_M.value & " " & TimeOfIncident_P.value
        Range(headerFind("Time Category of Incident", petitionHead) & emptyRow).value _
                = calcTimeGroup(TimeOfIncident_H.value, TimeOfIncident_P.value)
        Range(headerFind("Incident District", petitionHead) & emptyRow).value _
                = IncidentDistrict.value
        Range(headerFind("Incident Address", petitionHead) & emptyRow).value _
                = IncidentAddress.value
        Range(headerFind("Incident Zipcode", petitionHead) & emptyRow).value _
                = IncidentZipcode.value
    
        'Finding address lat and lon
        Dim incidentCoords As Variant
        'Probably want to try-catch the whole block after this line
        incidentCoords = FindLatLon(IncidentAddress.value, IncidentZipcode.value)
        Range(headerFind("Latitude", petitionHead) & emptyRow).value = incidentCoords(1)
        Range(headerFind("Longitude", petitionHead) & emptyRow).value = incidentCoords(2)
        Range(headerFind("Incident Zipcode", petitionHead) & emptyRow).value = incidentCoords(3)
    
        If Not StrComp(IncidentZipcode.value, incidentCoords(3)) = 0 Then
            MsgBox ("ALERT: The zipcode entered and zipcode found by geolocating services are different. Please check the new zipcode entered to make sure it is correct.")
        End If
    
        Range(headerFind("DA") & emptyRow).value = Lookup("DA_Last_Name_Name")(DA.value)
        
    
        Range(headerFind("Gun Case?") & emptyRow).value = Lookup("Generic_YNOU_Name")(GunCase.value)
        Range(headerFind("Gun Involved Arrest?") & emptyRow).value = Lookup("Generic_YNOU_Name")(GunInvolved.value)
        
        Range(headerFind("General Notes from Intake") & emptyRow).value = GeneralNotes.value
    
        Dim Num As Long
        Dim i As Integer
        Dim j As Integer
    
        For Num = 1 To PetitionBox.ListCount
            tempHead = headerFind("Petition #" & Num, petitionHead)
            
            If DiversionProgram.value = "Yes" Or InitialHearingLocation.value = "Adult" And count = 2 Then
                Range(headerFind("Petition Filed?", tempHead) & emptyRow).value _
                        = Lookup("Generic_YNOU_Name")("No")
            Else
                Range(headerFind("Petition Filed?", tempHead) & emptyRow).value _
                        = Lookup("Generic_YNOU_Name")("Yes")
            End If
            Range(headerFind("Was Petition Transferred from Other County?", tempHead) & emptyRow).value _
                    = Lookup("Generic_YNOU_Name")(PetitionBox.List(Num - 1, 6))
            Range(tempHead & emptyRow).value _
                    = PetitionBox.List(Num - 1, 1)
            Range(headerFind("Date Filed", tempHead) & emptyRow).value _
                    = PetitionBox.List(Num - 1, 0)
            Range(headerFind("Lead Charge Code", tempHead) & emptyRow).value _
                    = PetitionBox.List(Num - 1, 4)
            Range(headerFind("Lead Charge Name", tempHead) & emptyRow).value _
                    = PetitionBox.List(Num - 1, 5)
            Range(headerFind("Charge Category #1", tempHead) & emptyRow).value _
                    = Lookup("Charge_Name")(PetitionBox.List(Num - 1, 3))
            Range(headerFind("Charge Grade (specific) #1", tempHead) & emptyRow).value _
                    = Lookup("Charge_Grade_Specific_Name")(PetitionBox.List(Num - 1, 2))
            Range(headerFind("Charge Grade (broad) #1", tempHead) & emptyRow).value _
                    = calcChargeBroad(PetitionBox.List(Num - 1, 2))
    
            j = 2
            For i = 0 To ChargeBox.ListCount - 1
                If ChargeBox.ListCount > 0 Then
                    If ChargeBox.List(i, 0) = PetitionBox.List(Num - 1, 1) Then
                        Range(headerFind("Charge Code #" & j, tempHead) & emptyRow).value _
                                = ChargeBox.List(i, 3)
                        Range(headerFind("Charge Name #" & j, tempHead) & emptyRow).value _
                                = ChargeBox.List(i, 4)
                        Range(headerFind("Charge Category #" & j, tempHead) & emptyRow).value _
                                = Lookup("Charge_Name")(ChargeBox.List(i, 2))
                        Range(headerFind("Charge Grade (specific) #" & j, tempHead) & emptyRow).value _
                                = Lookup("Charge_Grade_Specific_Name")(ChargeBox.List(i, 1))
                        Range(headerFind("Charge Grade (broad) #" & j, tempHead) & emptyRow).value _
                                = calcChargeBroad(ChargeBox.List(i, 1))
                        j = j + 1
                    End If
                End If
            Next i
        Next Num
    
        Range(headerFind("LOS Until DA Referral", petitionHead) & emptyRow).value _
                    = timeDiff(Range(headerFind("Time of Arrest", petitionHead) & emptyRow).value, _
               Range(headerFind("Time of Referral to DA") & emptyRow).value)
    Next count

    ''''''''''''''''''
    'SET LEGAL STATUS'
    ''''''''''''''''''

    Dim bucketHead As String

    If DiversionProgram.value = "Yes" Then
        Call startLegalStatus( _
            clientRow:=emptyRow, _
            statusType:="Diversion", _
            Courtroom:="PJJSC", _
            DA:=DA.value, _
            startDate:=DiversionProgramReferralDate.value)

    Else
        If ConfOutcome.value = "Hold for Detention" _
        Or ConfOutcome.value = "Roll to Detention Hearing" Then
            Call startLegalStatus( _
                clientRow:=emptyRow, _
                statusType:="Pretrial", _
                Courtroom:="PJJSC", _
                DA:=DA.value, _
                startDate:=PetitionBox.List(0, 0))
        Else
            If InitialHearingLocation.value = "Adult" Then
                Call startLegalStatus( _
                    clientRow:=emptyRow, _
                    statusType:="Adult", _
                    Courtroom:="Adult", _
                    DA:=DA.value, _
                    startDate:=ArrestDate.value)
            Else
                Call startLegalStatus( _
                    clientRow:=emptyRow, _
                    statusType:="Pretrial", _
                    Courtroom:="Intake Conf.", _
                    DA:=DA.value, _
                    startDate:=PetitionBox.List(0, 0))
            End If
        End If
    End If




    '''''''''''''''''''
    ''''''CALL IN''''''
    '''''''''''''''''''

    tempHead = headerFind("CALL-IN", petitionHead)

    If CallInRecord.value = "Yes" Then
        Range(headerFind("Did Youth Have Call-In?", tempHead) & emptyRow).value _
                = Lookup("Generic_NYNOU_Name")("Yes")

        Range(headerFind("Date of Call-In", tempHead) & emptyRow).value _
                    = CallInDate.value

        Range(headerFind("Was DRAI Administered?", tempHead) & emptyRow).value _
                = Lookup("Generic_NYNOU_Name")(Was_DRAI_Administered.value)

        Range(headerFind("DRAI Score", tempHead) & emptyRow).value _
                = DRAI_Score.value

        Select Case DRAI_Score.value
            Case Is < 10
                Range(hFind("DRAI Recommendation", "CALL-IN") & emptyRow).value _
                        = Lookup("DRAI_Recommendation_Name")("Release")
            Case Is < 15
                Range(hFind("DRAI Recommendation", "CALL-IN") & emptyRow).value _
                        = Lookup("DRAI_Recommendation_Name")("Release w/ Supervision")
            Case Is < 30
                Range(hFind("DRAI Recommendation", "CALL-IN") & emptyRow).value _
                        = Lookup("DRAI_Recommendation_Name")("Release w/ Supervision")
            Case Else
                Range(hFind("DRAI Recommendation", "CALL-IN") & emptyRow).value _
                        = Lookup("DRAI_Recommendation_Name")("Unknown")
        End Select
            
        
        Range(headerFind("DRAI Recommendation", tempHead) & emptyRow).value _
                = Lookup("DRAI_Recommendation_Name")(DRAI_Rec.value)

        Range(headerFind("DRAI Action", tempHead) & emptyRow).value _
                = Lookup("DRAI_Action_Name")(DRAI_Action.value)

        Select Case DRAI_Action.value
            Case "Override - Hold", "Follow - Hold"
                Range(headerFind("End Date", tempHead) & emptyRow).value _
                        = InConfDate.value
                Range(headerFind("LOS in Detention", tempHead) & emptyRow).value _
                        = calcLOS(CallInDate.value, InConfDate.value)
                Call addSupervision( _
                    clientRow:=emptyRow, _
                    serviceType:="Detention (not respite)", _
                    legalStatus:="Pretrial", _
                    Courtroom:="Call-In", _
                    CourtroomOfOrder:="Call-In", _
                    DA:=DA.value, _
                    agency:="PJJSC", _
                    startDate:=CallInDate.value, _
                    endDate:=InConfDate.value, _
                    Re1:="", _
                    Re2:="", _
                    Re3:="", _
                    Notes:="Held at call-in")

        End Select
        Range(headerFind("LOS Until Next Hearing", tempHead) & emptyRow).value _
                        = calcLOS(CallInDate.value, InConfDate.value)

        Range(headerFind("Detention Facility", tempHead) & emptyRow).value _
                    = Lookup("Detention_Facility_Name")(DetentionFacility.value)

        Range(headerFind("Reason #1 for Override Hold", tempHead) & emptyRow).value _
                = Lookup("DRAI_Override_Reason_Name")(OverrideHoldRe1.value)
        Range(headerFind("Reason #2 for Override Hold", tempHead) & emptyRow).value _
                = Lookup("DRAI_Override_Reason_Name")(OverrideHoldRe2.value)
        Range(headerFind("Reason #3 for Override Hold", tempHead) & emptyRow).value _
                = Lookup("DRAI_Override_Reason_Name")(OverrideHoldRe3.value)
    Else
        Range(headerFind("Did Youth Have Call-In?", tempHead) & emptyRow).value _
                = Lookup("Generic_NYNOU_Name")("Unknown")
    End If


    '''''''''''''''''''
    'Intake Conference'
    '''''''''''''''''''

    tempHead = headerFind("INTAKE CONFERENCE", petitionHead)



    If InConfRecord.value = "Yes" Then
        Range(headerFind("Did Youth Have an Intake Conference?", tempHead) & emptyRow).value _
                = Lookup("Generic_NYNOU_Name")("Yes")

        Range(headerFind("Date of Intake Conference", tempHead) & emptyRow).value _
                = InConfDate.value

        Range(headerFind("Intake Conference Type", tempHead) & emptyRow).value _
                = Lookup("Intake_Conference_Type_Name")(InConfType.value)

        Range(headerFind("DA", tempHead) & emptyRow).value _
                = Lookup("DA_Last_Name_Name")(DA.value)


        Range(headerFind("Intake Conference Outcome", tempHead) & emptyRow).value _
                = Lookup("Intake_Conference_Outcome_Name")(ConfOutcome.value)


        Range(headerFind("Location of Next Event", tempHead) & emptyRow).value _
                = Lookup("Courtroom_Name")(InitialHearingLocation.value)

        Range(headerFind("Next Event Date", tempHead) & emptyRow).value _
                = InitialHearingDate.value

        Range(headerFind("LOS from Arrest Until Conference", tempHead) & emptyRow).value _
                    = calcLOS(ArrestDate.value, InConfDate.value)

        tempHead = headerFind("Supervision Ordered #1", tempHead)

        Range(tempHead & emptyRow).value _
                = Lookup("Supervision_Program_Name")(Supv1.value)
        Range(headerFind("Community-Based Agency #1", tempHead) & emptyRow).value _
                = Lookup("Community_Based_Supervision_Provider_Name")(Supv1Pro.value)

        Range(headerFind("Reason #1 for Supervision Referral", tempHead) & emptyRow).value _
                = Lookup("Supervision_Referral_Reason_Name")(Supv1Re1.value)
        Range(headerFind("Reason #2 for Supervision Referral", tempHead) & emptyRow).value _
                = Lookup("Supervision_Referral_Reason_Name")(Supv1Re2.value)
        Range(headerFind("Reason #3 for Supervision Referral", tempHead) & emptyRow).value _
                = Lookup("Supervision_Referral_Reason_Name")(Supv1Re3.value)

        tempHead = headerFind("Supervision Ordered #2", tempHead)

        Range(tempHead & emptyRow).value _
                = Lookup("Supervision_Program_Name")(Supv2.value)
        Range(headerFind("Community-Based Agency #2", tempHead) & emptyRow).value _
                = Lookup("Community_Based_Supervision_Provider_Name")(Supv2Pro.value)


        Range(headerFind("Reason #1 for Supervision Referral", tempHead) & emptyRow).value _
                = Lookup("Supervision_Referral_Reason_Name")(Supv2Re1.value)
        Range(headerFind("Reason #2 for Supervision Referral", tempHead) & emptyRow).value _
                = Lookup("Supervision_Referral_Reason_Name")(Supv2Re2.value)
        Range(headerFind("Reason #3 for Supervision Referral", tempHead) & emptyRow).value _
                = Lookup("Supervision_Referral_Reason_Name")(Supv2Re3.value)

        Range(headerFind("Other Condition #1", tempHead) & emptyRow).value _
                = Lookup("Condition_Name")(Cond1.value)
        Range(headerFind("Other Condition #1 Provider", tempHead) & emptyRow).value _
                = Lookup("Condition_Provider_Name")(Cond1Pro.value)

        Range(headerFind("Other Condition #2", tempHead) & emptyRow).value _
                = Lookup("Condition_Name")(Cond2.value)
        Range(headerFind("Other Condition #2 Provider", tempHead) & emptyRow).value _
                = Lookup("Condition_Provider_Name")(Cond2Pro.value)

        Range(headerFind("Other Condition #3", tempHead) & emptyRow).value _
                = Lookup("Condition_Name")(Cond3.value)
        Range(headerFind("Other Condition #3 Provider", tempHead) & emptyRow).value _
                = Lookup("Condition_Provider_Name")(Cond3Pro.value)

        Select Case ConfOutcome.value
            Case "Hold for Detention"
                Range(headerFind("Active Courtroom") & emptyRow).value _
                         = Lookup("Courtroom_Name")("PJJSC")
                Call flagNo(Range(hFind("Did Youth Have Initial Detention Hearing?", "DETENTION") & emptyRow))
                Range(hFind("Detention Facility", "DETENTION") & emptyRow).value _
                         = Lookup("Detention_Facility_Name")(DetentionFacility.value)
                Call addSupervision( _
                    clientRow:=emptyRow, _
                    serviceType:="Detention (not respite)", _
                    legalStatus:="Pretrial", _
                    Courtroom:="Intake Conf.", _
                    DA:=DA.value, _
                    agency:="", _
                    startDate:=InConfDate.value, _
                    Re1:="", _
                    Re2:="", _
                    Re3:="", _
                    Notes:="Held at intake conference")

            Case "Release for Court"
                Call ReferClientTo( _
                    referralDate:=InConfDate.value, _
                    clientRow:=emptyRow, _
                    fromCR:="Intake Conf.", _
                    toCR:=InitialHearingLocation.value _
                    )
                If InitialHearingLocation.value = "5E" Then
                    Range(hFind("Courtroom of Origin", "Crossover") & emptyRow).value _
                            = Lookup("Courtroom_Name")("Intake Conf.")
                Else
                    Range(hFind("Courtroom of Origin", InitialHearingLocation.value) & emptyRow).value _
                            = Lookup("Courtroom_Name")("Intake Conf.")
                End If

                'add supervisions and conditions if assigned
                If Not Supv1.value = "None" Then
                    Call addSupervision( _
                        clientRow:=emptyRow, _
                        serviceType:=Supv1.value, _
                        legalStatus:="Pretrial", _
                        Courtroom:="Intake Conf.", _
                        CourtroomOfOrder:="Intake Conf.", _
                        DA:=DA.value, _
                        agency:=Supv1Pro.value, _
                        startDate:=InConfDate.value, _
                        NextCourtDate:=InitialHearingDate.value, _
                        Re1:=Supv1Re1.value, _
                        Re2:=Supv1Re2.value, _
                        Re3:=Supv1Re3.value, _
                        Notes:="Referred at intake conference")
                End If

                If Not Supv2.value = "None" Then
                    Call addSupervision( _
                        clientRow:=emptyRow, _
                        serviceType:=Supv2.value, _
                        legalStatus:="Pretrial", _
                        Courtroom:="Intake Conf.", _
                        CourtroomOfOrder:="Intake Conf.", _
                        DA:=DA.value, _
                        agency:=Supv2Pro.value, _
                        startDate:=InConfDate.value, _
                        NextCourtDate:=InitialHearingDate.value, _
                        Re1:=Supv2Re1.value, _
                        Re2:=Supv2Re2.value, _
                        Re3:=Supv2Re3.value, _
                        Notes:="Referred at intake conference")
                End If

                If Not Cond1.value = "None" Then
                    Call addCondition( _
                        clientRow:=emptyRow, _
                        condition:=Cond1.value, _
                        legalStatus:="Pretrial", _
                        Courtroom:="Intake Conf.", _
                        CourtroomOfOrder:="Intake Conf.", _
                        DA:=DA.value, _
                        agency:=Cond1Pro.value, _
                        startDate:=InConfDate.value, _
                        Re1:="N/A", _
                        Re2:="N/A", _
                        Re3:="N/A", _
                        Notes:="Referred at intake conference")
                End If

                If Not Cond2.value = "None" Then
                    Call addCondition( _
                        clientRow:=emptyRow, _
                        condition:=Cond2.value, _
                        legalStatus:="Pretrial", _
                        Courtroom:="Intake Conf.", _
                        CourtroomOfOrder:="Intake Conf.", _
                        DA:=DA.value, _
                        agency:=Cond2Pro.value, _
                        startDate:=InConfDate.value, _
                        Re1:="N/A", _
                        Re2:="N/A", _
                        Re3:="N/A", _
                        Notes:="Referred at intake conference")
                End If

                If Not Cond3.value = "None" Then
                    Call addCondition( _
                        clientRow:=emptyRow, _
                        condition:=Cond3.value, _
                        legalStatus:="Pretrial", _
                        Courtroom:="Intake Conf.", _
                        CourtroomOfOrder:="Intake Conf.", _
                        DA:=DA.value, _
                        agency:=Cond3Pro.value, _
                        startDate:=InConfDate.value, _
                        Re1:="N/A", _
                        Re2:="N/A", _
                        Re3:="N/A", _
                        Notes:="Referred at intake conference")
                End If

            Case "Release for Diversion"

        End Select
    Else
        Range(headerFind("Did Youth Have an Intake Conference?", tempHead) & emptyRow).value _
                = Lookup("Generic_NYNOU_Name")("Unknown")

        Select Case InitialHearingLocation.value
            Case "4G", "4E", "6F", "6H", "3E", "JTC", "5E", "WRAP", "Adult"
                Call ReferClientTo( _
                    referralDate:=PetitionBox.List(0, 0), _
                    clientRow:=emptyRow, _
                    toCR:=InitialHearingLocation.value _
                    )
        End Select
    End If







    'Range(headerFind("DA") & emptyRow).value = Lookup("DA_Last_Name_Name")(DA.value)
    Range(headerFind("General Notes from Intake") & emptyRow).value = GeneralNotes.value


    '''''''''''''''''''
    '''''DIVERSION'''''
    '''''''''''''''''''


    Dim diversionHead As String

    diversionHead = headerFind("DIVERSION")

    Range(headerFind("Referred to Diversion?", petitionHead) & emptyRow) _
            = Lookup("Generic_YNOU_Name")(DiversionProgram.value)
    Range(headerFind("Referred to Diversion?", diversionHead) & emptyRow) _
            = Lookup("Generic_YNOU_Name")(DiversionProgram.value)

    If DiversionProgram.value = "Yes" Then

        Range(headerFind("Which Diversion Program Used", petitionHead) & emptyRow) _
                = Lookup("Diversion_Program_Name")(NameOfProgram.value)
        Range(headerFind("Diversion Referral Date", petitionHead) & emptyRow) _
                = DiversionProgramReferralDate.value

        Range(headerFind("Referral Date", diversionHead) & emptyRow) _
                = DiversionProgramReferralDate.value
        Range(headerFind("Referral Source", diversionHead) & emptyRow) _
                = Lookup("Diversion_Referral_Source_Name")(ReferralSource.value)
        Range(headerFind("Age at Diversion Referral", diversionHead) & emptyRow) _
                = ageAtTime(DiversionProgramReferralDate.value, emptyRow)
        Range(headerFind("Diversion Program Ordered", diversionHead) & emptyRow) _
                = Lookup("Diversion_Program_Name")(NameOfProgram.value)

        If IsNumeric(YAPDistrict.value) Then
            Range(headerFind("YAP Panel District #", diversionHead) & emptyRow) _
                    = Lookup("Police_District_Name")(CInt(YAPDistrict.value))
        Else
            Range(headerFind("YAP Panel District #", diversionHead) & emptyRow) _
                    = Lookup("Police_District_Name")(YAPDistrict.value)
        End If


        Range(headerFind("Victim First Name", diversionHead) & emptyRow) _
                = VictimFirstName.value
        Range(headerFind("Victim Last Name", diversionHead) & emptyRow) _
                = VictimLastName.value

        Range(headerFind("Legal Status") & emptyRow).value _
                = Lookup("Legal_Status_Name")("Diversion")

        Range(headerFind("Did Youth Receive a Review Hearing?", diversionHead) & emptyRow) _
                = 2
        Range(headerFind("Did Youth Receive an Exit Hearing?", diversionHead) & emptyRow) _
                = 2
    End If

    If DiversionProgram.value = "No" Then
        Range(headerFind("Reason #1 Not Diverted", diversionHead) & emptyRow) _
                = Lookup("Diversion_Rejection_Reason_Name")(NoDiversionReason1.value)
        Range(headerFind("Reason #2 Not Diverted", diversionHead) & emptyRow) _
                = Lookup("Diversion_Rejection_Reason_Name")(NoDiversionReason2.value)
        Range(headerFind("Reason #3 Not Diverted", diversionHead) & emptyRow) _
                = Lookup("Diversion_Rejection_Reason_Name")(NoDiversionReason3.value)
    End If


    Call addNotes( _
        Courtroom:=InitialHearingLocation.value, _
        dateOf:=InConfDate.value, _
        userRow:=emptyRow, _
        Notes:=GeneralNotes, _
        DA:=DA.value _
    )




done:
    'Call SaveAs_Countdown
    Call Save_Countdown
    Call UnloadAll

    Worksheets("User Entry").Activate
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
    End With
    Exit Sub
err:

    Range("C" & emptyRow & ":" & headerFind("END") & emptyRow).value = restorer

    Stop 'press F8 twice to see the error point
    Resume
    MsgBox "Something went wrong. Database has been restored to state prior to submission. " _
      & vbNewLine & vbNewLine & "Message: " & vbNewLine & err.Description _
      & vbNewLine & vbNewLine & "Source: " & vbNewLine & err.Source

    Call UnloadAll
End Sub

Private Sub TestFillPetition_Click()
    With PetitionBox
        .ColumnCount = 7
        .ColumnWidths = "50;50;30;50;65;50;0"
        .AddItem "01/01/2019"
        .List(0, 1) = "13579"
        .List(0, 2) = "F"
        .List(0, 3) = "Assaults"
        .List(0, 4) = "18 - 2702"
        .List(0, 5) = "AGGRAVATED ASSAULT"
        .List(0, 6) = "No"
    End With
    
    CallInDate.value = "02/01/2019"
    Was_DRAI_Administered.value = "Yes"
    DRAI_Score.value = "4"
    DRAI_Rec.value = "Release"
    DRAI_Action.value = "Follow - Release"
    
    InConfDate.value = "02/01/2019"
    ConfOutcome.value = "Release for Court"
    
    NoDiversionReason1 = "Charge Ineligible"
    
    InitialHearingLocation.value = "3E"

    InitialHearingDate = "2/12/2019"

    DA.value = "Keller"

    GeneralNotes.value = "Gen Notes Test"

End Sub

Private Sub TestFillDiversion_Click()
    'Petitions

    With PetitionBox
        .ColumnCount = 7
        .ColumnWidths = "50;50;30;50;65;50;0"
        .AddItem "08/01/2018"
        .List(0, 1) = "13579"
        .List(0, 2) = "F"
        .List(0, 3) = "Assaults"
        .List(0, 4) = "18 - 2702"
        .List(0, 5) = "AGGRAVATED ASSAULT"
        .List(0, 6) = "No"
    End With

    CallInDate.value = "02/01/2019"
    Was_DRAI_Administered.value = "Yes"
    DRAI_Score.value = "4"
    DRAI_Rec.value = "Release"
    DRAI_Action.value = "Follow - Release"
    
    InConfDate.value = "02/01/2019"
    ConfOutcome.value = "Release for Diversion"
    
    DiversionProgram.value = "Yes"
    DiversionProgramReferralDate.value = "2/1/19"
    ReferralSource.value = "Pre-Petition DA"
    NameOfProgram.value = "YAP"
    YAPDistrict.value = 2
    
    GeneralNotes.value = "Gen Notes Test"

    InitialHearingDate.value = "02/01/2019"
    InitialHearingLocation.value = "Diversion"
    ListingType.value = "Diversion"

End Sub

Private Sub TestFillIntake_Click()
    With PetitionBox
        .ColumnCount = 7
        .ColumnWidths = "50;50;30;50;65;50;0"
        .AddItem "09/08/2018"
        .List(0, 1) = "13579"
        .List(0, 2) = "F"
        .List(0, 3) = "Assaults"
        .List(0, 4) = "18 - 2702"
        .List(0, 5) = "AGGRAVATED ASSAULT"
        .List(0, 6) = "No"
    End With

    InConfRecord.value = "Yes"
    InConfType.value = "DA"
    Was_DRAI_Administered = "Yes"
    DRAI_Score.value = 25
    DRAI_Rec.value = "Hold"
    DRAI_Action = "Follow - Hold"
    OverrideHoldRe1.value = "B/W"
    OverrideHoldRe2.value = "Drug Screens"
    OverrideHoldRe3.value = "N/A"
    ConfOutcome.value = "Release for Court"

    GeneralNotes.value = "Gen Notes Test"

    InitialHearingDate.value = "02/01/2019"
    InitialHearingLocation.value = "3E"

    DiversionProgram.value = "No"

End Sub




