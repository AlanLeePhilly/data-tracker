VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewClientForm 
   Caption         =   "NewClientUserForm"
   ClientHeight    =   10575
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   18075
   OleObjectBlob   =   "NewClientForm.frx":0000
End
Attribute VB_Name = "NewClientForm"
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


Private Sub SameDate_2_Click()
    CallInDate.value = ArrestDate.value
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



Private Sub Rearrest_Click()
    If isNotEmptyOrZero(Range(hFind("Arrest Date #" & Rearrest_Num.value, "REARRESTS", "AGGREGATES") & Rearrest_Row.value)) Then
        Call RearrestIntake(Rearrest_Row.value, Rearrest_Num.value)
    Else
        MsgBox "Arrest not found. Sorry!"
    End If

End Sub

Private Sub Reload_Click()
    Worksheets("Entry").Activate
    Call Generate_Dictionaries
    '''Demographics
    Dim emptyRow

    emptyRow = CLng(Reload_Row.value)
    FirstName.value = Range(headerFind("First Name") & emptyRow).value
    LastName.value = Range(headerFind("Last Name") & emptyRow).value
    DateOfBirth.value = Range(headerFind("DOB") & emptyRow).value
    Race.value = Lookup("Race_Num")(Range(headerFind("Race") & emptyRow).value)
    Sex.value = Lookup("Sex_Num")(Range(headerFind("Sex") & emptyRow).value)
    Latino.value = Lookup("Latino_Num")(Range(headerFind("Latino/Not Latino") & emptyRow).value)

    '''Community

    GuardianFirstName.value = Range(headerFind("Guardian First") & emptyRow).value
    GuardianLastName.value = Range(headerFind("Guardian Last") & emptyRow).value

    Address.value = Range(headerFind("Address") & emptyRow).value
    Zipcode.value = Range(headerFind("Zipcode") & emptyRow).value


    PhoneNumber.value = Range(headerFind("Phone #") & emptyRow).value
    School.value = Range(headerFind("School") & emptyRow).value
    Grade.value = Range(headerFind("Grade") & emptyRow).value

    '''Incident and Arrest
    petitionHead = headerFind("PETITION")

    IncidentDate.value = Range(headerFind("Incident Date") & emptyRow).value
    TimeOfIncident_H.value = getHour(Range(headerFind("Time of Incident") & emptyRow).value)
    TimeOfIncident_M.value = getMinute(Range(headerFind("Time of Incident") & emptyRow).value)
    TimeOfIncident_P.value = getPeriod(Range(headerFind("Time of Incident") & emptyRow).value)
    IncidentDistrict.value = Lookup("Police_District_Name")(Range(headerFind("Incident District") & emptyRow).value)
    IncidentAddress.value = Range(headerFind("Incident Address") & emptyRow).value
    IncidentZipcode.value = Range(headerFind("Incident Zipcode") & emptyRow).value

    ArrestDate.value = Range(headerFind("Arrest Date", petitionHead) & emptyRow).value

    TimeOfArrest_H.value = getHour(Range(headerFind("Time of Arrest") & emptyRow).value)
    TimeOfArrest_M.value = getMinute(Range(headerFind("Time of Arrest") & emptyRow).value)
    TimeOfArrest_P.value = getPeriod(Range(headerFind("Time of Arrest") & emptyRow).value)

    TimeReferredToDA_H.value = getHour(Range(headerFind("Time of Referral to DA") & emptyRow).value)
    TimeReferredToDA_M.value = getMinute(Range(headerFind("Time of Referral to DA") & emptyRow).value)
    TimeReferredToDA_P.value = getPeriod(Range(headerFind("Time of Referral to DA") & emptyRow).value)
    ArrestingDistrict.value = Range(headerFind("Arresting District", petitionHead) & emptyRow).value

    ActiveAtArrest.value = Lookup("Generic_YNOU_Num")(Range(headerFind("Active in System at Time of Arrest?") & emptyRow).value)
    NumOfPriorArrests.value = Lookup("Num_Prior_Arrests_Num")(Range(headerFind("# of Prior Arrests") & emptyRow).value)
    DCNum.value = Range(headerFind("DC #", petitionHead) & emptyRow).value
    PIDNum.value = Range(headerFind("PID #") & emptyRow).value
    SIDNum.value = Range(headerFind("SID #") & emptyRow).value

    Officer1.value = Range(headerFind("Officer #1") & emptyRow).value
    Officer2.value = Range(headerFind("Officer #2") & emptyRow).value
    Officer3.value = Range(headerFind("Officer #3") & emptyRow).value
    Officer4.value = Range(headerFind("Officer #4") & emptyRow).value
    Officer4.value = Range(headerFind("Officer #5") & emptyRow).value

    VictimFirstName = Range(headerFind("Victim First Name") & emptyRow)
    VictimLastName = Range(headerFind("Victim Last Name") & emptyRow)

    Dim i As Integer
    Dim j As Integer
    Dim sectionHead As String
    Dim bucketHead As String

    sectionHead = headerFind("PETITION")
    For i = 1 To 5
        If isNotEmptyOrZero(Range(headerFind("Petition #" & i, sectionHead) & emptyRow)) Then
            bucketHead = headerFind("Petition #" & i, sectionHead)

            With NewClientForm.PetitionBox
                .ColumnCount = 7
                .ColumnWidths = "50;50;30;50;65;50;0"
                ' 0 Date Filed
                ' 1 Petition Number
                ' 2 Charge Grade
                ' 3 Charge Group
                ' 4 Charge Code
                ' 5 Charge Name
                ' 6 Was Petition from other county?
                .AddItem Range(headerFind("Date Filed", bucketHead) & emptyRow).value
                .List(.ListCount - 1, 0) = Range(headerFind("Date Filed", bucketHead) & emptyRow).value
                .List(.ListCount - 1, 1) = Range(bucketHead & emptyRow).value
                .List(.ListCount - 1, 2) = Lookup("Charge_Grade_Specific_Num")(Range(headerFind("Charge Grade (specific) #1", bucketHead) & emptyRow).value)
                .List(.ListCount - 1, 3) = Lookup("Charge_Num")(Range(headerFind("Charge Category #1", bucketHead) & emptyRow).value)
                .List(.ListCount - 1, 4) = Range(headerFind("Lead Charge Code", bucketHead) & emptyRow).value
                .List(.ListCount - 1, 5) = Range(headerFind("Lead Charge Name", bucketHead) & emptyRow).value
                .List(.ListCount - 1, 6) = Lookup("Generic_YNOU_Num")(Range(headerFind("Was Petition Transferred from Other County?", bucketHead) & emptyRow).value)
            End With

            For j = 2 To 5
                If isNotEmptyOrZero(Range(headerFind("Charge Code #" & j, bucketHead) & emptyRow)) Then
                    With NewClientForm.ChargeBox
                        .ColumnCount = 5
                        .ColumnWidths = "50;50;30;50;65;"
                        ' 0 Petition Number
                        ' 1 Charge Grade
                        ' 2 Charge Group (specific)
                        ' 3 Charge Code
                        ' 4 Charge Name
                        .AddItem Range(bucketHead & emptyRow).value
                        .List(.ListCount - 1, 0) = Range(bucketHead & emptyRow).value
                        .List(.ListCount - 1, 1) = Lookup("Charge_Grade_Specific_Num")(Range(headerFind("Charge Grade (specific) #" & j, bucketHead) & emptyRow).value)
                        .List(.ListCount - 1, 2) = Lookup("Charge_Num")(Range(headerFind("Charge Category #" & j, bucketHead) & emptyRow).value)
                        .List(.ListCount - 1, 3) = Range(headerFind("Charge Code #" & j, bucketHead) & emptyRow).value
                        .List(.ListCount - 1, 4) = Range(headerFind("Charge Name #" & j, bucketHead) & emptyRow).value
                    End With
                End If
            Next j
        End If
    Next i

    tempHead = headerFind("INTAKE CONFERENCE", petitionHead)
    'thing = Lookup("Generic_NYNOU_Num")(Range(headerFind("Did Youth Have an Intake Conference?", tempHead) & emptyRow).value)
    'IntakeConference.value = Lookup("Generic_NYNOU_Num")(Range(headerFind("Did Youth Have an Intake Conference?", tempHead) & emptyRow).value)
    InConfDate.value = Range(headerFind("Date of Intake Conference", tempHead) & emptyRow).value
    InConfType.value = Lookup("Intake_Conference_Type_Num")(Range(headerFind("Intake Conference Type", tempHead) & emptyRow).value)

    tempHead = headerFind("CALL-IN")
    CallInDate.value = Range(headerFind("Date of Call-In", tempHead) & emptyRow).value
    Was_DRAI_Administered = Lookup("Generic_NYNOU_Num")(Range(headerFind("Was DRAI Administered?", tempHead) & emptyRow).value)
    DRAI_Score.value = Range(headerFind("DRAI Score", tempHead) & emptyRow).value
    DRAI_Rec.value = Lookup("DRAI_Recommendation_Num")(Range(headerFind("DRAI Recommendation", tempHead) & emptyRow).value)
    DRAI_Action = Lookup("DRAI_Action_Num")(Range(headerFind("DRAI Action", tempHead) & emptyRow).value)
    OverrideHoldRe1.value = Lookup("DRAI_Override_Reason_Num")(Range(headerFind("Reason #1 for Override Hold", tempHead) & emptyRow).value)
    OverrideHoldRe2.value = Lookup("DRAI_Override_Reason_Num")(Range(headerFind("Reason #2 for Override Hold", tempHead) & emptyRow).value)
    OverrideHoldRe3.value = Lookup("DRAI_Override_Reason_Num")(Range(headerFind("Reason #3 for Override Hold", tempHead) & emptyRow).value)

    ConfOutcome.value = Lookup("Intake_Conference_Outcome_Num")(Range(headerFind("Intake Conference Outcome", tempHead) & emptyRow).value)

    Supv1 = Lookup("Supervision_Program_Num")(Range(headerFind("Supervision Ordered #1", tempHead) & emptyRow).value)
    Supv1Pro = Lookup("Community_Based_Supervision_Provider_Num")(Range(headerFind("Community-Based Agency #1", tempHead) & emptyRow).value)
    Supv2 = Lookup("Supervision_Program_Num")(Range(headerFind("Supervision Ordered #2", tempHead) & emptyRow).value)
    Supv2Pro = Lookup("Community_Based_Supervision_Provider_Num")(Range(headerFind("Community-Based Agency #2", tempHead) & emptyRow).value)
    DetentionFacility = Lookup("Detention_Facility_Num")(Range(hFind("Detention Facility", "DETENTION") & emptyRow).value)

    Cond1 = Lookup("Condition_Num")(Range(headerFind("Other Condition #1", tempHead) & emptyRow).value)
    Cond1Pro = Lookup("Condition_Provider_Num")(Range(headerFind("Other Condition #1 Provider", tempHead) & emptyRow).value)
    Cond2 = Lookup("Condition_Num")(Range(headerFind("Other Condition #2", tempHead) & emptyRow).value)
    Cond2Pro = Lookup("Condition_Provider_Num")(Range(headerFind("Other Condition #2 Provider", tempHead) & emptyRow).value)
    Cond3 = Lookup("Condition_Num")(Range(headerFind("Other Condition #3", tempHead) & emptyRow).value)
    Cond3Pro = Lookup("Condition_Provider_Num")(Range(headerFind("Other Condition #3 Provider", tempHead) & emptyRow).value)


    DiversionProgram.value = Lookup("Generic_YNOU_Num")(Range(headerFind("Referred to Diversion?", petitionHead) & emptyRow).value)
    DiversionProgramReferralDate.value = Range(headerFind("Diversion Referral Date", petitionHead) & emptyRow).value
    ReferralSource.value = Lookup("Diversion_Referral_Source_Num")(Range(headerFind("Referral Source", diversionHead) & emptyRow).value)
    NameOfProgram.value = Lookup("Diversion_Program_Num")(Range(headerFind("Diversion Program Ordered", diversionHead) & emptyRow).value)
    YAPDistrict.value = Lookup("Police_District_Num")(Range(headerFind("YAP Panel District #", diversionHead) & emptyRow).value)


    InitialHearingDate.value = Range(headerFind("Initial Hearing Date") & emptyRow).value
    InitialHearingLocation.value = Lookup("Courtroom_Num")(Range(headerFind("Initial Hearing Location") & emptyRow).value)
    ListingType.value = Lookup("Listing_Type_Num")(Range(headerFind("Listing Type") & emptyRow).value)
    DA.value = Lookup("DA_Last_Name_Num")(Range(headerFind("DA") & emptyRow).value)

    GeneralNotes.value = Range(headerFind("General Notes from Intake") & emptyRow).value



End Sub


Private Sub Supv2_Change()
    If Supv2.value = "None " Then
        MsgBox "wow wtf"
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
        Modal_NewClient_Add_Petition.headline.Caption = "New Client"
        Modal_NewClient_Add_Petition.Show
    Else
        MsgBox "Maximum of five petitions for new client"
    End If
End Sub


'Private Sub ConfOutcome_Change()
'    Select Case ConfOutcome.value
'        Case "Hold for Detention"
'            DetentionFacilityLabel.Enabled = True
'            DetentionFacility.Enabled = True
'            InitialHearingLocation = "PJJSC"
'        Case "Roll to Detention Hearing"
'            DetentionFacilityLabel.Enabled = False
'            DetentionFacility.Enabled = False
'           InitialHearingLocation = "PJJSC"
'       Case "Release for Diversion"
'           DetentionFacilityLabel.Enabled = False
'           DetentionFacility.Enabled = False
'           DetentionFacility.value = ""
'       Case Else
'           DetentionFacilityLabel.Enabled = False
'           DetentionFacility.Enabled = False
'          DetentionFacility.value = ""
'  End Select
'End Sub

Private Sub DateOfBirth_Enter()
    DateOfBirth.value = CalendarForm.GetDate(RangeOfYears:=30)
End Sub
Private Sub DateOfBirth_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Set ctl = Me.DateOfBirth
    Call DateValidation(ctl, Cancel)
End Sub

'for a date texbox called "ArrestDate"
Private Sub ArrestDate_Enter() 'when a user "Enters" (clicks) the text box...
    'this boxe's value is defined as a result of the picker being called and completed
    ArrestDate.value = CalendarForm.GetDate(RangeOfYears:=5) '(set range from today)
End Sub

Private Sub ArrestDate_Exit(ByVal Cancel As MSForms.ReturnBoolean) 'when a user "Exits" (clicks outside of) the text box...
    Set ctl = Me.ArrestDate 'set the text box to a variable (the whole text box, not its value)
    Call DateValidation(ctl, Cancel) 'send the text box to a custom date validation function
End Sub


Private Sub IncidentDate_Enter()
    IncidentDate.value = CalendarForm.GetDate(RangeOfYears:=5)
End Sub
Private Sub IncidentDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Set ctl = Me.IncidentDate
    Call DateValidation(ctl, Cancel)
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

Private Sub EditPetition_Click()
    Dim petitionNum As String
    Dim i As Integer, j As Integer
    Dim listIndex As Integer
    Dim pBox, cBox

    Set pBox = PetitionBox

    If pBox.listIndex = -1 Then
        Exit Sub
    End If

    petitionNum = pBox.List(pBox.listIndex, 1)
    MsgBox "Editing Petition #" & petitionNum

    With Modal_NewClient_Add_Petition
        .headline.Caption = "Edit Petition"
        .DateFiled.value = pBox.List(pBox.listIndex, 0)
        .Num.value = pBox.List(pBox.listIndex, 1)
        .ChargeGrade1.value = pBox.List(pBox.listIndex, 2)
        .ChargeGroup1.value = pBox.List(pBox.listIndex, 3)
        With .ChargeList1
            .ColumnCount = 2
            .ColumnWidths = "85;400;"
            .AddItem pBox.List(pBox.listIndex, 4)
            .List(.ListCount - 1, 1) = pBox.List(pBox.listIndex, 5)
            .listIndex = 0
        End With

        .isTransferred.value = pBox.List(pBox.listIndex, 6)

        listIndex = ChargeBox.ListCount - 1
        j = 2
        For i = 0 To listIndex
            If ChargeBox.List(i, 0) = petitionNum Then
                Select Case j
                    Case 2
                        Call .LoadBox(.ChargeList2, ChargeBox.List(i, 3), ChargeBox.List(i, 4))
                        .ChargeGrade2 = ChargeBox.List(i, 1)
                        .ChargeGroup2 = ChargeBox.List(i, 2)
                        j = j + 1
                    Case 3
                        Call .LoadBox(.ChargeList3, ChargeBox.List(i, 3), ChargeBox.List(i, 4))
                        .ChargeGrade3 = ChargeBox.List(i, 1)
                        .ChargeGroup3 = ChargeBox.List(i, 2)
                        j = j + 1
                    Case 4
                        Call .LoadBox(.ChargeList4, ChargeBox.List(i, 3), ChargeBox.List(i, 4))
                        .ChargeGrade4 = ChargeBox.List(i, 1)
                        .ChargeGroup4 = ChargeBox.List(i, 2)
                        j = j + 1
                    Case 5
                        Call .LoadBox(.ChargeList5, ChargeBox.List(i, 3), ChargeBox.List(i, 4))
                        .ChargeGrade5 = ChargeBox.List(i, 1)
                        .ChargeGroup5 = ChargeBox.List(i, 2)
                        j = j + 1
                End Select
            End If
        Next i
        Call DeletePetition_Click
        .Show
    End With
End Sub


Private Sub InitialHearingDate_Enter()
    InitialHearingDate.value = CalendarForm.GetDate(RangeOfYears:=5)
End Sub
Private Sub InitialHearingDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Set ctl = Me.InitialHearingDate
    Call DateValidation(ctl, Cancel)
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
    NewClientUserForm.Hide
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

Private Sub SameDate_Click()
    ArrestDate.value = IncidentDate.value
End Sub
Private Sub SameTime_Click()
    TimeOfArrest_H.value = TimeOfIncident_H.value
    TimeOfArrest_M.value = TimeOfIncident_M.value
    TimeOfArrest_P.value = TimeOfIncident_P.value
End Sub
Private Sub SameDistrict_Click()
    ArrestingDistrict.value = IncidentDistrict.value
End Sub

'Here is the updated FindLatLon function that should be used in BigCahuna, complete with message boxes for incorrect input and error catching. After the function is the updated code to be used in the Submit_Click Sub (or what we should use in ClientEdit or ClientUpdate).


Private Function FindLatLon(Address As String, Zipcode As String)
    Dim hReq As Object
    Dim Json As Object
    Dim try As Object
    Dim strUrl As String
    Dim addressStr As String
    Dim cityStr As String

    On Error GoTo LatLonErr

    cityStr = ""

    'Check to see if address is "Homeless" or if zipcode is "19100"
    If StrComp(UCase(Address), "HOMELESS") = 0 Then
        MsgBox ("ALERT: Address of 'homeless' entered was not mappable; no latitude or longitude coordinates added")
        Dim responseArr() As Variant
        responseArr = Array("", "", "", Zipcode)
        FindLatLon = responseArr
        Exit Function
    End If

    If StrComp(Zipcode, "19100") = 0 Then
        MsgBox ("ALERT: Address entered with zipcode '19100'; city 'Philadelphia' used with no zipcode to attempt mapping instead")
        cityStr = "Philadelphia"
        Zipcode = ""
    End If

    'Probably want try-catch here
    strUrl = "https://nominatim.openstreetmap.org/search?format=json&addressdetails=1&limit=1&q=" & WorksheetFunction.EncodeURL(Address) & "%20" & WorksheetFunction.EncodeURL(cityStr) & "%20" & WorksheetFunction.EncodeURL(Zipcode)

    Set hReq = CreateObject("MSXML2.XMLHTTP")
    With hReq
        .Open "GET", strUrl, False
        .Send
    End With

    addressStr = "[{" & Mid(hReq.ResponseText, 107, Len(hReq.ResponseText))

    Set Json = JsonConverter.ParseJson(addressStr)

    Dim a(1 To 3) As Double
    a(1) = Json(1)("lat")
    a(2) = Json(1)("lon")
    a(3) = Json(1)("address")("postcode")
    FindLatLon = a
    Exit Function

LatLonErr:
    MsgBox ("ALERT: Error occurred in finding location coordinates: " & err.Description & "; Setting coordinates to null. Please check address and zipcode and edit if location coordinates desired.")
    Dim errorArr() As Variant
    errorArr = Array("", "", "", Zipcode)
    FindLatLon = errorArr
    Exit Function

End Function


Private Sub Submit_Click()
    Dim restorer As Variant

    Call Generate_Dictionaries
    'define variable Long(a big integer) named emptyRow
    Dim emptyRow As Long

    'activate the spreadsheet as default selector
    Worksheets("Entry").Activate

    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With

    '''''''''''''
    'Validations'
    '''''''''''''

    'confirm that client first name is present in the form
    If FirstName.value = "" Then
        MsgBox "First Name Required"
        Exit Sub
    End If

    'confirm that client last name is present in the form
    If LastName.value = "" Then
        MsgBox "Last Name Required"
        Exit Sub
    End If


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

    'find empty row by finding first 'first name' value from bottom
    emptyRow = Range("C" & Rows.count).End(xlUp).row + 1
    restorer = Range("C" & emptyRow & ":" & hFind("END") & emptyRow).value

    If Not Reload_Row = "" Then
        If MsgBox("Warning, you are about to overwrite row " & Reload_Row.value, vbOKCancel) = vbCancel Then
            Exit Sub
        End If

        emptyRow = Reload_Row
        Range("C" & emptyRow & ":" & hFind("END") & emptyRow).ClearContents
    End If


    On Error GoTo err

    ''''''''''''''
    'DEMOGRAPHICS'
    ''''''''''''''

    Range(headerFind("First Name") & emptyRow).value = FirstName.value
    Range(headerFind("Last Name") & emptyRow).value = LastName.value
    Range(headerFind("Next Court Date") & emptyRow).value = InitialHearingDate.value
    Range(headerFind("Listing Type") & emptyRow).value = Lookup("Listing_Type_Name")(ListingType.value)
    Range(headerFind("Previous Court Dates") & emptyRow).value = InConfDate.value
    Range(headerFind("Arrest Date (current petition)") & emptyRow).value = ArrestDate.value
    Range(headerFind("Initial Hearing Date") & emptyRow).value = InitialHearingDate.value
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
            = ageAtTime(InitialHearingDate, emptyRow)
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
    petitionHead = hFind("PETITION")

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

    Range(headerFind("General Notes from Intake") & emptyRow).value = GeneralNotes.value

    Dim Num As Long
    Dim i As Integer
    Dim j As Integer

    For Num = 1 To PetitionBox.ListCount
        tempHead = headerFind("Petition #" & Num, petitionHead)

        Range(headerFind("Petition Filed?", tempHead) & emptyRow).value _
                = Lookup("Generic_YNOU_Name")("Yes")
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
            Call startLegalStatus( _
                clientRow:=emptyRow, _
                statusType:="Pretrial", _
                Courtroom:="Intake Conf.", _
                DA:=DA.value, _
                startDate:=PetitionBox.List(0, 0))
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

    '''''''''''''''''
    '''DETENTION'''''
    '''''''''''''''''

    Range(hFind("Did Youth Have Initial Detention Hearing?", "DETENTION") & emptyRow).value = 2 '"No"

    'TOTAL OUTCOMES
    Dim arrestHead As String
    arrestHead = headerFind("ARREST GRAPH")


    Dim chargeKey As String

    Select Case Lookup("Charge_Name")(PetitionBox.List(0, 3))
        Case 1, 2, 3, 4, 6, 14
            chargeKey = "Violent"
        Case 5
            chargeKey = "Property"
        Case 8
            chargeKey = "Weapons"
        Case 9
            chargeKey = "Drugs"
        Case 10
            chargeKey = "Sexual"
        Case 15
            chargeKey = "Public Order"
        Case Else
            chargeKey = "Other"
    End Select

    tempHead = headerFind(chargeKey, arrestHead)

    Range(headerFind("Time of Incident", tempHead) & emptyRow).value _
            = TimeValue(TimeOfIncident_H.value & ":" & TimeOfIncident_M.value & " " & TimeOfIncident_P.value)

    Range(headerFind("Day of Incident", tempHead) & emptyRow).value _
            = 2 * Weekday(IncidentDate.value, vbMonday) - 1 + Lookup("Day_Adjustment_Name")(chargeKey)



    Call addNotes( _
        Courtroom:=InitialHearingLocation.value, _
        dateOf:=InConfDate.value, _
        userRow:=emptyRow, _
        Notes:=GeneralNotes, _
        DA:=DA.value _
    )

    'ZERO FILL

    Dim counter As Long


    'For counter = (alphaToNum(headerFind("PETITION")) + 1) _
    '    To (alphaToNum(headerFind("DRAI")) - 1)
    '
    '    If Cells(emptyRow, counter).value = "" Then
    '        Cells(emptyRow, counter).value = 0
    '    End If
    'Next counter

    'For counter = (alphaToNum(headerFind("DEMOGRAPHICS")) + 1) _
    '    To (alphaToNum(headerFind("PETITION")) - 1)
    '
    '    If Cells(emptyRow, counter).value = "" Then
    '        Cells(emptyRow, counter).value = 0
    '    End If
    'Next counter
    Call aggFlag(emptyRow)
    Call courtsFlag(emptyRow)
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
    '''Demographics

    FirstName.value = "Adam"
    LastName.value = "PetSerlin"
    DateOfBirth.value = "3/7/2002"

    Race.value = "White"
    Sex.value = "Male"
    Latino = "Not Latino"


    '''Community

    GuardianFirstName.value = "Mother Serlin"
    GuardianLastName.value = "Serlin"

    Address.value = "817 N. 2nd St."
    Zipcode.value = "19123"


    PhoneNumber.value = "123-456-7890"
    School.value = "Franklin High School"
    Grade.value = "12"

    '''Incident and Arrest
    IncidentDate.value = "02/01/2019"
    TimeOfIncident_H.value = "04"
    TimeOfIncident_M.value = "00"
    TimeOfIncident_P.value = "PM"
    IncidentDistrict.value = "25"

    ArrestDate.value = "01/29/2019"
    TimeOfArrest_H.value = "01"
    TimeOfArrest_M.value = "15"
    TimeOfArrest_P.value = "PM"

    DCNum.value = "12345"
    PIDNum.value = "5467"
    SIDNum.value = "87980"

    VictimFirstName = "VicFirstTest"
    VictimLastName = "VicLastTest"

    ArrestingDistrict.value = "25"

    NumOfPriorArrests.value = 0
    ActiveAtArrest.value = "No"
    InitialHearingDate.value = "01/01/2019"
    InitialHearingLocation.value = "3E"

    Officer1.value = "AO1Test"
    Officer2.value = "AO2Test"
    Officer3.value = "AO3Test"
    Officer4.value = "AO4Test"

    TimeReferredToDA_H.value = "08"
    TimeReferredToDA_M.value = "55"
    TimeReferredToDA_P.value = "AM"
    DiversionProgram.value = "No"

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
    FirstName.value = "David"
    LastName.value = "Sterling"
    DateOfBirth.value = "3/7/2002"

    Race.value = "White"
    Sex.value = "Male"
    Latino = "Not Latino"

    GuardianFirstName.value = "Mother Serlin"
    GuardianLastName.value = "Serlin"

    Address.value = "817 N. 2nd St."
    Zipcode.value = "19123"


    PhoneNumber.value = "123-456-7890"
    School.value = "Franklin High School"
    Grade.value = "12"


    IncidentDate.value = "02/01/2019"
    TimeOfIncident_H.value = "04"
    TimeOfIncident_M.value = "00"
    TimeOfIncident_P.value = "PM"
    IncidentDistrict.value = "22"
    IncidentAddress.value = "1501 Market St."
    IncidentZipcode.value = "19102"

    ArrestDate.value = "02/01/2019"
    TimeOfArrest_H.value = "02"
    TimeOfArrest_M.value = "30"
    TimeOfArrest_P.value = "PM"
    TimeReferredToDA_H.value = "10"
    TimeReferredToDA_M.value = "00"
    TimeReferredToDA_P.value = "AM"
    ArrestingDistrict.value = "22"

    ActiveAtArrest.value = "No"
    NumOfPriorArrests.value = 0
    DCNum.value = "12345"
    PIDNum.value = "5467"
    SIDNum.value = "87980"

    Officer1.value = "AO1Test"
    Officer2.value = "AO2Test"
    Officer3.value = "AO3Test"
    Officer4.value = "AO4Test"
    Officer4.value = "AO5Test"

    VictimFirstName = "VicFirstTest"
    VictimLastName = "VicLastTest"

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
    '''Demographics

    FirstName.value = "Chad"
    LastName.value = "Merlin"
    DateOfBirth.value = "3/22/2002"

    Race.value = "White"
    Sex.value = "Male"
    Latino = "Not Latino"


    '''Community

    GuardianFirstName.value = "Jane"
    GuardianLastName.value = "Merlin"

    Address.value = "716 S. 12th St."
    Zipcode.value = "19148"


    PhoneNumber.value = "142-534-6346"
    School.value = "Wells High School"
    Grade.value = "10"

    IncidentDate.value = "02/23/2019"
    TimeOfIncident_H.value = "06"
    TimeOfIncident_M.value = "15"
    TimeOfIncident_P.value = "PM"
    IncidentDistrict.value = "19"
    IncidentAddress.value = "1000 Sansom St."
    IncidentZipcode.value = "19031"

    ArrestDate.value = "02/24/2019"
    TimeOfArrest_H.value = "02"
    TimeOfArrest_M.value = "30"
    TimeOfArrest_P.value = "PM"
    TimeReferredToDA_H.value = "10"
    TimeReferredToDA_M.value = "00"
    TimeReferredToDA_P.value = "AM"
    ArrestingDistrict.value = "22"

    ActiveAtArrest.value = "No"
    NumOfPriorArrests.value = 0
    DCNum.value = "12345"
    PIDNum.value = "5467"
    SIDNum.value = "87980"

    Officer1.value = "Off1"
    Officer2.value = "Off2"
    Officer3.value = "Off3"
    Officer4.value = "Off4"
    Officer4.value = "Off5"

    VictimFirstName = "VicFirst"
    VictimLastName = "VicLast"

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


