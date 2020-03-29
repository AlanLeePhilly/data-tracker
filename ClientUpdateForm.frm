VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ClientUpdateForm 
   Caption         =   "ClientUpdateForm"
   ClientHeight    =   10575
   ClientLeft      =   45
   ClientTop       =   -75
   ClientWidth     =   15975
   OleObjectBlob   =   "ClientUpdateForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ClientUpdateForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public dataStore As Collection
Option Explicit





Private Sub Adult_NextCourtDate_Enter()
    Adult_NextCourtDate.value = CalendarForm.GetDate(RangeOfYears:=5)
End Sub
Private Sub Adult_NextCourtDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Set ctl = ClientUpdateForm.Adult_NextCourtDate

    Call DateValidation(ctl, Cancel)
End Sub

Private Sub JTC_NCD_NA_Click()
    NextCourtDate.value = "N/A"
End Sub



Private Sub LogPayment_Click()
    If DA.value = "" Then
        MsgBox "DA Required"
        Exit Sub
    End If
    
    Log_Payment.Show
    
End Sub



Private Sub RearrestIntake_Click()
    Modal_Rearrest_Intake.Show
    
End Sub

Private Sub Standard_FTA_Yes_Click()
    If Worksheets("Entry").Range(hFind("Active B/W?") & updateRow).value = Lookup("Generic_YNOU_Name")("Yes") Then
        Standard_FTA_Yes.BackColor = selectedColor
        Standard_FTA_No.BackColor = unselectedColor
    Else
        Modal_FTA.Show
    End If
End Sub

Private Sub Standard_FTA_No_Click()
    Standard_FTA_Yes.BackColor = unselectedColor
    Standard_FTA_No.BackColor = selectedColor
End Sub

Private Sub DRev_Cancel_Click()
    Unload Me
End Sub

Private Sub DRev_DetentionDecision_Change()
    Select Case DRev_DetentionDecision.value
        Case "Held"
            If Range(hFind("Did Youth Have Initial Detention Hearing?", "DETENTION") & updateRow).value = Lookup("Generic_YN_Name")("Yes") Then
                MultiPage2.value = 3
            Else
                MultiPage2.value = 2
            End If
        Case "Released"
            MultiPage2.value = 1
            DRev_Facility.value = "N/A"
            DRev_Facility.Enabled = False
        Case "Remain as Commit"
            If Not Range(hFind("Did Youth Have Initial Detention Hearing?", "DETENTION") & updateRow).value _
                = Lookup("Generic_YN_Name")("Yes") Then
                MsgBox "It looks like this is the client's initial detention hearing. They cannot 'Remain as Commit', but they can be 'Held'"
                DRev_DetentionDecision.value = "Held"
            End If
        Case Else
            MultiPage2.value = 0
    End Select
End Sub


Private Sub DRev_HoldAndTransfer_Change()
    If DRev_HoldAndTransfer.value = "Yes" Then
        DRev_NextHearingLocation.Enabled = True
        DRev_NextHearingLocationLabel.Enabled = True
    Else
        DRev_NextHearingLocation.Enabled = False
        DRev_NextHearingLocationLabel.Enabled = False
        DRev_NextHearingLocation.value = "N/A"
    End If
End Sub

Private Sub DRevSup1_Change()
    If isResidential(DRevSup1) Then
        DRevSup1_Agency.RowSource = "Residential_Supervision_Provider"
    Else
        DRevSup1_Agency.RowSource = "Community_Based_Supervision_Provider"
    End If

    DRevSup1_Agency.value = "None"
End Sub

Private Sub DRevSup2_Change()
    If isResidential(DRevSup1) Then
        DRevSup2_Agency.RowSource = "Residential_Supervision_Provider"
    Else
        DRevSup2_Agency.RowSource = "Community_Based_Supervision_Provider"
    End If

    DRevSup2_Agency.value = "None"
End Sub



Private Sub RearrestButton_Click()
    Load Modal_New_Arrest
    Modal_New_Arrest.Fetch_First_Name = Range(hFind("First Name") & updateRow).value
    Modal_New_Arrest.Fetch_Last_Name = Range(hFind("Last Name") & updateRow).value
    Modal_New_Arrest.Active_Row = updateRow

    Modal_New_Arrest.Show
End Sub

Private Sub Standard_Court_Transfer_Click()
    Modal_Standard_Court_Transfer.Show
End Sub

Private Sub Standard_Lift_BW_Click()
    Modal_Lift_BW.Show
End Sub

Private Sub Standard_NCD_NA_Click()
    Standard_NextCourtDate.value = "N/A"
End Sub
Private Sub Adult_NCD_NA_Click()
    Adult_NextCourtDate.value = "N/A"
End Sub

Private Sub Standard_NextCourtDate_Enter()
    Standard_NextCourtDate.value = CalendarForm.GetDate(RangeOfYears:=5)
End Sub
Private Sub Standard_NextCourtDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Set ctl = Me.Standard_NextCourtDate
    Call DateValidation(ctl, Cancel)
End Sub
Private Sub PJJSC_NextCourtDate_Enter()
    PJJSC_NextCourtDate.value = CalendarForm.GetDate(RangeOfYears:=5)
End Sub
Private Sub PJJSC_NextCourtDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Set ctl = Me.PJJSC_NextCourtDate
    Call DateValidation(ctl, Cancel)
End Sub



''''''''''''''''
'INITIALIZATION'
''''''''''''''''

Sub UserForm_Initialize()
    Call Generate_Dictionaries
    JTC_Stepdown_Label.Visible = False
    JTC_Accept_Reject_Date_Label.Visible = False
    JTC_Referred_To_Label.Visible = False
    JTC_Accept.Visible = False
    JTC_Reject.Visible = False
    JTC_Expungement.Visible = False
    MultiPage1.value = 0
    MultiPage2.value = 0
    Me.ScrollTop = 0
End Sub
'''''''''''''
'VALIDATIONS'
'''''''''''''

Private Sub DateOfHearing_Enter()
    DateOfHearing.value = CalendarForm.GetDate(RangeOfYears:=5)
End Sub


Sub DateOfHearing_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Set ctl = Me.DateOfHearing

    Call DateValidation(ctl, Cancel)
End Sub

Private Sub NextCourtDate_Enter()

    NextCourtDate.value = CalendarForm.GetDate(RangeOfYears:=5)

End Sub

Sub NextCourtDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Set ctl = Me.NextCourtDate

    Call DateValidation(ctl, Cancel)
End Sub



''''''''''''''
'YOUTH_SEARCH'
''''''''''''''

Sub Yesterday_Click()
    DateOfHearing.value = DateAdd("d", -1, Date)
End Sub

Sub Today_Click()
    DateOfHearing.value = Date
End Sub



Sub SearchButton_Click()
    On Error Resume Next

    'define variable Long(a big integer) named updateRow
    Dim lastRow As Long
    Dim Query As String
    Dim lookRow As Long
    Dim lookCell As String
    'activate the spreadsheet as default selector
    Worksheets("Entry").Activate

    'define variable of search query in UPPERCASE named 'query'
    Query = UCase(SearchTextBox.value)

    SearchResultsBox.Clear

    lastRow = Range("C" & Rows.count).End(xlUp).row

    For lookRow = 3 To lastRow

        lookCell = UCase(Range(headerFind(Search_Type.value) & lookRow))

        If InStr(1, lookCell, Query) > 0 Then
            With SearchResultsBox
                .ColumnCount = 9
                .ColumnWidths = "30;70;80;80;70;70;70;70;70"
                .AddItem lookRow
                .List(SearchResultsBox.ListCount - 1, 1) = Range(headerFind("First Name") & lookRow)
                .List(SearchResultsBox.ListCount - 1, 2) = Range(headerFind("Last Name") & lookRow)
                .List(SearchResultsBox.ListCount - 1, 3) = Range(headerFind("DOB") & lookRow)
                .List(SearchResultsBox.ListCount - 1, 4) = Range(headerFind("Arrest Date") & lookRow)
                .List(SearchResultsBox.ListCount - 1, 5) = Range(headerFind("Petition #1") & lookRow)
                .List(SearchResultsBox.ListCount - 1, 6) = Lookup("Courtroom_Num")(Range(headerFind("Active Courtroom") & lookRow).value)
                .List(SearchResultsBox.ListCount - 1, 7) = Lookup("Legal_Status_Num")(Range(headerFind("Legal Status") & lookRow).value)
                .List(SearchResultsBox.ListCount - 1, 8) = Lookup("Supervision_Program_Num")(Range(headerFind("Active Supervision") & lookRow).value)
            End With
        End If
    Next lookRow
End Sub

Sub SearchResultsBox_Click()
    updateRow = SearchResultsBox.value
    Modal_New_Arrest.Active_Row = updateRow
    Courtroom.value = SearchResultsBox.List(SearchResultsBox.listIndex, 6)
End Sub

''''''''''''''''''''''
''''''DATA_FETCH''''''
''''''''''''''''''''''

Sub Lookup_Button_Click()
    'VaLIDATION: must enter hearing date before lookup
    If DateOfHearing.value = "" Then
        MsgBox "Please enter hearing date"
        Exit Sub
    End If

    'VALIDATION: must enter courtroom before lookup
    If Courtroom.value = "" Then
        MsgBox "Please enter courtroom"
        Exit Sub
    End If
    Worksheets("Entry").Activate
    'VALIDATION: if courtroom is PJJSC, must provide whether Initial or Review hearing
    On Error GoTo err

    'Select page to show depending on courtroom selected and fetch relevant data
    Select Case Courtroom.value
        Case "N/A"
            MultiPage1.value = 0
        Case "PJJSC"
            If Range(hFind("Did Youth Have Initial Detention Hearing?", "DETENTION") & updateRow).value = Lookup("Generic_YN_Name")("Yes") Then
                DetentionHeader.Caption = "DETENTION REVIEW"
            Else
                DetentionHeader.Caption = "INITIAL DETENTION HEARING"
            End If
            MultiPage1.value = 1
        Case "4G", "4E", "6F", "6H", "3E", "5E", "WRAP"
            MultiPage1.value = 4
            Call Standard_Fetch
        Case "JTC"
            MultiPage1.value = 2
            Call JTC_Fetch
        Case "Adult"
            MultiPage1.value = 3
            Call Adult_Fetch
        Case "Intake Conf."
            MultiPage1.value = 5
        Case Else
            MsgBox "Please select a valid courtroom to continue!"
            Exit Sub
    End Select
    Worksheets("User Entry").Activate
    'function to stop the view from shifting down when a page is selected
    Me.ScrollTop = 0
done:

    Exit Sub
err:
    MsgBox "Something went wrong. Client may not be referred to that courtroom. " _
      & vbNewLine & vbNewLine & "Message: " & vbNewLine & err.Description _
      & vbNewLine & vbNewLine & "Source: " & vbNewLine & err.Source

End Sub









''''''''''''''''
'FORM_FUNCTIONS'
''''''''''''''''

Sub Cancel_Click()
    Call UnloadAll
End Sub

Sub Clear_Click()

    'for each control (generic name for any field in form
    For Each ctl In Me.Controls

        'determine the type of control it is and reset value accordingly
        Select Case TypeName(ctl)
            Case "TextBox"
                ctl.value = ""
            Case "CheckBox", "ToggleButton"
                ctl.value = False
            Case "OptionGroup"
                ctl = Null
            Case "OptionButton"
                ' Do not reset an optionbutton if it is part of an OptionGroup
                If TypeName(ctl.Parent) <> "OptionGroup" Then ctl.value = False
            Case "ComboBox", "ListBox"
                ctl.listIndex = -1
        End Select
    Next ctl

    Unload Modal_JTC_Accept
    Unload Modal_JTC_Add_Service
    Unload Modal_JTC_Discharge
    Unload Modal_JTC_Drop_Service
    Unload Modal_JTC_Expungement
    Unload Modal_JTC_Phase_Pushback
    Unload Modal_JTC_Provider
    Unload Modal_JTC_Reject
    Unload Modal_JTC_Stepdown
    Unload Modal_JTC_Stepup
    Unload Me

    'run sub that normally fires when form opens
    'Call UserForm_Initialize

End Sub

Sub Adult_Submit_Click()
    On Error GoTo err

    'VALIDATIONS
    If Adult_Legal_Status_Remain.BackColor = unselectedColor And _
            Adult_Legal_Status_Update.BackColor = unselectedColor Then
        MsgBox "Please select a result button for Legal Status"
        Exit Sub
    End If

    If Adult_Decertification_Remain.BackColor = unselectedColor And _
            Adult_Decertification_Update.BackColor = unselectedColor Then
        MsgBox "Please select a result button for Decertification"
        Exit Sub
    End If

    If Adult_Admission_Remain.BackColor = unselectedColor And _
            Adult_Admission_Update.BackColor = unselectedColor Then
        MsgBox "Please select a result button for Admission"
        Exit Sub
    End If

    If Adult_Adjudication_Remain.BackColor = unselectedColor And _
            Adult_Adjudication_Update.BackColor = unselectedColor Then
        MsgBox "Please select a result button for Adjudication"
        Exit Sub
    End If

    If Adult_Continuance_Remain.BackColor = unselectedColor And _
            Adult_Continuance_Update.BackColor = unselectedColor Then
        MsgBox "Please select a result button for Continuance"
        Exit Sub
    End If


    If Not HasContent(Adult_NextCourtDate) Then
        MsgBox "Please enter the next court date"
        Exit Sub
    End If

    Call formSubmitStart(updateRow)
    

    Dim oldCourtHead As String
    Dim oldCourtroom As String
    Dim newCourtHead As String
    Dim newCourtroom As String
    oldCourtroom = "Adult"
    newCourtroom = "Adult"

    If Adult_Return_Reslate.BackColor = selectedColor Then
        newCourtroom = Adult_Reslate_Juvenile_Petition.NextHearingLocation.value
    End If

    oldCourtHead = getCourtroomHead(oldCourtroom)
    newCourtHead = getCourtroomHead(newCourtroom)

    'append PCD
    Call append(Range(headerFind("Previous Court Dates") & updateRow), DateOfHearing.value)

    Range(headerFind("Next Court Date") & updateRow) = Adult_NextCourtDate.value
    Range(headerFind("Listing Type") & updateRow) = Lookup("Listing_Type_Name")(Adult_Next_Hearing_Type.value)


    Dim oldLegalHead As String
    Dim newLegalHead As String
    Dim bucketHead As String
    Dim i As Long

    '''''''''''''''
    ''''RESLATE''''
    '''''''''''''''

    If Adult_Reslate_Update.BackColor = selectedColor Then
        If Modal_Adult_Reslate.Hearing_Outcome.value = "Granted" Then
        
            ''''''''''''''''''''''
            ''Close out Adult CR''
            ''''''''''''''''''''''
            Dim referralDate As String
            referralDate = DateOfHearing.value
                
                
                Range(headerFind("Discharging Courtroom", oldCourtHead) & updateRow).value _
                    = Lookup("Courtroom_Name")(oldCourtroom)
                Range(headerFind("Discharging DA", oldCourtHead) & updateRow).value _
                    = Lookup("DA_Last_Name_Name")(DA.value)
            
                Range(headerFind("End Date", oldCourtHead) & updateRow).value _
                    = referralDate
                Range(headerFind("Nature of Discharge", oldCourtHead) & updateRow).value _
                    = 1 'Positive
                Range(headerFind("Detailed Status Outcome", oldCourtHead) & updateRow).value _
                    = 16 'Transfer to New Del. Room - Positive
                Range(headerFind("LOS", oldCourtHead) & updateRow).value _
                    = calcLOS( _
                        Range(headerFind("Start Date", oldCourtHead) & updateRow).value, _
                        Range(headerFind("End Date", oldCourtHead) & updateRow).value)
                        
                Range(headerFind("Notes on Outcome", oldCourtHead) & updateRow).value _
                    = Adult_Notes.value
                Range(headerFind("Date of Overall Discharge", oldCourtHead) & updateRow).value _
                    = referralDate
                Range(headerFind("Courtroom of Discharge", oldCourtHead) & updateRow).value _
                    = Lookup("Courtroom_Name")(oldCourtroom)
                Range(hFind("DA", "OUTCOMES", "ADULT") & updateRow).value _
                    = Lookup("DA_Last_Name_Name")(DA.value)
                Range(headerFind("Legal Status of Discharge", oldCourtHead) & updateRow).value _
                    = 10 'Adult
                Range(headerFind("Active or Discharged", oldCourtHead) & updateRow).value _
                    = 2 'Discharged
                Range(headerFind("Nature of Courtroom Outcome", oldCourtHead) & updateRow).value _
                    = 1 'Positive
                Range(headerFind("Detailed Courtroom Outcome", oldCourtHead) & updateRow).value _
                    = 16 'Transfer to New Del. Room - Positive
                Range(headerFind("Acquittal or Supervision Discharge?", oldCourtHead) & updateRow).value _
                    = 0 'N/A
                Range(headerFind("Total LOS in Adult", oldCourtHead) & updateRow).value _
                    = calcLOS( _
                        Range(headerFind("Start Date", oldCourtHead) & updateRow).value, _
                        Range(headerFind("End Date", oldCourtHead) & updateRow).value)
                Range(headerFind("Total LOS From Arrest", oldCourtHead) & updateRow).value _
                    = calcLOS( _
                        Range(headerFind("Arrest Date") & updateRow).value, _
                        Range(headerFind("End Date", oldCourtHead) & updateRow).value)
            
            
            ''''''''''''
            ''PETITION''
            ''''''''''''

            Dim petitionHead As String
            petitionHead = hFind("PETITION")
            Dim juvePetitionHead As String
            juvePetitionHead = hFind("JUVENILE PETITION")


            'COPY VALUES FROM PETITION TO JUVENILE PETITION'
            ''''''''''''''''''''''''''''''''''''''''''''''''''
            

            Range(headerFind("Initial Court Date", juvePetitionHead) & updateRow).value _
          = Range(headerFind("Initial Court Date", petitionHead) & updateRow).value
            Range(headerFind("# of Prior Arrests", juvePetitionHead) & updateRow).value _
          = Range(headerFind("# of Prior Arrests", petitionHead) & updateRow).value
            Range(headerFind("Active in System at Time of Arrest?", juvePetitionHead) & updateRow).value _
          = Range(headerFind("Active in System at Time of Arrest?", petitionHead) & updateRow).value
            Range(headerFind("Prior Closed Petitons (Prior Results)", juvePetitionHead) & updateRow).value _
          = Range(headerFind("Prior Closed Petitons (Prior Results)", petitionHead) & updateRow).value

            
            Range(headerFind("Arrest Date", juvePetitionHead) & updateRow).value _
          = Range(headerFind("Arrest Date", petitionHead) & updateRow).value
            Range(headerFind("Day of Arrest", juvePetitionHead) & updateRow).value _
          = Range(headerFind("Day of Arrest", petitionHead) & updateRow).value
            Range(headerFind("Time of Arrest", juvePetitionHead) & updateRow).value _
          = Range(headerFind("Time of Arrest", petitionHead) & updateRow).value
            Range(headerFind("Time Category of Arrest", juvePetitionHead) & updateRow).value _
          = Range(headerFind("Time Category of Arrest", petitionHead) & updateRow).value
            Range(headerFind("Arresting District", juvePetitionHead) & updateRow).value _
          = Range(headerFind("Arresting District", petitionHead) & updateRow).value

            Range(headerFind("Time of Referral to DA", juvePetitionHead) & updateRow).value _
          = Range(headerFind("Time of Referral to DA", petitionHead) & updateRow).value
            Range(headerFind("DC #", juvePetitionHead) & updateRow).value _
          = Range(headerFind("DC #", petitionHead) & updateRow).value
            Range(headerFind("PID #", juvePetitionHead) & updateRow).value _
          = Range(headerFind("PID #", petitionHead) & updateRow).value
            Range(headerFind("DC-PID #", juvePetitionHead) & updateRow).value _
          = Range(headerFind("DC-PID #", petitionHead) & updateRow).value
            Range(headerFind("SID #", juvePetitionHead) & updateRow).value _
          = Range(headerFind("SID #", petitionHead) & updateRow).value

            Range(headerFind("Officer #1", juvePetitionHead) & updateRow).value _
          = Range(headerFind("Officer #1", petitionHead) & updateRow).value
            Range(headerFind("Officer #2", juvePetitionHead) & updateRow).value _
          = Range(headerFind("Officer #2", petitionHead) & updateRow).value
            Range(headerFind("Officer #3", juvePetitionHead) & updateRow).value _
          = Range(headerFind("Officer #3", petitionHead) & updateRow).value
            Range(headerFind("Officer #4", juvePetitionHead) & updateRow).value _
          = Range(headerFind("Officer #4", petitionHead) & updateRow).value
            Range(headerFind("Officer #5", juvePetitionHead) & updateRow).value _
          = Range(headerFind("Officer #5", petitionHead) & updateRow).value

            Range(headerFind("Victim First Name", juvePetitionHead) & updateRow).value _
          = Range(headerFind("Victim First Name", petitionHead) & updateRow).value
            Range(headerFind("Victim Last Name", juvePetitionHead) & updateRow).value _
          = Range(headerFind("Victim Last Name", petitionHead) & updateRow).value
            Range(headerFind("Incident Date", juvePetitionHead) & updateRow).value _
          = Range(headerFind("Incident Date", petitionHead) & updateRow).value
            Range(headerFind("Day of Incident", juvePetitionHead) & updateRow).value _
          = Range(headerFind("Day of Incident", petitionHead) & updateRow).value
            Range(headerFind("Time of Incident", juvePetitionHead) & updateRow).value _
          = Range(headerFind("Time of Incident", petitionHead) & updateRow).value

            Range(headerFind("Time Category of Incident", juvePetitionHead) & updateRow).value _
          = Range(headerFind("Time Category of Incident", petitionHead) & updateRow).value
            Range(headerFind("Incident District", juvePetitionHead) & updateRow).value _
          = Range(headerFind("Incident District", petitionHead) & updateRow).value
            Range(headerFind("Incident Address", juvePetitionHead) & updateRow).value _
          = Range(headerFind("Incident Address", petitionHead) & updateRow).value
            Range(headerFind("Incident Zipcode", juvePetitionHead) & updateRow).value _
          = Range(headerFind("Incident Zipcode", petitionHead) & updateRow).value
            Range(headerFind("Latitude", juvePetitionHead) & updateRow).value _
          = Range(headerFind("Latitude", petitionHead) & updateRow).value

            Range(headerFind("Longitude", juvePetitionHead) & updateRow).value _
          = Range(headerFind("Longitude", petitionHead) & updateRow).value
            Range(headerFind("Incident Violence Zone", juvePetitionHead) & updateRow).value _
          = Range(headerFind("Incident Violence Zone", petitionHead) & updateRow).value
            Range(headerFind("Referred to Diversion?", juvePetitionHead) & updateRow).value _
          = Range(headerFind("Referred to Diversion?", petitionHead) & updateRow).value
            Range(headerFind("Which Diversion Program Used", juvePetitionHead) & updateRow).value _
          = Range(headerFind("Which Diversion Program Used", petitionHead) & updateRow).value
            Range(headerFind("Diversion Referral Date", juvePetitionHead) & updateRow).value _
          = Range(headerFind("Diversion Referral Date", petitionHead) & updateRow).value

            Range(headerFind("DA", juvePetitionHead) & updateRow).value _
          = Range(headerFind("DA", petitionHead) & updateRow).value

            Range(headerFind("Direct Filed?", juvePetitionHead) & updateRow).value _
          = Range(headerFind("Direct Filed?", petitionHead) & updateRow).value
          
            Range(headerFind("School-Based Incident?", juvePetitionHead) & updateRow).value _
          = Range(headerFind("School-Based Incident?", petitionHead) & updateRow).value
            Range(headerFind("School-Based Incident Type", juvePetitionHead) & updateRow).value _
          = Range(headerFind("School-Based Incident Type", petitionHead) & updateRow).value
            Range(headerFind("Home-Based Incident?", juvePetitionHead) & updateRow).value _
          = Range(headerFind("Home-Based Incident?", petitionHead) & updateRow).value
            Range(headerFind("Home-Based Incident Type", juvePetitionHead) & updateRow).value _
          = Range(headerFind("Home-Based Incident Type", petitionHead) & updateRow).value

            With Adult_Reslate_Juvenile_Petition

                Range(headerFind("Initial Hearing Location", juvePetitionHead) & updateRow).value _
                        = 10 'adult
                ''''ADD arrest values from PETITION

                Range(headerFind("Gun Case?", juvePetitionHead) & updateRow).value = Lookup("Generic_YNOU_Name")(.GunCase.value)
                Range(headerFind("Gun Involved Arrest?", juvePetitionHead) & updateRow).value = Lookup("Generic_YNOU_Name")(.GunInvolved.value)

                Range(headerFind("General Notes from Intake", juvePetitionHead) & updateRow).value = Adult_Notes.value
                
                Dim Num As Long
                Dim j As Integer

                For Num = 1 To .PetitionBox.ListCount
                    tempHead = headerFind("Petition #" & Num, juvePetitionHead)

                    If .DiversionProgram.value = "Yes" Then
                        Range(headerFind("Petition Filed?", tempHead) & updateRow).value _
                                = Lookup("Generic_YNOU_Name")("No")
                    Else
                        Range(headerFind("Petition Filed?", tempHead) & updateRow).value _
                                = Lookup("Generic_YNOU_Name")("Yes")
                    End If
                    Range(headerFind("Was Petition Transferred from Other County?", tempHead) & updateRow).value _
                            = Lookup("Generic_YNOU_Name")(.PetitionBox.List(Num - 1, 6))
                    Range(tempHead & updateRow).value _
                            = .PetitionBox.List(Num - 1, 1)
                    Range(headerFind("Date Filed", tempHead) & updateRow).value _
                            = .PetitionBox.List(Num - 1, 0)
                    Range(headerFind("Lead Charge Code", tempHead) & updateRow).value _
                            = .PetitionBox.List(Num - 1, 4)
                    Range(headerFind("Lead Charge Name", tempHead) & updateRow).value _
                            = .PetitionBox.List(Num - 1, 5)
                    Range(headerFind("Charge Category #1", tempHead) & updateRow).value _
                            = Lookup("Charge_Name")(.PetitionBox.List(Num - 1, 3))
                    Range(headerFind("Charge Grade (specific) #1", tempHead) & updateRow).value _
                            = Lookup("Charge_Grade_Specific_Name")(.PetitionBox.List(Num - 1, 2))
                    Range(headerFind("Charge Grade (broad) #1", tempHead) & updateRow).value _
                            = calcChargeBroad(.PetitionBox.List(Num - 1, 2))

                    j = 2
                    For i = 0 To .ChargeBox.ListCount - 1
                        If .ChargeBox.ListCount > 0 Then
                            If .ChargeBox.List(i, 0) = .PetitionBox.List(Num - 1, 1) Then
                                Range(headerFind("Charge Code #" & j, tempHead) & updateRow).value _
                                        = .ChargeBox.List(i, 3)
                                Range(headerFind("Charge Name #" & j, tempHead) & updateRow).value _
                                        = .ChargeBox.List(i, 4)
                                Range(headerFind("Charge Category #" & j, tempHead) & updateRow).value _
                                        = Lookup("Charge_Name")(.ChargeBox.List(i, 2))
                                Range(headerFind("Charge Grade (specific) #" & j, tempHead) & updateRow).value _
                                        = Lookup("Charge_Grade_Specific_Name")(.ChargeBox.List(i, 1))
                                Range(headerFind("Charge Grade (broad) #" & j, tempHead) & updateRow).value _
                                        = calcChargeBroad(.ChargeBox.List(i, 1))
                                j = j + 1
                            End If
                        End If
                    Next i
                Next Num

                Range(headerFind("LOS Until DA Referral", juvePetitionHead) & updateRow).value _
                            = timeDiff(Range(headerFind("Time of Arrest", juvePetitionHead) & updateRow).value, _
                       Range(headerFind("Time of Referral to DA") & updateRow).value)

                ''''''''''''''''''
                'SET LEGAL STATUS'
                ''''''''''''''''''

                If .DiversionProgram.value = "Yes" Then
                    Call legalStatusStart( _
                        clientRow:=updateRow, _
                        statusType:="Diversion", _
                        Courtroom:="PJJSC", _
                        DA:=DA.value, _
                        startDate:=.DiversionProgramReferralDate.value)

                Else
                    If .ConfOutcome.value = "Hold for Detention" _
                    Or .ConfOutcome.value = "Roll to Detention Hearing" Then
                        Call legalStatusStart( _
                            clientRow:=updateRow, _
                            statusType:="Pretrial", _
                            Courtroom:="PJJSC", _
                            DA:=DA.value, _
                            startDate:=.PetitionBox.List(0, 0))
                    Else
                        Call legalStatusStart( _
                            clientRow:=updateRow, _
                            statusType:="Pretrial", _
                            Courtroom:="Intake Conf.", _
                            DA:=DA.value, _
                            startDate:=.PetitionBox.List(0, 0))
                    End If
                End If


                '''''''''''''''''''
                ''''''CALL IN''''''
                '''''''''''''''''''

                tempHead = headerFind("CALL-IN", juvePetitionHead)

                If .CallInRecord.value = "Yes" Then
                    Range(headerFind("Did Youth Have Call-In?", tempHead) & updateRow).value _
                        = Lookup("Generic_NYNOU_Name")("Yes")

                    Range(headerFind("Date of Call-In", tempHead) & updateRow).value _
                            = .CallInDate.value

                    Range(headerFind("Was DRAI Administered?", tempHead) & updateRow).value _
                        = Lookup("Generic_NYNOU_Name")(.Was_DRAI_Administered.value)

                    Range(headerFind("DRAI Score", tempHead) & updateRow).value _
                        = .DRAI_Score.value

                    Select Case .DRAI_Score.value
                        Case Is < 10
                            Range(hFind("DRAI Recommendation", "CALL-IN") & updateRow).value _
                                = Lookup("DRAI_Recommendation_Name")("Release")
                        Case Is < 15
                            Range(hFind("DRAI Recommendation", "CALL-IN") & updateRow).value _
                                = Lookup("DRAI_Recommendation_Name")("Release w/ Supervision")
                        Case Is < 30
                            Range(hFind("DRAI Recommendation", "CALL-IN") & updateRow).value _
                                = Lookup("DRAI_Recommendation_Name")("Release w/ Supervision")
                        Case Else
                            Range(hFind("DRAI Recommendation", "CALL-IN") & updateRow).value _
                                = Lookup("DRAI_Recommendation_Name")("Unknown")
                    End Select


                    Range(headerFind("DRAI Recommendation", tempHead) & updateRow).value _
                        = Lookup("DRAI_Recommendation_Name")(.DRAI_Rec.value)

                    Range(headerFind("DRAI Action", tempHead) & updateRow).value _
                        = Lookup("DRAI_Action_Name")(.DRAI_Action.value)

                    Select Case .DRAI_Action.value
                        Case "Override - Hold", "Follow - Hold"
                            Range(headerFind("End Date", tempHead) & updateRow).value _
                                = .InConfDate.value
                            Range(headerFind("LOS in Detention", tempHead) & updateRow).value _
                                = calcLOS(.CallInDate.value, .InConfDate.value)
                            Call addSupervision( _
                                clientRow:=updateRow, _
                                serviceType:="Detention (not respite)", _
                                legalStatus:="Pretrial", _
                                Courtroom:="Call-In", _
                                CourtroomOfOrder:="Call-In", _
                                DA:=DA.value, _
                                agency:="PJJSC", _
                                startDate:=.CallInDate.value, _
                                endDate:=.InConfDate.value, _
                                re1:="", _
                                re2:="", _
                                re3:="", _
                                Notes:="Held at call-in")

                    End Select
                    Range(headerFind("LOS Until Next Hearing", tempHead) & updateRow).value _
                                = calcLOS(.CallInDate.value, .InConfDate.value)

                    Range(headerFind("Detention Facility", tempHead) & updateRow).value _
                            = Lookup("Detention_Facility_Name")(.DetentionFacility.value)

                    Range(headerFind("Reason #1 for Override Hold", tempHead) & updateRow).value _
                        = Lookup("DRAI_Override_Reason_Name")(.OverrideHoldRe1.value)
                    Range(headerFind("Reason #2 for Override Hold", tempHead) & updateRow).value _
                        = Lookup("DRAI_Override_Reason_Name")(.OverrideHoldRe2.value)
                    Range(headerFind("Reason #3 for Override Hold", tempHead) & updateRow).value _
                        = Lookup("DRAI_Override_Reason_Name")(.OverrideHoldRe3.value)
                Else
                    Range(headerFind("Did Youth Have Call-In?", tempHead) & updateRow).value _
                        = Lookup("Generic_NYNOU_Name")("Unknown")
                End If


                '''''''''''''''''''
                'Intake Conference'
                '''''''''''''''''''

                tempHead = headerFind("INTAKE CONFERENCE", juvePetitionHead)



                If .InConfRecord.value = "Yes" Then
                    Range(headerFind("Did Youth Have an Intake Conference?", tempHead) & updateRow).value _
                        = Lookup("Generic_NYNOU_Name")("Yes")

                    Range(headerFind("Date of Intake Conference", tempHead) & updateRow).value _
                        = .InConfDate.value

                    Range(headerFind("Intake Conference Type", tempHead) & updateRow).value _
                        = Lookup("Intake_Conference_Type_Name")(.InConfType.value)

                    Range(headerFind("DA", tempHead) & updateRow).value _
                        = Lookup("DA_Last_Name_Name")(DA.value)

                    Range(headerFind("Intake Conference Outcome", tempHead) & updateRow).value _
                        = Lookup("Intake_Conference_Outcome_Name")(.ConfOutcome.value)

                    Range(hFind("Status at Arrest", "DHS") & updateRow).value _
                        = Lookup("DHS_Status_at_Arrest_Name")(.DHS_Status.value)
                    
                    If .DHS_Status.value = "N/A" Or .DHS_Status.value = "None" Or .DHS_Status.value = "Unknown" Then
                        Range(hFind("Did youth have any DHS contact?", "DHS") & updateRow).value = 2 'no
                    Else
                        Range(hFind("Did youth have any DHS contact?", "DHS") & updateRow).value = 1 'yes
                    End If
                    
                    
                    Range(headerFind("Location of Next Event", tempHead) & updateRow).value _
                        = Lookup("Courtroom_Name")(.NextHearingLocation.value)

                    Range(headerFind("Next Event Date", tempHead) & updateRow).value _
                        = Adult_NextCourtDate.value

                    tempHead = headerFind("Supervision Ordered #1", tempHead)

                    Range(tempHead & updateRow).value _
                        = Lookup("Supervision_Program_Name")(.Supv1.value)
                    Range(headerFind("Community-Based Agency #1", tempHead) & updateRow).value _
                        = Lookup("Community_Based_Supervision_Provider_Name")(.Supv1Pro.value)

                    Range(headerFind("Reason #1 for Supervision Referral", tempHead) & updateRow).value _
                        = Lookup("Supervision_Referral_Reason_Name")(.Supv1Re1.value)
                    Range(headerFind("Reason #2 for Supervision Referral", tempHead) & updateRow).value _
                        = Lookup("Supervision_Referral_Reason_Name")(.Supv1Re2.value)
                    Range(headerFind("Reason #3 for Supervision Referral", tempHead) & updateRow).value _
                        = Lookup("Supervision_Referral_Reason_Name")(.Supv1Re3.value)

                    tempHead = headerFind("Supervision Ordered #2", tempHead)

                    Range(tempHead & updateRow).value _
                        = Lookup("Supervision_Program_Name")(.Supv2.value)
                    Range(headerFind("Community-Based Agency #2", tempHead) & updateRow).value _
                        = Lookup("Community_Based_Supervision_Provider_Name")(.Supv2Pro.value)


                    Range(headerFind("Reason #1 for Supervision Referral", tempHead) & updateRow).value _
                        = Lookup("Supervision_Referral_Reason_Name")(.Supv2Re1.value)
                    Range(headerFind("Reason #2 for Supervision Referral", tempHead) & updateRow).value _
                        = Lookup("Supervision_Referral_Reason_Name")(.Supv2Re2.value)
                    Range(headerFind("Reason #3 for Supervision Referral", tempHead) & updateRow).value _
                        = Lookup("Supervision_Referral_Reason_Name")(.Supv2Re3.value)

                    Range(headerFind("Other Condition #1", tempHead) & updateRow).value _
                        = Lookup("Condition_Name")(.Cond1.value)
                    Range(headerFind("Other Condition #1 Provider", tempHead) & updateRow).value _
                        = Lookup("Condition_Provider_Name")(.Cond1Pro.value)

                    Range(headerFind("Other Condition #2", tempHead) & updateRow).value _
                        = Lookup("Condition_Name")(.Cond2.value)
                    Range(headerFind("Other Condition #2 Provider", tempHead) & updateRow).value _
                        = Lookup("Condition_Provider_Name")(.Cond2Pro.value)

                    Range(headerFind("Other Condition #3", tempHead) & updateRow).value _
                        = Lookup("Condition_Name")(.Cond3.value)
                    Range(headerFind("Other Condition #3 Provider", tempHead) & updateRow).value _
                        = Lookup("Condition_Provider_Name")(.Cond3Pro.value)

                    
                    Range(headerFind("Diagnosis #1") & updateRow).value = Lookup("Diagnosis_Name")(.Diagnosis1.value)
                    Range(headerFind("Diagnosis #2") & updateRow).value = Lookup("Diagnosis_Name")(.Diagnosis2.value)
                    Range(headerFind("Diagnosis #3") & updateRow).value = Lookup("Diagnosis_Name")(.Diagnosis3.value)
                    Range(headerFind("Trauma Type #1") & updateRow).value = Lookup("Trauma_Type_Name")(.TraumaType1.value)
                    Range(headerFind("Trauma Type #2") & updateRow).value = Lookup("Trauma_Type_Name")(.TraumaType2.value)
                    Range(headerFind("Trauma Type #3") & updateRow).value = Lookup("Trauma_Type_Name")(.TraumaType3.value)
                    Range(headerFind("Treatment #1") & updateRow).value = Lookup("Treatment_Name")(.Treatment1.value)
                    Range(headerFind("Treatment #2") & updateRow).value = Lookup("Treatment_Name")(.Treatment2.value)
                    Range(headerFind("Treatment #3") & updateRow).value = Lookup("Treatment_Name")(.Treatment3.value)
                    
                    
                    
                    
                    
                    Select Case .ConfOutcome.value
                        Case "Hold for Detention"
                            Range(headerFind("Active Courtroom") & updateRow).value _
                                 = Lookup("Courtroom_Name")("PJJSC")
                            Call flagNo(Range(hFind("Did Youth Have Initial Detention Hearing?", "DETENTION") & updateRow))
                            Range(hFind("Detention Facility", "DETENTION") & updateRow).value _
                                 = Lookup("Detention_Facility_Name")(.DetentionFacility.value)
                            Call addSupervision( _
                                clientRow:=updateRow, _
                                serviceType:="Detention (not respite)", _
                                legalStatus:="Pretrial", _
                                Courtroom:="Intake Conf.", _
                                DA:=DA.value, _
                                agency:="", _
                                startDate:=.InConfDate.value, _
                                re1:="", _
                                re2:="", _
                                re3:="", _
                                Notes:="Held at intake conference")
                        Case "Roll to Detention Hearing"
                            Range(headerFind("Active Courtroom") & updateRow).value _
                                 = Lookup("Courtroom_Name")("PJJSC")
                        Case "Release for Court"
                            Call ReferClientTo( _
                                referralDate:=.InConfDate.value, _
                                clientRow:=updateRow, _
                                fromCR:="Intake Conf.", _
                                toCR:=.NextHearingLocation.value, _
                                DA:=DA.value _
                                )
                            If .NextHearingLocation.value = "5E" Then
                                Range(hFind("Courtroom of Origin", "Crossover") & updateRow).value _
                                    = Lookup("Courtroom_Name")("Intake Conf.")
                            Else
                                Range(hFind("Courtroom of Origin", .NextHearingLocation.value) & updateRow).value _
                                    = Lookup("Courtroom_Name")("Intake Conf.")
                            End If

                            'add supervisions and conditions if assigned
                            If Not .Supv1.value = "None" Then
                                Call addSupervision( _
                                clientRow:=updateRow, _
                                serviceType:=.Supv1.value, _
                                legalStatus:="Pretrial", _
                                Courtroom:="Intake Conf.", _
                                CourtroomOfOrder:="Intake Conf.", _
                                DA:=DA.value, _
                                agency:=.Supv1Pro.value, _
                                startDate:=.InConfDate.value, _
                                NextCourtDate:=Adult_NextCourtDate.value, _
                                re1:=.Supv1Re1.value, _
                                re2:=.Supv1Re2.value, _
                                re3:=.Supv1Re3.value, _
                                Notes:="Referred at intake conference")
                            End If

                            If Not .Supv2.value = "None" Then
                                Call addSupervision( _
                                clientRow:=updateRow, _
                                serviceType:=.Supv2.value, _
                                legalStatus:="Pretrial", _
                                Courtroom:="Intake Conf.", _
                                CourtroomOfOrder:="Intake Conf.", _
                                DA:=DA.value, _
                                agency:=.Supv2Pro.value, _
                                startDate:=.InConfDate.value, _
                                NextCourtDate:=Adult_NextCourtDate.value, _
                                re1:=.Supv2Re1.value, _
                                re2:=.Supv2Re2.value, _
                                re3:=.Supv2Re3.value, _
                                Notes:="Referred at intake conference")
                            End If

                            If Not .Cond1.value = "None" Then
                                Call addCondition( _
                                clientRow:=updateRow, _
                                condition:=.Cond1.value, _
                                legalStatus:="Pretrial", _
                                Courtroom:="Intake Conf.", _
                                CourtroomOfOrder:="Intake Conf.", _
                                DA:=DA.value, _
                                agency:=.Cond1Pro.value, _
                                startDate:=.InConfDate.value, _
                                re1:="N/A", _
                                re2:="N/A", _
                                re3:="N/A", _
                                Notes:="Referred at intake conference")
                            End If

                            If Not .Cond2.value = "None" Then
                                Call addCondition( _
                                clientRow:=updateRow, _
                                condition:=.Cond2.value, _
                                legalStatus:="Pretrial", _
                                Courtroom:="Intake Conf.", _
                                CourtroomOfOrder:="Intake Conf.", _
                                DA:=DA.value, _
                                agency:=.Cond2Pro.value, _
                                startDate:=.InConfDate.value, _
                                re1:="N/A", _
                                re2:="N/A", _
                                re3:="N/A", _
                                Notes:="Referred at intake conference")
                            End If

                            If Not .Cond3.value = "None" Then
                                Call addCondition( _
                                clientRow:=updateRow, _
                                condition:=.Cond3.value, _
                                legalStatus:="Pretrial", _
                                Courtroom:="Intake Conf.", _
                                CourtroomOfOrder:="Intake Conf.", _
                                DA:=DA.value, _
                                agency:=.Cond3Pro.value, _
                                startDate:=.InConfDate.value, _
                                re1:="N/A", _
                                re2:="N/A", _
                                re3:="N/A", _
                                Notes:="Referred at intake conference")
                            End If

                        Case "Release for Diversion"

                    End Select
                Else
                    Range(headerFind("Did Youth Have an Intake Conference?", tempHead) & updateRow).value _
                        = Lookup("Generic_NYNOU_Name")("Unknown")

                    Select Case .NextHearingLocation.value
                        Case "4G", "4E", "6F", "6H", "3E", "JTC", "5E", "WRAP"
                            Call ReferClientTo( _
                            referralDate:=.PetitionBox.List(0, 0), _
                            clientRow:=updateRow, _
                            fromCR:="Adult", _
                            toCR:=.NextHearingLocation.value, _
                            DA:=DA.value _
                            )
                    End Select
                End If

                'Range(headerFind("DA") & updateRow).value = Lookup("DA_Last_Name_Name")(DA.value)
                'Range(headerFind("General Notes from Intake") & updateRow).value = .GeneralNotes.value

                '''''''''''''''''''
                '''''DIVERSION'''''
                '''''''''''''''''''

                Dim diversionHead As String

                diversionHead = headerFind("DIVERSION")

                Range(headerFind("Referred to Diversion?", juvePetitionHead) & updateRow) _
                    = Lookup("Generic_YNOU_Name")(.DiversionProgram.value)
                Range(headerFind("Referred to Diversion?", diversionHead) & updateRow) _
                    = Lookup("Generic_YNOU_Name")(.DiversionProgram.value)

                If .DiversionProgram.value = "Yes" Then

                    Range(headerFind("Which Diversion Program Used", juvePetitionHead) & updateRow) _
                        = Lookup("Diversion_Program_Name")(.NameOfProgram.value)
                    Range(headerFind("Diversion Referral Date", juvePetitionHead) & updateRow) _
                        = .DiversionProgramReferralDate.value

                    Range(headerFind("Referral Date", diversionHead) & updateRow) _
                        = .DiversionProgramReferralDate.value
                    Range(headerFind("Referral Source", diversionHead) & updateRow) _
                        = Lookup("Diversion_Referral_Source_Name")(.ReferralSource.value)
                    Range(headerFind("Age at Diversion Referral", diversionHead) & updateRow) _
                        = ageAtTime(.DiversionProgramReferralDate.value, updateRow)
                    Range(headerFind("Diversion Program Ordered", diversionHead) & updateRow) _
                        = Lookup("Diversion_Program_Name")(.NameOfProgram.value)

                    If IsNumeric(.YAPDistrict.value) Then
                        Range(headerFind("YAP Panel District #", diversionHead) & updateRow) _
                            = Lookup("Police_District_Name")(CInt(.YAPDistrict.value))
                    Else
                        Range(headerFind("YAP Panel District #", diversionHead) & updateRow) _
                            = Lookup("Police_District_Name")(.YAPDistrict.value)
                    End If

                    Range(headerFind("Legal Status") & updateRow).value _
                        = Lookup("Legal_Status_Name")("Diversion")

                    Range(headerFind("Did Youth Receive a Review Hearing?", diversionHead) & updateRow) _
                        = 2
                    Range(headerFind("Did Youth Receive an Exit Hearing?", diversionHead) & updateRow) _
                        = 2
                End If

                If .DiversionProgram.value = "No" Then
                    Range(headerFind("Reason #1 Not Diverted", diversionHead) & updateRow) _
                        = Lookup("Diversion_Rejection_Reason_Name")(.NoDiversionReason1.value)
                    Range(headerFind("Reason #2 Not Diverted", diversionHead) & updateRow) _
                        = Lookup("Diversion_Rejection_Reason_Name")(.NoDiversionReason2.value)
                    Range(headerFind("Reason #3 Not Diverted", diversionHead) & updateRow) _
                        = Lookup("Diversion_Rejection_Reason_Name")(.NoDiversionReason3.value)
                End If
            End With
        Else
            MsgBox "Reslate debug: Reslate function not triggered because Hearing_Outcome was  '" & Modal_Adult_Reslate.Hearing_Outcome.value & "' and not 'Granted'"
        End If
    End If


    '''''''''''''''
    'CERTIFICATION'
    '''''''''''''''

    If Adult_Decertification_Update.BackColor = selectedColor Then
        If Adult_Fetch_Decertification.Caption = "Filed" Then
            Call certificationUpdate( _
                updateRow, _
                headerFind("Decertification", oldCourtHead), _
                Modal_Adult_Decertification.Motion_Result, _
                DateOfHearing.value _
            )
            Call certificationUpdate( _
                updateRow, _
                hFind("Certification", "COURT PROCEEDINGS", "AGGREGATES"), _
                Modal_Adult_Decertification.Motion_Result, _
                DateOfHearing.value _
            )
        Else
            Call certificationStart( _
                updateRow, _
                headerFind("Decertification", newCourtHead), _
                Lookup("Legal_Status_Num")(Range(headerFind("Legal Status") & updateRow).value), _
                newCourtroom, _
                DA.value, _
                Modal_Adult_Decertification.Motion_Date.value _
            )
            Call certificationStart( _
                updateRow, _
                hFind("Certification", "COURT PROCEEDINGS", "AGGREGATES"), _
                Lookup("Legal_Status_Num")(Range(headerFind("Legal Status") & updateRow).value), _
                newCourtroom, _
                DA.value, _
                Modal_Adult_Decertification.Motion_Date.value _
            )
        End If
    End If

    '''''''''''
    'ADMISSION'
    '''''''''''
    If Adult_Admission_Update.BackColor = selectedColor Then
        With Modal_Adult_Admission
            Call admissionStart( _
                clientRow:=updateRow, _
                petitionNum:=Modal_Adult_Admission.PetitionBox.value, _
                statusType:="Adult", _
                Courtroom:=newCourtroom, _
                DA:=DA.value, _
                startDate:=Modal_Adult_Admission.Admission_Date.value, _
                Result:=Modal_Adult_Admission.Result.value, _
                detailed:=Modal_Adult_Admission.Detailed_Result.value, _
                leadChargeCode:=.PetitionBox.List(.PetitionBox.listIndex, 4), _
                leadChargeName:=.PetitionBox.List(.PetitionBox.listIndex, 5), _
                chargeCategory:=.PetitionBox.List(.PetitionBox.listIndex, 3), _
                chargeGradeSpecific:=.PetitionBox.List(.PetitionBox.listIndex, 2) _
            )
        End With
    End If




    'Call closeCallIn(DateOfHearing.value, updateRow)
    'Call closeIntakeConference(DateOfHearing.value, updateRow)
    

    Call addNotes( _
        Courtroom:="Adult", _
        DateOf:=DateOfHearing.value, _
        userRow:=updateRow, _
        Notes:=Adult_Notes, _
        DA:=DA.value _
    )
    
    Call UnloadAll
    
    Call formSubmitEnd


done:

    Exit Sub
err:
    Call loadFromCache(2)

    MsgBox "Something went wrong. Database has been restored to state prior to submission. " _
      & vbNewLine & vbNewLine & "Message: " & vbNewLine & err.Description _
      & vbNewLine & vbNewLine & "Source: " & vbNewLine & err.Source
    Call UnloadAll

    'Stop   'press F8 twice to see the error point
    'Resume
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'JTC_SUBMIT_CLICK'''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Sub JTC_Submit_Click()

    On Error GoTo err

    Call formSubmitStart(updateRow)

    Dim oldPhaseHead As String
    Dim newPhaseHead As String
    'VALIDATIONS


    If JTC_Phase_Remain.BackColor = unselectedColor And _
            JTC_Phase_Stepup.BackColor = unselectedColor And _
            JTC_Phase_Pushback.BackColor = unselectedColor And _
            JTC_Discharge.BackColor = unselectedColor Then

        If JTC_Fetch_Phase.Caption = "Referred" _
                Or JTC_Fetch_Phase.Caption = "Graduated, Awaiting Expungment" _
                Or JTC_Fetch_Phase.Caption = "Graduated, Record Expunged" Then
            'do nothing
        Else
            MsgBox "Please select a phase status button"
            Exit Sub
        End If
    End If

    If JTC_Treatment_Provider_Remain.BackColor = unselectedColor And _
            JTC_Treatment_Stepdown.BackColor = unselectedColor And _
            JTC_Treatment_Discharge.BackColor = unselectedColor And _
            JTC_Treatment_Provider_Update.BackColor = unselectedColor Then

        If JTC_Fetch_Phase.Caption = "Referred" _
                Or JTC_Fetch_Phase.Caption = "Graduated, Awaiting Expungment" _
                Or JTC_Fetch_Phase.Caption = "Graduated, Record Expunged" Then
            'do nothing
        Else
            MsgBox "Please select a treatment status button"
            Exit Sub
        End If
    End If


    If NextCourtDate.value = "" Then
        If JTC_Fetch_Phase.Caption = "Referred" _
                Or JTC_Fetch_Phase.Caption = "Graduated, Awaiting Expungment)" _
                Or JTC_Fetch_Phase.Caption = "Graduated, Record Expunged" Then
            'do nothing
        Else
            MsgBox "Please enter the next court date"
            Exit Sub
        End If
    End If
    
    'NOTES
    
    If Not JTC_Notes.value = "" Then
        Select Case JTC_Fetch_Phase.Caption
            Case 2
                Range(headerFind("Notes on Phase 2") & updateRow).value _
                        = DateOfHearing & " - " & JTC_Notes.value & "; " _
                        & vbNewLine & Range(headerFind("Notes on Phase 2") & updateRow).value
            Case 3
                Range(headerFind("Notes on Phase 3") & updateRow).value _
                        = DateOfHearing & " - " & JTC_Notes.value & "; " _
                        & vbNewLine & Range(headerFind("Notes on Phase 3") & updateRow).value
            Case Else
                Range(headerFind("Notes on Phase 1") & updateRow).value _
                        = DateOfHearing & " - " & JTC_Notes.value & "; " _
                        & vbNewLine & Range(headerFind("Notes on Phase 1") & updateRow).value
        End Select
    End If
    Call addNotes( _
        Courtroom:="JTC", _
        DateOf:=DateOfHearing.value, _
        userRow:=updateRow, _
        Notes:=JTC_Notes.value, _
        DA:=DA.value _
    )
    

    'set tempHead at column at beginning of JTC section
    courtHead = headerFind("JTC")

    'set oldPhaseHead to column at beginning of old phase
    Select Case JTC_Fetch_Phase
        Case 1
            oldPhaseHead = headerFind("PHASE 1", courtHead)
        Case 2
            oldPhaseHead = headerFind("PHASE 2", courtHead)
        Case 3
            oldPhaseHead = headerFind("PHASE 3", courtHead)
        Case "Graduated, Awaiting Expungment"
            oldPhaseHead = headerFind("PHASE 3", courtHead)
        Case "Graduated, Record Expunged"
            oldPhaseHead = headerFind("PHASE 3", courtHead)
        Case Else
            oldPhaseHead = courtHead
    End Select

    'set newPhaseHead to column at beginning of new phase
    Select Case JTC_Return_Phase
        Case 1
            newPhaseHead = headerFind("PHASE 1", courtHead)
        Case 2
            newPhaseHead = headerFind("PHASE 2", courtHead)
        Case 3
            newPhaseHead = headerFind("PHASE 3", courtHead)
        Case "Graduated, Awaiting Expungment"
            newPhaseHead = headerFind("PHASE 3", courtHead)
        Case "Graduated, Record Expunged"
            newPhaseHead = headerFind("PHASE 3", courtHead)
        Case Else
            newPhaseHead = courtHead
    End Select

    'Add DoH (Date of Hearing) to "Previous Court Dates"
    'to PCD in current courtroom
    Call append(Range(headerFind("Previous Court Dates", courtHead) & updateRow), DateOfHearing.value)
    'to PCD at front of record
    Call append(Range(headerFind("Previous Court Dates") & updateRow), DateOfHearing.value)

    'Add next court date to "Next Court Date" in court section and at front of record
    Range(headerFind("Next Court Date", courtHead) & updateRow) = NextCourtDate.value
    Range(headerFind("Next Court Date") & updateRow) = NextCourtDate.value
    Range(headerFind("Listing Type") & updateRow) = Lookup("Listing_Type_Name")(JTC_ListingType.value)



    ''''''''''''''''
    'FAIL TO APPEAR'
    ''''''''''''''''

    Dim bucketHead As String

    If JTC_FTA_Yes.BackColor = selectedColor Then
        Call flagYes(Range(hFind("Did Youth FTA?", "AGGREGATES") & updateRow))


        For i = 1 To 15
            If isEmptyOrZero(Range(hFind("FTA #" & i & " Date", "AGGREGATES") & updateRow)) Then

                bucketHead = hFind("FTA #" & i & " Date", "AGGREGATES")

                Range(bucketHead & updateRow).value = DateOfHearing.value
                Range(headerFind("Day of FTA", bucketHead) & updateRow).value _
                        = Weekday(DateOfHearing.value, vbMonday) * 2 - 1
                Range(headerFind("Courtroom", bucketHead) & updateRow).value = Range(hFind("Active Courtroom") & updateRow).value
                Range(headerFind("Legal Status", bucketHead) & updateRow).value = Range(hFind("Legal Status") & updateRow).value

                If i = 1 Then
                    Range(headerFind("LOS to FTA", bucketHead) & updateRow).value _
                            = calcLOS(Range(hFind("Arrest Date") & updateRow).value, DateOfHearing.value)
                Else
                    Range(headerFind("LOS Between FTAs", bucketHead) & updateRow).value _
                            = calcLOS(Range(hFind("FTA #" & (i - 1) & " Date", "AGGREGATES") & updateRow).value, DateOfHearing.value)
                End If


                If Range(hFind("Active B/W?") & updateRow).value = Lookup("Generic_YNOU_Name")("Yes") Then
                    Range(headerFind("B/W Action", bucketHead) & updateRow).value = Lookup("BW_Action_Name")("Continue B/W")
                End If

                If Range(hFind("Active B/W?") & updateRow).value = Lookup("Generic_YNOU_Name")("No") Then
                    If Modal_FTA.BW.value = "Yes" Then
                        Range(headerFind("B/W Action", bucketHead) & updateRow).value = Lookup("BW_Action_Name")("Begin B/W")
                        Call flagYes(Range(hFind("Active B/W?") & updateRow))
                    Else
                        Range(headerFind("B/W Action", bucketHead) & updateRow).value = Lookup("BW_Action_Name")("N/A")
                    End If
                End If

                i = 15
            End If
        Next i
    End If


    If JTC_Lift_BW.BackColor = selectedColor Then
        Range(hFind("Active B/W?") & updateRow).value = Lookup("Generic_YNOU_Name")("No")

        For i = 15 To 1 Step -1
            If isNotEmptyOrZero(Range(hFind("FTA #" & i & " Date", "AGGREGATES") & updateRow)) _
                    And Range(hFind("B/W Action", "FTA #" & i & " Date", "AGGREGATES") & updateRow).value _
                        = Lookup("BW_Action_Name")("Begin B/W") Then

                bucketHead = hFind("FTA #" & i & " Date", "AGGREGATES")

                Range(headerFind("B/W Lifted Date", bucketHead) & updateRow).value _
                        = Modal_Lift_BW.DateBox.value

                Range(headerFind("LOS B/W", bucketHead) & updateRow).value _
                            = calcLOS(Range(bucketHead & updateRow).value, Modal_Lift_BW.DateBox.value)
                i = 1
            End If
        Next i
    End If

    '''''''''''''''
    ''EXPUNGEMENT''
    '''''''''''''''

    If JTC_Fetch_Phase = "Graduated, Awaiting Expungment" Then
        If JTC_Return_Phase = "Graduated, Record Expunged" Then
            Range(headerFind("Phase", courtHead) & updateRow) = Lookup("JTC_Phase_Name")("Graduated, Record Expunged")
            Range(headerFind("Record Expunged?", newPhaseHead) & updateRow) = Lookup("Generic_YN_Name")("Yes")
            Range(headerFind("Date of Expungement", newPhaseHead) & updateRow) = JTC_Accept_Reject_Date.Caption
            Range(headerFind("LOS (expungement)", newPhaseHead) & updateRow) _
                = calcLOS(Range(headerFind("Accepted Date", courtHead) & updateRow), JTC_Accept_Reject_Date.Caption)
        Else
            MsgBox "Invalid submission."
            Exit Sub
        End If
    End If

    '''''''''''''''''
    ''PHASE SECTION''
    '''''''''''''''''
    If JTC_Reject.BackColor = selectedColor Then
        Range(headerFind("Phase", courtHead) & updateRow).value = 7 'Rejected
        Range(headerFind("Accepted (Y/N)", courtHead) & updateRow).value = 2
        Range(headerFind("Rejected Date", courtHead) & updateRow).value = DateOfHearing.value
        Range(headerFind("Next Hearing Location (if rejected)", courtHead) & updateRow).value = _
            Lookup("Courtroom_Name")(Modal_JTC_Reject.ReferredTo.value)
            
        tempHead = headerFind("JTC OUTCOMES")

        Range(headerFind("Notes on Outcome", tempHead) & updateRow) = JTC_Notes.value
        Range(headerFind("Date of Overall Discharge", tempHead) & updateRow) = DateOfHearing.value
        Range(headerFind("Courtroom of Discharge", tempHead) & updateRow) = 8 'JTC
        Range(headerFind("DA", tempHead) & updateRow) = Lookup("DA_Last_Name_Name")(DA.value)
        Range(headerFind("Legal Status of Discharge", tempHead) & updateRow) = 2 'Pretrial
        Range(headerFind("Active or Discharged", tempHead) & updateRow) = Lookup("Active_Name")("Discharged")
        Range(headerFind("Nature of Courtroom Outcome", tempHead) & updateRow) = 3 'neutral
        Range(headerFind("Detailed Courtroom Outcome", tempHead) & updateRow) = 9 'Acceptance Not Granted
        
        Range(headerFind("Total LOS in JTC", tempHead) & updateRow) _
        = calcLOS(Range(headerFind("Referral Date", courtHead) & updateRow), DateOfHearing.value)
        
        Range(headerFind("Total LOS from Arrest", tempHead) & updateRow) _
        = calcLOS(Range(headerFind("Arrest Date") & updateRow), DateOfHearing.value)
 

        Call ReferClientTo( _
            referralDate:=DateOfHearing.value, _
            clientRow:=updateRow, _
            toCR:=Modal_JTC_Reject.ReferredTo.value, _
            fromCR:="JTC", _
            Notes:="Rejected from JTC", _
            DA:=DA.value)
        'Call Cancel_Click
        'Worksheets("User Entry").Activate
        'Exit Sub
    End If

    If JTC_Accept.BackColor = selectedColor Then
        Range(headerFind("Phase", courtHead) & updateRow).value = "1"
        Range(headerFind("Accepted (Y/N)", courtHead) & updateRow).value = 1 'Yes
        Range(headerFind("Accepted Date", courtHead) & updateRow).value = JTC_Accept_Reject_Date.Caption
        Range(headerFind("Start Date", newPhaseHead) & updateRow).value = JTC_Accept_Reject_Date.Caption
        Range(headerFind("Scheduled Step-Up Date", newPhaseHead) & updateRow) = JTC_Return_Stepup_Date.Caption

        Call legalStatusEnd( _
            clientRow:=updateRow, _
            statusType:=Lookup("Legal_Status_Num")(Range(hFind("Legal Status") & updateRow).value), _
            Courtroom:=Lookup("Courtroom_Num")(Range(hFind("Courtroom of Origin", "JTC") & updateRow).value), _
            DA:=DA.value, _
            endDate:=JTC_Accept_Reject_Date.Caption, _
            Nature:="Neutral", _
            detailed:="Neutral Transfer of Status", _
            Notes:="Accepted to JTC", _
            withAgg:=True, _
            dischargingCourtroom:="JTC")

        Call legalStatusStart( _
            clientRow:=updateRow, _
            statusType:="JTC", _
            Courtroom:="JTC", _
            DA:=DA.value, _
            startDate:=DateOfHearing.value, _
            Notes:="Accepted to JTC")

        Call closeOpenLegalStatuses( _
            clientRow:=updateRow, _
            DateOf:=DateOfHearing.value, _
            Courtroom:="JTC", _
            legalStatus:="JTC", _
            DA:=DA.value)
    End If

    'if step-up is selected
    If JTC_Phase_Stepup.BackColor = selectedColor Then
        'set DoH as old phase end date and new phase begin date
        Range(headerFind("End Date", oldPhaseHead) & updateRow) = DateOfHearing.value
        Range(headerFind("Start Date", newPhaseHead) & updateRow) = DateOfHearing.value
        Range(headerFind("LOS", oldPhaseHead) & updateRow) _
                        = calcLOS(Range(headerFind("Start Date", oldPhaseHead) & updateRow), DateOfHearing.value)
        'Update phase
        Range(headerFind("Phase", courtHead) & updateRow) = JTC_Return_Phase.Caption
        'set new stepup
        Range(headerFind("Scheduled Step-Up Date", newPhaseHead) & updateRow) = JTC_Return_Stepup_Date.Caption
    End If

    'if pushback
    If JTC_Phase_Pushback.BackColor = selectedColor Then
        'find first empty Push-Back date and enter date and reasons
        Select Case True
            Case IsEmpty(Range(headerFind("Push-Back Date #1", oldPhaseHead) & updateRow))
                Range(headerFind("Push-Back Date #1", oldPhaseHead) & updateRow) = _
                                JTC_Return_Stepup_Date.Caption

                Range(headerFind("Reason #1", headerFind("Push-Back Date #1", oldPhaseHead)) & updateRow) = _
                                Lookup("Negative_Discharge_Reason_Name")(JTC_Pushback_Reason1.Caption)
                Range(headerFind("Reason #2", headerFind("Push-Back Date #1", oldPhaseHead)) & updateRow) = _
                                Lookup("Negative_Discharge_Reason_Name")(JTC_Pushback_Reason2.Caption)
                Range(headerFind("Reason #3", headerFind("Push-Back Date #1", oldPhaseHead)) & updateRow) = _
                                Lookup("Negative_Discharge_Reason_Name")(JTC_Pushback_Reason3.Caption)

            Case IsEmpty(Range(headerFind("Push-Back Date #2", oldPhaseHead) & updateRow))
                Range(headerFind("Push-Back Date #2", oldPhaseHead) & updateRow) = _
                                JTC_Return_Stepup_Date.Caption

                Range(headerFind("Reason #1", headerFind("Push-Back Date #2", oldPhaseHead)) & updateRow) = _
                                Lookup("Negative_Discharge_Reason_Name")(JTC_Pushback_Reason1.Caption)
                Range(headerFind("Reason #2", headerFind("Push-Back Date #2", oldPhaseHead)) & updateRow) = _
                                Lookup("Negative_Discharge_Reason_Name")(JTC_Pushback_Reason2.Caption)
                Range(headerFind("Reason #3", headerFind("Push-Back Date #2", oldPhaseHead)) & updateRow) = _
                                Lookup("Negative_Discharge_Reason_Name")(JTC_Pushback_Reason3.Caption)

            Case IsEmpty(Range(headerFind("Push-Back Date #3", oldPhaseHead) & updateRow))
                Range(headerFind("Push-Back Date #3", oldPhaseHead) & updateRow) = _
                                JTC_Return_Stepup_Date.Caption

                Range(headerFind("Reason #1", headerFind("Push-Back Date #3", oldPhaseHead)) & updateRow) = _
                                Lookup("Negative_Discharge_Reason_Name")(JTC_Pushback_Reason1.Caption)
                Range(headerFind("Reason #2", headerFind("Push-Back Date #3", oldPhaseHead)) & updateRow) = _
                                Lookup("Negative_Discharge_Reason_Name")(JTC_Pushback_Reason2.Caption)
                Range(headerFind("Reason #3", headerFind("Push-Back Date #3", oldPhaseHead)) & updateRow) = _
                                Lookup("Negative_Discharge_Reason_Name")(JTC_Pushback_Reason3.Caption)
            Case Else
        End Select
    End If

    
    ''''''''''''''''
    ''IOP PROVIDER''
    ''''''''''''''''

    'if change provider
    If JTC_Treatment_Provider_Update.BackColor = selectedColor Then
        Range(headerFind("IOP Provider") & updateRow) = Lookup("IOP_Provider_Name")(JTC_Return_Treatment_Provider.Caption)
        'find latest entry & update provider
        Select Case True
            Case isEmptyOrZero(Range(hFind("IOP Provider #1", "JTC") & updateRow))
                Range(hFind("IOP Provider #1", "JTC") & updateRow) = _
                                Lookup("IOP_Provider_Name")(JTC_Return_Treatment_Provider.Caption)
                Range(hFind("Referral Date", "IOP Provider #1", "JTC") & updateRow) = _
                                Modal_JTC_Provider.Referral_Date

            Case isEmptyOrZero(Range(hFind("IOP Provider #2", "JTC") & updateRow))
                Range(hFind("IOP Provider #2", "JTC") & updateRow) = _
                                Lookup("IOP_Provider_Name")(JTC_Return_Treatment_Provider.Caption)
                Range(hFind("Referral Date", "IOP Provider #2", "JTC") & updateRow) = Modal_JTC_Provider.Referral_Date
                Range(hFind("Discharge Date", "IOP Provider #1", "JTC") & updateRow) = DateOfHearing
                Range(hFind("LOS IOP", "IOP Provider #1", "JTC") & updateRow) = _
                                calcLOS(Range(headerFind("Referral Date", hFind("IOP Provider #1", "JTC")) & updateRow).value, Modal_JTC_Provider.Referral_Date.value)

            Case isEmptyOrZero(Range(hFind("IOP Provider #3", "JTC") & updateRow))
                Range(hFind("IOP Provider #3", "JTC") & updateRow) = _
                                Lookup("IOP_Provider_Name")(JTC_Return_Treatment_Provider.Caption)
                Range(hFind("Referral Date", "IOP Provider #3", "JTC") & updateRow) = Modal_JTC_Provider.Referral_Date
                Range(hFind("Discharge Date", "IOP Provider #2", "JTC") & updateRow) = DateOfHearing
                Range(hFind("LOS IOP", "IOP Provider #2", "JTC") & updateRow) = _
                                calcLOS(Range(headerFind("Referral Date", headerFind("IOP Provider #2", courtHead)) & updateRow).value, Modal_JTC_Provider.Referral_Date.value)

            Case Else
                Range(hFind("Discharge Date", hFind("IOP Provider #3", "JTC")) & updateRow) = DateOfHearing
                Range(hFind("LOS IOP", "IOP Provider #3", "JTC") & updateRow) = _
                                calcLOS(Range(hFind("Referral Date", "IOP Provider #3", "JTC") & updateRow).value, Modal_JTC_Provider.Referral_Date.value)
        End Select

    End If

    'if stepdown
    If JTC_Treatment_Stepdown.BackColor = selectedColor Then
        'find latest entry & enter DoH as stepdown date
        Select Case False
            Case isEmptyOrZero(Range(hFind("IOP Provider #3", "JTC") & updateRow))
                Range(hFind("Step-Down Date", "IOP Provider #3", "JTC") & updateRow) = JTC_Return_Stepdown_Date
            Case isEmptyOrZero(Range(hFind("IOP Provider #2", "JTC") & updateRow))
                Range(hFind("Step-Down Date", "IOP Provider #2", "JTC") & updateRow) = JTC_Return_Stepdown_Date
            Case isEmptyOrZero(Range(hFind("IOP Provider #1", "JTC") & updateRow))
                Range(hFind("Step-Down Date", "IOP Provider #1", "JTC") & updateRow) = JTC_Return_Stepdown_Date
        End Select
    End If

    'if treatment discharge
    If JTC_Treatment_Discharge.BackColor = selectedColor Then
        'find latest entry & enter DoH as discharge date
        Select Case True
            Case Range(hFind("IOP Provider #3", "JTC") & updateRow).value = Lookup("IOP_Provider_Name")(JTC_Fetch_Treatment_Provider.Caption)
                Range(hFind("Discharge Date", "IOP Provider #3", "JTC") & updateRow) = DateOfHearing
                Range(hFind("LOS IOP", "IOP Provider #3", "JTC") & updateRow) _
                    = calcLOS(Range(hFind("Referral Date", "IOP Provider #3", "JTC") & updateRow), _
                      Range(hFind("Discharge Date", "IOP Provider #3", "JTC") & updateRow))

            Case Range(hFind("IOP Provider #2", "JTC") & updateRow).value = Lookup("IOP_Provider_Name")(JTC_Fetch_Treatment_Provider.Caption)
                Range(hFind("Discharge Date", "IOP Provider #2", "JTC") & updateRow) = DateOfHearing
                Range(hFind("LOS IOP", "IOP Provider #2", "JTC") & updateRow) _
                    = calcLOS(Range(hFind("Referral Date", "IOP Provider #2", "JTC") & updateRow), _
                      Range(hFind("Discharge Date", "IOP Provider #2", "JTC") & updateRow))

            Case Range(hFind("IOP Provider #1", "JTC") & updateRow).value = Lookup("IOP_Provider_Name")(JTC_Fetch_Treatment_Provider.Caption)
                Range(hFind("Discharge Date", "IOP Provider #1", "JTC") & updateRow) = DateOfHearing
                Range(hFind("LOS IOP", "IOP Provider #1", "JTC") & updateRow) _
                      = calcLOS(Range(hFind("Referral Date", "IOP Provider #1", "JTC") & updateRow), _
                        Range(hFind("Discharge Date", "IOP Provider #1", "JTC") & updateRow))
        End Select
    End If




    If JTC_Certification_Update.BackColor = selectedColor Then
        If JTC_Fetch_Certification.Caption = "Filed" Then
            Call certificationUpdate( _
                    updateRow, _
                    hFind("Certification", "COURT PROCEEDINGS", "JTC"), _
                    Modal_JTC_Certification.Motion_Result, _
                    DateOfHearing.value _
                )
            Call certificationUpdate( _
                    updateRow, _
                    hFind("Certification", "COURT PROCEEDINGS", "AGGREGATES"), _
                    Modal_JTC_Certification.Motion_Result, _
                    DateOfHearing.value _
                )
        Else
            Call certificationStart( _
                    updateRow, _
                    hFind("Certification", "COURT PROCEEDINGS", "JTC"), _
                    "JTC", _
                    "JTC", _
                    DA.value, _
                    Modal_JTC_Certification.Motion_Date.value _
                )
            Call certificationStart( _
                    updateRow, _
                    hFind("Certification", "COURT PROCEEDINGS", "AGGREGATES"), _
                    "JTC", _
                    "JTC", _
                    DA.value, _
                    Modal_JTC_Certification.Motion_Date.value _
                )
        End If
    End If

    If JTC_Admission_Update.BackColor = selectedColor Then
        With Modal_JTC_Admission
        
            Call admissionStart( _
                clientRow:=updateRow, _
                petitionNum:=.PetitionBox.value, _
                statusType:="JTC", _
                Courtroom:="JTC", _
                DA:=DA.value, _
                startDate:=.Admission_Date.value, _
                Result:=.Result.value, _
                detailed:=.Detailed_Result.value, _
                leadChargeCode:=.PetitionBox.List(.PetitionBox.listIndex, 4), _
                leadChargeName:=.PetitionBox.List(.PetitionBox.listIndex, 5), _
                chargeCategory:=.PetitionBox.List(.PetitionBox.listIndex, 3), _
                chargeGradeSpecific:=.PetitionBox.List(.PetitionBox.listIndex, 2) _
            )
            End With
    End If

    If JTC_Adjudication_Update.BackColor = selectedColor Then
        With Modal_JTC_Adjudication
            Call adjudicationStart( _
                clientRow:=updateRow, _
                petitionNum:=.PetitionBox.value, _
                Courtroom:="JTC", _
                DA:=DA.value, _
                startDate:=.Adjudication_Date.value, _
                Type_of:=.Type_of.value, _
                re1:=.Reason1.value, _
                re2:=.Reason2.value, _
                re3:=.Reason3.value, _
                re4:=.Reason4.value, _
                re5:=.Reason5.value, _
                leadChargeCode:=.PetitionBox.List(.PetitionBox.listIndex, 4), _
                leadChargeName:=.PetitionBox.List(.PetitionBox.listIndex, 5), _
                chargeCategory:=.PetitionBox.List(.PetitionBox.listIndex, 3), _
                chargeGradeSpecific:=.PetitionBox.List(.PetitionBox.listIndex, 2) _
            )
        End With
    End If

    If JTC_Continuance_Update.BackColor = selectedColor Then
        Call continuanceStart( _
                updateRow, _
                Modal_JTC_Continuance.Status, _
                "JTC", _
                DA.value, _
                DateOfHearing.value, _
                NextCourtDate.value, _
                Modal_JTC_Continuance.Continuance_Type.value, _
                Modal_JTC_Continuance.Reason1.value, Modal_JTC_Continuance.Reason2.value, Modal_JTC_Continuance.Reason3.value _
            )
    End If


    'In listbox representing after-court status
    With JTC_Return_Service_Box

        Dim j As Long
        Dim service As String

        'we will use this for loop to iterate through all of the rows in the listbox
        '
        For i = 0 To .ListCount - 1
            If .List(i, 4) = "New" Then
                Call addSupervision( _
                    clientRow:=updateRow, _
                    serviceType:=.List(i, 0), _
                    legalStatus:="JTC", _
                    Courtroom:="JTC", _
                    DA:=DA.value, _
                    agency:=.List(i, 1), _
                    startDate:=.List(i, 2), _
                    NextCourtDate:=NextCourtDate.value, _
                    re1:=decodeReasons(.List(i, 6))(0), _
                    re2:=decodeReasons(.List(i, 6))(1), _
                    re3:=decodeReasons(.List(i, 6))(2), _
                    re4:=decodeReasons(.List(i, 6))(3), _
                    re5:=decodeReasons(.List(i, 6))(4), _
                    Notes:=.List(i, 9), _
                    phase:=JTC_Return_Phase.Caption)
            Else
            
                If .List(i, 4) = "JTC" Then
                'if service ordered from this courtroom
                
                    If IsDate(.List(i, 3)) Then
                    'if has End Date
                    
                        Call dropSupervision( _
                            clientRow:=updateRow, _
                            Courtroom:=.List(i, 4), _
                            serviceType:=.List(i, 0), _
                            startDate:=.List(i, 2), _
                            endDate:=.List(i, 3), _
                            Nature:=.List(i, 5), _
                            re1:=decodeReasons(.List(i, 6))(0), _
                            re2:=decodeReasons(.List(i, 6))(1), _
                            re3:=decodeReasons(.List(i, 6))(2), _
                            re4:=decodeReasons(.List(i, 6))(3), _
                            re5:=decodeReasons(.List(i, 6))(4), _
                            Notes:=.List(i, 9))
                    End If
                Else
                ' if service ordered from somewhere else
                
                    If IsDate(.List(i, 3)) Then
                    'if has end date
            
                        Call dropSupervision( _
                            clientRow:=updateRow, _
                            Courtroom:=.List(i, 4), _
                            serviceType:=.List(i, 0), _
                            startDate:=.List(i, 2), _
                            endDate:=.List(i, 3), _
                            Nature:=.List(i, 5), _
                            re1:=decodeReasons(.List(i, 6))(0), _
                            re2:=decodeReasons(.List(i, 6))(1), _
                            re3:=decodeReasons(.List(i, 6))(2), _
                            re4:=decodeReasons(.List(i, 6))(3), _
                            re5:=decodeReasons(.List(i, 6))(4), _
                            Notes:=.List(i, 9))
                    Else
                    'if continued from somewhere else
            
                        Call dropSupervision( _
                            clientRow:=updateRow, _
                            Courtroom:=.List(i, 4), _
                            serviceType:=.List(i, 0), _
                            startDate:=.List(i, 2), _
                            endDate:=DateOfHearing.value, _
                            Nature:="Neutral", _
                            re1:="N/A", _
                            re2:="N/A", _
                            re3:="N/A", _
                            Notes:="Continued in JTC")
    
                        Call addSupervision( _
                            clientRow:=updateRow, _
                            serviceType:=.List(i, 0), _
                            legalStatus:="JTC", _
                            Courtroom:="JTC", _
                            DA:=DA.value, _
                            agency:=.List(i, 1), _
                            startDate:=DateOfHearing.value, _
                            NextCourtDate:=NextCourtDate.value, _
                            Notes:="Continued from " & .List(i, 4), _
                            phase:=JTC_Return_Phase.Caption)
                    End If
                End If
            End If
        Next i
    End With

    With JTC_Return_Condition_Box
        Dim condition As String

        'we will use this for loop to iterate through all of the rows in the listbox
        '
        For i = 0 To .ListCount - 1
            If .List(i, 4) = "New" Then
                Call addCondition( _
                    clientRow:=updateRow, _
                    condition:=.List(i, 0), _
                    legalStatus:="JTC", _
                    Courtroom:="JTC", _
                    DA:=DA.value, _
                    agency:=.List(i, 1), _
                    startDate:=.List(i, 2), _
                    re1:=decodeReasons(.List(i, 6))(0), _
                    re2:=decodeReasons(.List(i, 6))(1), _
                    re3:=decodeReasons(.List(i, 6))(2), _
                    re4:=decodeReasons(.List(i, 6))(3), _
                    re5:=decodeReasons(.List(i, 6))(4), _
                    Notes:=.List(i, 9), _
                    phase:=JTC_Return_Phase.Caption)
                    
            Select Case .List(i, 0)
                    Case "Restitution"
                        Call startRestitution( _
                            Amount:=ClientUpdateForm.JTC_Restitution.Caption, _
                            Courtroom:="JTC", _
                            DA:=DA.value, _
                            DateOf:=.List(i, 2), _
                            userRow:=updateRow)
                    
                    Case "Comm. Serv"
                        Call startCommService( _
                            Amount:=ClientUpdateForm.JTC_Comm_Service.Caption, _
                            Courtroom:="JTC", _
                            DA:=DA.value, _
                            DateOf:=.List(i, 2), _
                            userRow:=updateRow)
                    Case "Court Costs"
                        Call startCourtCosts( _
                            Amount:=ClientUpdateForm.JTC_Court_Costs.Caption, _
                            Courtroom:="JTC", _
                            DA:=DA.value, _
                            DateOf:=.List(i, 2), _
                            userRow:=updateRow)
                End Select
                
            Else
                If .List(i, 4) = "JTC" Then
                'if service ordered from this courtroom
                
                    If IsDate(.List(i, 3)) Then
                    'if has End Date
                    
                        Call dropCondition( _
                            clientRow:=updateRow, _
                            Courtroom:=.List(i, 4), _
                            condition:=.List(i, 0), _
                            startDate:=.List(i, 2), _
                            endDate:=.List(i, 3), _
                            Nature:=.List(i, 5), _
                            re1:=decodeReasons(.List(i, 6))(0), _
                            re2:=decodeReasons(.List(i, 6))(1), _
                            re3:=decodeReasons(.List(i, 6))(2), _
                            re4:=decodeReasons(.List(i, 6))(3), _
                            re5:=decodeReasons(.List(i, 6))(4), _
                            Notes:=.List(i, 9))
                    End If
                Else
                ' if service ordered from somewhere else
                
                    If IsDate(.List(i, 3)) Then
                    'if has end date
            
                        Call dropCondition( _
                            clientRow:=updateRow, _
                            Courtroom:=.List(i, 4), _
                            condition:=.List(i, 0), _
                            startDate:=.List(i, 2), _
                            endDate:=.List(i, 3), _
                            Nature:=.List(i, 5), _
                            re1:=decodeReasons(.List(i, 6))(0), _
                            re2:=decodeReasons(.List(i, 6))(1), _
                            re3:=decodeReasons(.List(i, 6))(2), _
                            re4:=decodeReasons(.List(i, 6))(3), _
                            re5:=decodeReasons(.List(i, 6))(4), _
                            Notes:=.List(i, 9))
                    Else
                    'if continued from somewhere else
            
                        Call dropCondition( _
                            clientRow:=updateRow, _
                            Courtroom:=.List(i, 4), _
                            condition:=.List(i, 0), _
                            startDate:=.List(i, 2), _
                            endDate:=DateOfHearing.value, _
                            Nature:="Neutral", _
                            re1:="N/A", _
                            re2:="N/A", _
                            re3:="N/A", _
                            Notes:="Continued in JTC")
    
                        Call addCondition( _
                            clientRow:=updateRow, _
                            condition:=.List(i, 0), _
                            legalStatus:="JTC", _
                            Courtroom:="JTC", _
                            DA:=DA.value, _
                            agency:=.List(i, 1), _
                            startDate:=DateOfHearing.value, _
                            Notes:="Continued from " & .List(i, 4), _
                            phase:=JTC_Return_Phase.Caption)
                    End If
                End If
            End If
        Next i
    End With
    
    
    'if discharge
    If JTC_Discharge.BackColor = selectedColor Then
        Dim endHead As String
        endHead = hFind("Petition Outcomes", "AGGREGATES")
        tempHead = headerFind("JTC OUTCOMES")
        'set current phase end date
        Range(headerFind("End Date", oldPhaseHead) & updateRow) = DateOfHearing.value
        Range(headerFind("LOS", oldPhaseHead) & updateRow) _
                = calcLOS(Range(headerFind("Start Date", oldPhaseHead) & updateRow), DateOfHearing.value)
        'outcome notes
        Range(headerFind("Notes on Outcome", tempHead) & updateRow) = JTC_Notes.value
        'set "Date of Overall Discharge"
        Range(headerFind("Date of Overall Discharge", tempHead) & updateRow) = DateOfHearing.value
        'set "Active or Discharged" to discharged
        
        Range(headerFind("Courtroom of Discharge", tempHead) & updateRow) = Lookup("Courtroom_Name")("JTC")
        Range(headerFind("DA", tempHead) & updateRow) = Lookup("DA_Last_Name_Name")(DA.value)
        
        If Modal_JTC_Discharge.Legal_Status.value = "" Then
            Range(headerFind("Legal Status of Discharge", tempHead) & updateRow) = Lookup("Legal_Status_Name")("JTC")
        Else
            Range(headerFind("Legal Status of Discharge", tempHead) & updateRow) = Lookup("Legal_Status_Name")(Modal_JTC_Discharge.Legal_Status.value)
        End If
        
        Range(headerFind("Active or Discharged", tempHead) & updateRow) = Lookup("Active_Name")("Discharged")
        'set "Nature of Discharge"

        Select Case Modal_JTC_Discharge.DetailedOutcome.value
            Case "Rearrested & Held (adult)"
                Call totalOutcome( _
                    clientRow:=updateRow, _
                    DateOf:=DateOfHearing.value, _
                    Courtroom:="JTC", _
                    DA:=DA.value, _
                    legalStatus:="JTC", _
                    Nature:=Modal_JTC_Discharge.NatureOfOutcome.value, _
                    detailed:="Rearrested & Held (adult)", _
                    Notes:=JTC_Notes.value)

                ''''''''''''''''''''''
            Case "Positive Completion"
                Call totalOutcome( _
                    clientRow:=updateRow, _
                    DateOf:=DateOfHearing.value, _
                    Courtroom:="JTC", _
                    DA:=DA.value, _
                    legalStatus:="JTC", _
                    Nature:=Modal_JTC_Discharge.NatureOfOutcome.value, _
                    detailed:="Petition Closed - Positive Comp. Terms", _
                    Notes:=JTC_Notes.value)

                If JTC_Fetch_Phase = 3 Then
                    Range(headerFind("Phase", courtHead) & updateRow) = Lookup("JTC_Phase_Name")("Graduated, Awaiting Expungment")
                    Range(headerFind("Record Expunged?", newPhaseHead) & updateRow) = Lookup("Generic_YN_Name")("No")
                    Range(headerFind("LOS", oldPhaseHead) & updateRow).value _
                        = calcLOS(Range(headerFind("Start Date", oldPhaseHead) & updateRow).value, DateOfHearing.value)
                End If
                Range(headerFind("LOS (discharged)") & updateRow).value _
                    = calcLOS(Range(headerFind("Arrest Date") & updateRow).value, DateOfHearing.value)

                '''''''''''''''''''''
            Case "Aged Out"
                Call totalOutcome( _
                    clientRow:=updateRow, _
                    DateOf:=DateOfHearing.value, _
                    Courtroom:="JTC", _
                    DA:=DA.value, _
                    legalStatus:="JTC", _
                    Nature:=Modal_JTC_Discharge.NatureOfOutcome.value, _
                    detailed:="Aged Out", _
                    Notes:=JTC_Notes.value)
                    
             Case "Admin. D/C - Reasonable Efforts"
                Call totalOutcome( _
                    clientRow:=updateRow, _
                    DateOf:=DateOfHearing.value, _
                    Courtroom:="JTC", _
                    DA:=DA.value, _
                    legalStatus:="JTC", _
                    Nature:=Modal_JTC_Discharge.NatureOfOutcome.value, _
                    detailed:="Admin. D/C - Reasonable Efforts", _
                    Notes:=JTC_Notes.value)

                '''''''''''''''''''''
            Case "Acceptance Not Granted", "Show Cause", "Hosp. (Mental Health)", "Hosp. (Physical Health)", "Other", "Unknown", "Not Fit to Stand Trial"
                Call ReferClientTo( _
                    referralDate:=DateOfHearing.value, _
                    clientRow:=updateRow, _
                    toCR:=Modal_JTC_Discharge.New_CR.value, _
                    fromCR:="JTC", _
                    newLegalStatus:=Modal_JTC_Discharge.Legal_Status.value, _
                    DA:=DA.value)

                '''''''''''''''''''''
            Case "Transfer to Dependent"
                Call ReferClientTo( _
                    referralDate:=DateOfHearing.value, _
                    clientRow:=updateRow, _
                    toCR:="5E", _
                    fromCR:="JTC", _
                    DA:=DA.value)

                '''''''''''''''''''''
            Case "Transfer to Other County"
                Call totalOutcome( _
                    clientRow:=updateRow, _
                    DateOf:=DateOfHearing.value, _
                    Courtroom:="JTC", _
                    DA:=DA.value, _
                    legalStatus:="JTC", _
                    Nature:=Modal_JTC_Discharge.NatureOfOutcome.value, _
                    detailed:="Transfer to Other County", _
                    Notes:=JTC_Notes.value)

        End Select


        'set detailed outcome
        Range(headerFind("Detailed Courtroom Outcome", tempHead) & updateRow) = _
                        Lookup("JTC_Outcome_Name")(Modal_JTC_Discharge.DetailedOutcome.value)
        Range(headerFind("Nature of Courtroom Outcome", tempHead) & updateRow) = _
                        Lookup("Nature_of_Discharge_Name")(Modal_JTC_Discharge.NatureOfOutcome.value)
        'if negative
        'set reasons for dischrage
        Range(headerFind("Reason #1 for Negative Discharge", tempHead) & updateRow) = _
                            Lookup("Negative_Discharge_Reason_Name")(Modal_JTC_Discharge.ReasonForDischarge1.value)
        Range(headerFind("Reason #2 for Negative Discharge", tempHead) & updateRow) = _
                            Lookup("Negative_Discharge_Reason_Name")(Modal_JTC_Discharge.ReasonForDischarge2.value)
        Range(headerFind("Reason #3 for Negative Discharge", tempHead) & updateRow) = _
                            Lookup("Negative_Discharge_Reason_Name")(Modal_JTC_Discharge.ReasonForDischarge3.value)
        'calc LOS in JTC
        Range(headerFind("Total LOS in JTC", tempHead) & updateRow) = _
                        calcLOS(Range(headerFind("Referral Date", courtHead) & updateRow), DateOfHearing)
        'calc LOS from Arrest
        Range(headerFind("Total LOS from Arrest", tempHead) & updateRow) = _
                        calcLOS(Range(headerFind("Arrest Date") & updateRow), DateOfHearing)
        '######set T/F discharge reasons
    End If



    Call closeCallIn(DateOfHearing.value, updateRow)
    Call closeIntakeConference(DateOfHearing.value, updateRow)
    
    Call UnloadAll

    Call formSubmitEnd

done:
    Exit Sub
err:

    Call loadFromCache(2)
    Stop   'press F8 twice to see the error point
    Resume
    MsgBox "Something went wrong. Database has been restored to state prior to submission. " _
      & vbNewLine & vbNewLine & "Message: " & vbNewLine & err.Description _
      & vbNewLine & vbNewLine & "Source: " & vbNewLine & err.Source
    Call UnloadAll
End Sub

Sub Standard_Submit_Click()
    On Error GoTo err

    Call formSubmitStart(updateRow)

    'VALIDATIONS
    If Standard_Legal_Status_Remain.BackColor = unselectedColor And _
            Standard_Legal_Status_Update.BackColor = unselectedColor Then
        MsgBox "Please select a result button for Legal Status"
        Exit Sub
    End If

    If Standard_Certification_Remain.BackColor = unselectedColor And _
            Standard_Certification_Update.BackColor = unselectedColor Then
        MsgBox "Please select a result button for Certification"
        Exit Sub
    End If

    If Standard_Admission_Remain.BackColor = unselectedColor And _
            Standard_Admission_Update.BackColor = unselectedColor Then
        MsgBox "Please select a result button for Admission"
        Exit Sub
    End If

    If Standard_Adjudication_Remain.BackColor = unselectedColor And _
            Standard_Adjudication_Update.BackColor = unselectedColor Then
        MsgBox "Please select a result button for Adjudication"
        Exit Sub
    End If

    If Standard_Continuance_Remain.BackColor = unselectedColor And _
            Standard_Continuance_Update.BackColor = unselectedColor Then
        MsgBox "Please select a result button for Continuance"
        Exit Sub
    End If


    If Not HasContent(Standard_NextCourtDate) Then
        MsgBox "Please enter the next court date"
        Exit Sub
    End If

    Dim oldCourtHead As String
    Dim oldCourtroom As String
    Dim newCourtHead As String
    Dim newCourtroom As String
    oldCourtroom = Standard_Title.Caption
    newCourtroom = Standard_Title.Caption

    If Standard_Court_Transfer.BackColor = selectedColor Then
        newCourtroom = Modal_Standard_Court_Transfer.Courtroom.value
    End If

    If oldCourtroom = "5E" Then
        oldCourtHead = headerFind("Crossover")
    Else
        oldCourtHead = headerFind(oldCourtroom)
    End If

    If newCourtHead = "5E" Then
        newCourtHead = headerFind("Crossover")
    Else
        newCourtHead = headerFind(newCourtroom)
    End If

    '''''''''''''''''
    ''''DEMOGRAPHICS'
    '''''''''''''''''
    'append PCD
    Call append(Range(headerFind("Previous Court Dates") & updateRow), DateOfHearing.value)

    Range(headerFind("Next Court Date") & updateRow) = Standard_NextCourtDate.value
    Range(headerFind("Listing Type") & updateRow) = Lookup("Listing_Type_Name")(Standard_ListingType.value)

    Range(headerFind("Legal Status") & updateRow) _
                = Lookup("Legal_Status_Name")(Standard_Return_Legal_Status.Caption)

   

    Dim oldLegalHead As String
    Dim newLegalHead As String
    Dim bucketHead As String
    Dim i As Long



    ''''''''''''''
    'LEGAL STATUS'
    ''''''''''''''
    If Standard_Legal_Status_Update.BackColor = selectedColor Then
        If Standard_Court_Transfer.BackColor = unselectedColor Then
            With Modal_Standard_Legal_Status
                Call legalStatusEnd( _
                    clientRow:=updateRow, _
                    statusType:=.Current_Legal_Status, _
                    Courtroom:=oldCourtroom, _
                    DA:=DA.value, _
                    endDate:=.Current_Discharge_Date, _
                    Nature:=.Current_Discharge_Nature, _
                    withAgg:=True, _
                    detailed:=.Current_Detailed_Outcome, _
                    Reason1:=.Reason1, Reason2:=.Reason2, Reason3:=.Reason3, Reason4:=.Reason4, Reason5:=.Reason5, _
                    Notes:=.Current_Notes)
    
                If isTerminal("Legal Status", .Current_Detailed_Outcome) Then
                    Call totalOutcome( _
                        clientRow:=updateRow, _
                        DateOf:=.Current_Discharge_Date, _
                        Courtroom:=oldCourtroom, _
                        DA:=DA.value, _
                        legalStatus:=.Current_Legal_Status.Caption, _
                        Nature:=.Courtroom_Outcome_Nature, _
                        detailed:=.Courtroom_Detailed_Outcome, _
                        Notes:=Standard_Notes.value)
                Else
                    Call legalStatusStart( _
                        clientRow:=updateRow, _
                        statusType:=.New_Legal_Status, _
                        Courtroom:=newCourtroom, _
                        DA:=DA.value, _
                        startDate:=.New_Start_Date, _
                        Notes:=.New_Notes)
                End If
            End With
        Else
            Call legalStatusStart( _
                clientRow:=updateRow, _
                statusType:=Standard_Return_Legal_Status.Caption, _
                Courtroom:=oldCourtroom, _
                DA:=DA.value, _
                startDate:=DateOfHearing.value, _
                Notes:=Standard_Notes.value, _
                zeroLocal:=True)
        End If
    Else
        Call legalStatusStart( _
            clientRow:=updateRow, _
            statusType:=Standard_Return_Legal_Status.Caption, _
            Courtroom:=oldCourtroom, _
            DA:=DA.value, _
            startDate:=DateOfHearing.value, _
            Notes:="Transferred from prior CR")
    End If



    ''''''''''''''''
    'FAIL TO APPEAR'
    ''''''''''''''''

    If Standard_FTA_Yes.BackColor = selectedColor Then
        Call flagYes(Range(hFind("Did Youth FTA?", "AGGREGATES") & updateRow))


        For i = 1 To 15
            If isEmptyOrZero(Range(hFind("FTA #" & i & " Date", "AGGREGATES") & updateRow)) Then

                bucketHead = hFind("FTA #" & i & " Date", "AGGREGATES")

                Range(bucketHead & updateRow).value = DateOfHearing.value
                Range(headerFind("Day of FTA", bucketHead) & updateRow).value _
                        = Weekday(DateOfHearing.value, vbMonday) * 2 - 1
                Range(headerFind("Courtroom", bucketHead) & updateRow).value = Range(hFind("Active Courtroom") & updateRow).value
                Range(headerFind("Legal Status", bucketHead) & updateRow).value = Range(hFind("Legal Status") & updateRow).value

                If i = 1 Then
                    Range(headerFind("LOS to FTA", bucketHead) & updateRow).value _
                            = calcLOS(Range(hFind("Arrest Date") & updateRow).value, DateOfHearing.value)
                Else
                    Range(headerFind("LOS Between FTAs", bucketHead) & updateRow).value _
                            = calcLOS(Range(hFind("FTA #" & (i - 1) & " Date", "AGGREGATES") & updateRow).value, DateOfHearing.value)
                End If


                If Range(hFind("Active B/W?") & updateRow).value = Lookup("Generic_YNOU_Name")("Yes") Then
                    Range(headerFind("B/W Action", bucketHead) & updateRow).value = Lookup("BW_Action_Name")("Continue B/W")
                End If

                If Range(hFind("Active B/W?") & updateRow).value = Lookup("Generic_YNOU_Name")("No") Then
                    If Modal_FTA.BW.value = "Yes" Then
                        Range(headerFind("B/W Action", bucketHead) & updateRow).value = Lookup("BW_Action_Name")("Begin B/W")
                        Call flagYes(Range(hFind("Active B/W?") & updateRow))
                    Else
                        Range(headerFind("B/W Action", bucketHead) & updateRow).value = Lookup("BW_Action_Name")("N/A")
                    End If
                End If

                i = 15
            End If
        Next i
    End If


    If Standard_Lift_BW.BackColor = selectedColor Then
        Range(hFind("Active B/W?") & updateRow).value = Lookup("Generic_YNOU_Name")("No")

        For i = 15 To 1 Step -1
            If isNotEmptyOrZero(Range(hFind("FTA #" & i & " Date", "AGGREGATES") & updateRow)) _
                    And Range(hFind("B/W Action", "FTA #" & i & " Date", "AGGREGATES") & updateRow).value _
                        = Lookup("BW_Action_Name")("Begin B/W") Then

                bucketHead = hFind("FTA #" & i & " Date", "AGGREGATES")

                Range(headerFind("B/W Lifted Date", bucketHead) & updateRow).value _
                        = Modal_Lift_BW.DateBox.value

                Range(headerFind("LOS B/W", bucketHead) & updateRow).value _
                            = calcLOS(Range(bucketHead & updateRow).value, Modal_Lift_BW.DateBox.value)
                i = 1
            End If
        Next i
    End If
    '''''''''''''''
    'CERTIFICATION'
    '''''''''''''''

    If Standard_Certification_Update.BackColor = selectedColor Then
        If Standard_Fetch_Certification.Caption = "Filed" Then
            Call certificationUpdate( _
                updateRow, _
                headerFind("Certification", oldCourtHead), _
                Modal_Standard_Certification.Motion_Result, _
                DateOfHearing.value _
            )
            Call certificationUpdate( _
                updateRow, _
                hFind("Certification", "COURT PROCEEDINGS", "AGGREGATES"), _
                Modal_Standard_Certification.Motion_Result, _
                DateOfHearing.value _
            )
        Else
            Call certificationStart( _
                updateRow, _
                headerFind("Certification", newCourtHead), _
                Lookup("Legal_Status_Num")(Range(headerFind("Legal Status") & updateRow).value), _
                newCourtroom, _
                DA.value, _
                Modal_Standard_Certification.Motion_Date.value _
            )
            Call certificationStart( _
                updateRow, _
                hFind("Certification", "COURT PROCEEDINGS", "AGGREGATES"), _
                Lookup("Legal_Status_Num")(Range(headerFind("Legal Status") & updateRow).value), _
                newCourtroom, _
                DA.value, _
                Modal_Standard_Certification.Motion_Date.value _
            )
        End If
    End If

    '''''''''''
    'ADMISSION'
    '''''''''''
    If Standard_Admission_Update.BackColor = selectedColor Then
        With Modal_Standard_Admission
            Call admissionStart( _
                clientRow:=updateRow, _
                petitionNum:=Modal_Standard_Admission.PetitionBox.value, _
                statusType:=Lookup("Legal_Status_Num")(Range(headerFind("Legal Status") & updateRow).value), _
                Courtroom:=newCourtroom, _
                DA:=DA.value, _
                startDate:=Modal_Standard_Admission.Admission_Date.value, _
                Result:=Modal_Standard_Admission.Result.value, _
                detailed:=Modal_Standard_Admission.Detailed_Result.value, _
                leadChargeCode:=.PetitionBox.List(.PetitionBox.listIndex, 4), _
                leadChargeName:=.PetitionBox.List(.PetitionBox.listIndex, 5), _
                chargeCategory:=.PetitionBox.List(.PetitionBox.listIndex, 3), _
                chargeGradeSpecific:=.PetitionBox.List(.PetitionBox.listIndex, 2) _
            )
        End With
    End If

    ''''''''''''''
    'Adjudication'
    ''''''''''''''
    If Standard_Adjudication_Update.BackColor = selectedColor Then
        With Modal_Standard_Adjudication
            Call adjudicationStart( _
                clientRow:=updateRow, _
                petitionNum:=.PetitionBox.value, _
                Courtroom:=newCourtroom, _
                DA:=DA.value, _
                startDate:=.Adjudication_Date.value, _
                Type_of:=.Type_of.value, _
                re1:=.Reason1.value, _
                re2:=.Reason2.value, _
                re3:=.Reason3.value, _
                re4:=.Reason4.value, _
                re5:=.Reason5.value, _
                leadChargeCode:=.PetitionBox.List(.PetitionBox.listIndex, 4), _
                leadChargeName:=.PetitionBox.List(.PetitionBox.listIndex, 5), _
                chargeCategory:=.PetitionBox.List(.PetitionBox.listIndex, 3), _
                chargeGradeSpecific:=.PetitionBox.List(.PetitionBox.listIndex, 2) _
            )
        End With
    End If
    
    '''''''''''''
    'CONTINUANCE'
    '''''''''''''
    If Standard_Continuance_Update.BackColor = selectedColor Then
        Call continuanceStart( _
            updateRow, _
            Modal_Standard_Continuance.Status, _
            newCourtroom, _
            DA.value, _
            DateOfHearing.value, _
            Standard_NextCourtDate.value, _
            Modal_Standard_Continuance.Continuance_Type.value, _
            Modal_Standard_Continuance.Reason1.value, Modal_Standard_Continuance.Reason2.value, Modal_Standard_Continuance.Reason3.value _
        )
    End If



    With Standard_Return_Supervision_Box
        Dim j As Long
        Dim service As String

        'we will use this for loop to iterate through all of the rows in the listbox
        '
        For i = 0 To .ListCount - 1
            If .List(i, 4) = "New" Then
                Call addSupervision( _
                    clientRow:=updateRow, _
                    serviceType:=.List(i, 0), _
                    legalStatus:=Standard_Return_Legal_Status.Caption, _
                    Courtroom:=oldCourtroom, _
                    DA:=DA.value, _
                    agency:=.List(i, 1), _
                    startDate:=.List(i, 2), _
                    NextCourtDate:=Standard_NextCourtDate.value, _
                    re1:=decodeReasons(.List(i, 6))(0), _
                    re2:=decodeReasons(.List(i, 6))(1), _
                    re3:=decodeReasons(.List(i, 6))(2), _
                    re4:=decodeReasons(.List(i, 6))(3), _
                    re5:=decodeReasons(.List(i, 6))(4), _
                    Notes:=.List(i, 9))
            Else
                If .List(i, 4) = oldCourtroom Then
                'if service ordered from this courtroom
                    
                    If IsDate(.List(i, 3)) Then
                    'if has End Date
                    
                        Call dropSupervision( _
                            clientRow:=updateRow, _
                            Courtroom:=.List(i, 4), _
                            serviceType:=.List(i, 0), _
                            startDate:=.List(i, 2), _
                            endDate:=.List(i, 3), _
                            Nature:=.List(i, 5), _
                            re1:=decodeReasons(.List(i, 6))(0), _
                            re2:=decodeReasons(.List(i, 6))(1), _
                            re3:=decodeReasons(.List(i, 6))(2), _
                            re4:=decodeReasons(.List(i, 6))(3), _
                            re5:=decodeReasons(.List(i, 6))(4), _
                            Notes:=.List(i, 9))
                    End If
                Else
                ' if service ordered from somewhere else
                
                    If IsDate(.List(i, 3)) Then
                    'if has end date
                    
                        Call dropSupervision( _
                            clientRow:=updateRow, _
                            Courtroom:=.List(i, 4), _
                            serviceType:=.List(i, 0), _
                            startDate:=.List(i, 2), _
                            endDate:=.List(i, 3), _
                            Nature:=.List(i, 5), _
                            re1:=decodeReasons(.List(i, 6))(0), _
                            re2:=decodeReasons(.List(i, 6))(1), _
                            re3:=decodeReasons(.List(i, 6))(2), _
                            re4:=decodeReasons(.List(i, 6))(3), _
                            re5:=decodeReasons(.List(i, 6))(4), _
                            Notes:=.List(i, 9))
                    Else
                    'if continued from somewhere else
                    
                        Call dropSupervision( _
                            clientRow:=updateRow, _
                            Courtroom:=.List(i, 4), _
                            serviceType:=.List(i, 0), _
                            startDate:=.List(i, 2), _
                            endDate:=DateOfHearing.value, _
                            Nature:="Neutral", _
                            re1:="N/A", _
                            re2:="N/A", _
                            re3:="N/A", _
                            Notes:="Continued in " & oldCourtroom)
    
                        Call addSupervision( _
                            clientRow:=updateRow, _
                            serviceType:=.List(i, 0), _
                            legalStatus:=Standard_Return_Legal_Status.Caption, _
                            Courtroom:=oldCourtroom, _
                            DA:=DA.value, _
                            agency:=.List(i, 1), _
                            startDate:=DateOfHearing.value, _
                            NextCourtDate:=Standard_NextCourtDate.value, _
                            re1:="N/A", _
                            re2:="N/A", _
                            re3:="N/A", _
                            Notes:="Continued from " & .List(i, 4))
                    End If
                End If
            End If
        Next i
    End With

    With Standard_Return_Condition_Box
        Dim condition As String

        'we will use this for loop to iterate through all of the rows in the listbox
        '
        For i = 0 To .ListCount - 1
            If .List(i, 4) = "New" Then
                Call addCondition( _
                    clientRow:=updateRow, _
                    condition:=.List(i, 0), _
                    legalStatus:=Standard_Return_Legal_Status.Caption, _
                    Courtroom:=oldCourtroom, _
                    DA:=DA.value, _
                    agency:=.List(i, 1), _
                    startDate:=.List(i, 2), _
                    re1:=decodeReasons(.List(i, 6))(0), _
                    re2:=decodeReasons(.List(i, 6))(1), _
                    re3:=decodeReasons(.List(i, 6))(2), _
                    re4:=decodeReasons(.List(i, 6))(3), _
                    re5:=decodeReasons(.List(i, 6))(4), _
                    Notes:=.List(i, 9))
                    
                Select Case .List(i, 0)
                    Case "Restitution"
                        Call startRestitution( _
                            Amount:=ClientUpdateForm.Standard_Restitution.Caption, _
                            Courtroom:=oldCourtroom, _
                            DA:=DA.value, _
                            DateOf:=.List(i, 2), _
                            userRow:=updateRow)
                    
                    Case "Comm. Serv"
                        Call startCommService( _
                            Amount:=ClientUpdateForm.Standard_Comm_Service.Caption, _
                            Courtroom:=oldCourtroom, _
                            DA:=DA.value, _
                            DateOf:=.List(i, 2), _
                            userRow:=updateRow)
                            
                    Case "Court Costs"
                        Call startCourtCosts( _
                            Amount:=ClientUpdateForm.Standard_Court_Costs.Caption, _
                            Courtroom:=oldCourtroom, _
                            DA:=DA.value, _
                            DateOf:=.List(i, 2), _
                            userRow:=updateRow)
                End Select
            Else
                If .List(i, 4) = oldCourtroom Then
                'if service ordered from this courtroom
                    
                    If IsDate(.List(i, 3)) Then
                    'if has End Date
                    
                        Call dropCondition( _
                            clientRow:=updateRow, _
                            Courtroom:=.List(i, 4), _
                            condition:=.List(i, 0), _
                            startDate:=.List(i, 2), _
                            endDate:=.List(i, 3), _
                            Nature:=.List(i, 5), _
                            re1:=decodeReasons(.List(i, 6))(0), _
                            re2:=decodeReasons(.List(i, 6))(1), _
                            re3:=decodeReasons(.List(i, 6))(2), _
                            re4:=decodeReasons(.List(i, 6))(3), _
                            re5:=decodeReasons(.List(i, 6))(4), _
                            Notes:=.List(i, 9))
                    End If
                Else
                ' if service ordered from somewhere else
                
                    If IsDate(.List(i, 3)) Then
                    'if has end date
                    
                        Call dropCondition( _
                            clientRow:=updateRow, _
                            Courtroom:=.List(i, 4), _
                            condition:=.List(i, 0), _
                            startDate:=.List(i, 2), _
                            endDate:=.List(i, 3), _
                            Nature:=.List(i, 5), _
                            re1:=decodeReasons(.List(i, 6))(0), _
                            re2:=decodeReasons(.List(i, 6))(1), _
                            re3:=decodeReasons(.List(i, 6))(2), _
                            re4:=decodeReasons(.List(i, 6))(3), _
                            re5:=decodeReasons(.List(i, 6))(4), _
                            Notes:=.List(i, 9))
                    Else
                    'if continued from somewhere else
                    
                        Call dropCondition( _
                            clientRow:=updateRow, _
                            Courtroom:=.List(i, 4), _
                            condition:=.List(i, 0), _
                            startDate:=.List(i, 2), _
                            endDate:=DateOfHearing.value, _
                            Nature:="Neutral", _
                            re1:="N/A", _
                            re2:="N/A", _
                            re3:="N/A", _
                            Notes:="Continued in " & oldCourtroom)
    
                        Call addCondition( _
                            clientRow:=updateRow, _
                            condition:=.List(i, 0), _
                            legalStatus:=Standard_Return_Legal_Status.Caption, _
                            Courtroom:=oldCourtroom, _
                            DA:=DA.value, _
                            agency:=.List(i, 1), _
                            startDate:=DateOfHearing.value, _
                            re1:="N/A", _
                            re2:="N/A", _
                            re3:="N/A", _
                            Notes:="Continued from " & .List(i, 4))
                    End If
                End If
            End If
        Next i
    End With

    If Standard_Court_Transfer.BackColor = selectedColor Then
        Dim outcomeHead As String
        outcomeHead = headerFind("OUTCOMES", oldCourtHead)

        Range(headerFind("Notes on Outcome", outcomeHead) & updateRow).value = "Transferred out of courtroom"
        Range(headerFind("Date of Overall Discharge", outcomeHead) & updateRow).value = DateOfHearing.value
        Range(headerFind("Courtroom of Discharge", outcomeHead) & updateRow).value = Lookup("Courtroom_Name")(oldCourtroom)
        Range(headerFind("Legal Status of Discharge", outcomeHead) & updateRow).value = Lookup("Legal_Status_Name")(Standard_Return_Legal_Status.Caption)
        Range(headerFind("DA", outcomeHead) & updateRow).value = Lookup("DA_Last_Name_Name")(DA.value)
        Range(headerFind("Active or Discharged", outcomeHead) & updateRow).value = 2 'discharged
        Range(headerFind("Nature of Courtroom Outcome", outcomeHead) & updateRow).value _
                = Lookup("Nature_of_Discharge_Name")(NatureFromDetailed(Modal_Standard_Court_Transfer.Detailed_Outcome.value))
        Range(headerFind("Detailed Courtroom Outcome", outcomeHead) & updateRow).value _
                = Lookup("Detailed_Courtroom_Outcome_Name")(Modal_Standard_Court_Transfer.Detailed_Outcome.value)

        If Standard_Legal_Status_Update.BackColor = selectedColor Then
            Call ReferClientTo( _
                referralDate:=DateOfHearing.value, _
                clientRow:=updateRow, _
                newLegalStatus:=Standard_Return_Legal_Status.Caption, _
                oldLegalStatus:=Standard_Fetch_Legal_Status.Caption, _
                toCR:=Modal_Standard_Court_Transfer.Courtroom.value, _
                fromCR:=oldCourtroom, _
                DA:=DA.value)

        Else
            Call ReferClientTo( _
                referralDate:=DateOfHearing.value, _
                clientRow:=updateRow, _
                toCR:=Modal_Standard_Court_Transfer.Courtroom.value, _
                fromCR:=oldCourtroom)
        End If
    End If

    Call closeCallIn(DateOfHearing.value, updateRow)
    Call closeIntakeConference(DateOfHearing.value, updateRow)
    Call UnloadAll

    Call addNotes( _
        Courtroom:=oldCourtroom, _
        DateOf:=DateOfHearing.value, _
        userRow:=updateRow, _
        Notes:=Standard_Notes, _
        DA:=DA.value _
    )


    Call CheckForConcurrency(updateRow, DateOfHearing.value)

    Call formSubmitEnd

done:
    Exit Sub
err:
    
    Call loadFromCache(2)
    
    MsgBox "Something went wrong. Database has been restored to state prior to submission. " _
      & vbNewLine & vbNewLine & "Message: " & vbNewLine & err.Description _
      & vbNewLine & vbNewLine & "Source: " & vbNewLine & err.Source
    Call UnloadAll

    'Stop   'press F8 twice to see the error point
    'Resume
End Sub



Private Sub PJJSC_Submit_Click()
    Dim Num As Integer
    Dim bucketHead As String
    Dim detentionHead As String

    Call formSubmitStart(updateRow)
    
    Call addNotes( _
        Courtroom:="PJJSC", _
        DateOf:=DateOfHearing.value, _
        userRow:=updateRow, _
        Notes:=PJJSC_NotesOnDetentionOutcome.value, _
        DA:=DA.value _
    )

    detentionHead = headerFind("DETENTION")

    'find empty review hearing bucket to enter data
    For Num = 1 To 10
        bucketHead = hFind("Date of Review #" & Num, "DETENTION")
        If isEmptyOrZero(Range(bucketHead & updateRow)) Then
            Num = 10
        End If
    Next Num

    If Range(hFind("Did Youth Have Initial Detention Hearing?", "DETENTION") & updateRow).value _
        = Lookup("Generic_YN_Name")("Yes") Then

    Else
        bucketHead = hFind("Date of Initial Detention Hearing", "DETENTION")
        Call flagYes(Range(hFind("Did Youth Have Initial Detention Hearing?", "DETENTION") & updateRow))
        Range(hFind("Type of Detention Hearing", "DETENTION") & updateRow).value _
            = Lookup("Type_of_Detention_Hearing_Name")("Initial")
        If DRev_DetentionDecision.value = "Held" Then
            Range(hFind("Reason #1 for Detention Commit", "DETENTION") & updateRow).value _
                = Lookup("Detention_Hearing_Reason_Name")(ReasonForDetentionCommit1.value)
            Range(hFind("Reason #2 for Detention Commit", "DETENTION") & updateRow).value _
                = Lookup("Detention_Hearing_Reason_Name")(ReasonForDetentionCommit2.value)
            Range(hFind("Reason #3 for Detention Commit", "DETENTION") & updateRow).value _
                = Lookup("Detention_Hearing_Reason_Name")(ReasonForDetentionCommit3.value)
            Range(hFind("Reason #4 for Detention Commit", "DETENTION") & updateRow).value _
                = Lookup("Detention_Hearing_Reason_Name")(ReasonForDetentionCommit4.value)
            Range(hFind("Reason #5 for Detention Commit", "DETENTION") & updateRow).value _
                = Lookup("Detention_Hearing_Reason_Name")(ReasonForDetentionCommit5.value)
        End If
    End If

    Range(bucketHead & updateRow).value = DateOfHearing
    Range(headerFind("DA", bucketHead) & updateRow).value = Lookup("DA_Last_Name_Name")(DA.value)
    Range(headerFind("DA Action", bucketHead) & updateRow).value = Lookup("DA_Action_Name")(DRev_DA_Action.value)
    Range(headerFind("DA Action Accepted?", bucketHead) & updateRow).value = Lookup("Generic_YNOU_Name")(DRev_ActionAccepted.value)
    Range(headerFind("Detention Decision", bucketHead) & updateRow).value = Lookup("Detention_Decision_Name")(DRev_DetentionDecision.value)
    Range(headerFind("Detention Facility", bucketHead) & updateRow).value = Lookup("Detention_Facility_Name")(DRev_Facility.value)
    Range(headerFind("Notes on Detention", bucketHead) & updateRow).value = PJJSC_NotesOnDetentionOutcome.value


    If DRev_DetentionDecision.value = "Held" _
        And DRev_HoldAndTransfer.value = "Yes" Then
        Call ReferClientTo( _
            referralDate:=DateOfHearing.value, _
            clientRow:=updateRow, _
            fromCR:="PJJSC", _
            toCR:=DRev_NextHearingLocation.value, _
            DA:=DA.value)

        Call addSupervision( _
            clientRow:=updateRow, _
            serviceType:="Detention (not respite)", _
            legalStatus:=Lookup("Legal_Status_Num")(Range(hFind("Legal Status") & updateRow).value), _
            Courtroom:="PJJSC", _
            DA:=DA.value, _
            agency:=DRev_Facility.value, _
            startDate:=DateOfHearing.value, _
            NextCourtDate:=PJJSC_NextCourtDate.value, _
            re1:=ReasonForDetentionCommit1.value, _
            re2:=ReasonForDetentionCommit2.value, _
            re3:=ReasonForDetentionCommit3.value)
    End If

    'IF RELEASED
    If DRev_DetentionDecision.value = "Released" Then
        Range(headerFind("Date of Release", bucketHead) & updateRow).value = DateOfHearing.value


        'Range(headerFind("Courtroom That Released", bucketHead) & updateRow).value =
        Range(headerFind("Referred to Courtroom", bucketHead) & updateRow).value _
            = Lookup("Courtroom_Name")(DRev_ReferredTo.value)

        'REFER TO COURTROOM
        Call ReferClientTo( _
            referralDate:=DateOfHearing.value, _
            clientRow:=updateRow, _
            fromCR:="PJJSC", _
            toCR:=DRev_ReferredTo.value, _
            DA:=DA.value _
        )

        'ADD SUPERVISION #1 TO DETENTION SECTION
        bucketHead = hFind("Supervision Ordered #1", "DETENTION")
        Range(bucketHead & updateRow).value = Lookup("Supervision_Program_Name")(DRevSup1.value)
        If isResidential(DRevSup1.value) Then
            Range(headerFind("Residential Agency", bucketHead) & updateRow).value _
                = Lookup("Residential_Supervision_Provider_Name")(DRevSup1_Agency.value)
        Else
            Range(headerFind("Community-Based Agency", bucketHead) & updateRow).value _
                = Lookup("Community_Based_Supervision_Provider_Name")(DRevSup1_Agency.value)
        End If

        Range(headerFind("Reason #1 for Supervision Referral", bucketHead) & updateRow).value = Lookup("Supervision_Referral_Reason_Name")(DRevSup1_Re1.value)
        Range(headerFind("Reason #2 for Supervision Referral", bucketHead) & updateRow).value = Lookup("Supervision_Referral_Reason_Name")(DRevSup1_Re2.value)
        Range(headerFind("Reason #3 for Supervision Referral", bucketHead) & updateRow).value = Lookup("Supervision_Referral_Reason_Name")(DRevSup1_Re3.value)
        Range(headerFind("Reason #4 for Supervision Referral", bucketHead) & updateRow).value = Lookup("Supervision_Referral_Reason_Name")(DRevSup1_Re4.value)
        Range(headerFind("Reason #5 for Supervision Referral", bucketHead) & updateRow).value = Lookup("Supervision_Referral_Reason_Name")(DRevSup1_Re5.value)


        'ADD SUPERVISION #1 TO NEW COURTROOM (AND AGG)
        If Not DRevSup1.value = "None" Then
            Call addSupervision( _
                clientRow:=updateRow, _
                serviceType:=DRevSup1.value, _
                legalStatus:=Lookup("Legal_Status_Num")(Range(hFind("Legal Status") & updateRow).value), _
                Courtroom:="PJJSC", _
                DA:=DA.value, _
                agency:=DRevSup1_Agency.value, _
                startDate:=DateOfHearing.value, _
                NextCourtDate:=PJJSC_NextCourtDate.value, _
                re1:=DRevSup1_Re1.value, _
                re2:=DRevSup1_Re2.value, _
                re3:=DRevSup1_Re3.value, _
                Notes:="Referred at detention")
        End If


        'ADD SUPERVISION #2 TO DETENTION SECTION
        bucketHead = hFind("Supervision Ordered #2", "DETENTION")
        Range(bucketHead & updateRow).value = Lookup("Supervision_Program_Name")(DRevSup2.value)
        If isResidential(DRevSup2.value) Then
            Range(headerFind("Residential Agency", bucketHead) & updateRow).value _
                = Lookup("Residential_Supervision_Provider_Name")(DRevSup1_Agency.value)
        Else
            Range(headerFind("Community-Based Agency", bucketHead) & updateRow).value _
                = Lookup("Community_Based_Supervision_Provider_Name")(DRevSup1_Agency.value)
        End If

        Range(headerFind("Reason #1 for Supervision Referral", bucketHead) & updateRow).value = Lookup("Supervision_Referral_Reason_Name")(DRevSup2_Re1.value)
        Range(headerFind("Reason #2 for Supervision Referral", bucketHead) & updateRow).value = Lookup("Supervision_Referral_Reason_Name")(DRevSup2_Re2.value)
        Range(headerFind("Reason #3 for Supervision Referral", bucketHead) & updateRow).value = Lookup("Supervision_Referral_Reason_Name")(DRevSup2_Re3.value)
        Range(headerFind("Reason #4 for Supervision Referral", bucketHead) & updateRow).value = Lookup("Supervision_Referral_Reason_Name")(DRevSup2_Re4.value)
        Range(headerFind("Reason #5 for Supervision Referral", bucketHead) & updateRow).value = Lookup("Supervision_Referral_Reason_Name")(DRevSup2_Re5.value)


        'ADD SUPERVISION #2 TO NEW COURTROOM (AND AGG)
        If Not DRevSup2.value = "None" Then
            Call addSupervision( _
                clientRow:=updateRow, _
                serviceType:=DRevSup2.value, _
                legalStatus:=Lookup("Legal_Status_Num")(Range(hFind("Legal Status") & updateRow).value), _
                Courtroom:="PJJSC", _
                DA:=DA.value, _
                agency:=DRevSup2_Agency.value, _
                startDate:=DateOfHearing.value, _
                NextCourtDate:=PJJSC_NextCourtDate.value, _
                re1:=DRevSup2_Re1.value, _
                re2:=DRevSup2_Re2.value, _
                re3:=DRevSup2_Re3.value, _
                Notes:="Referred at detention")
        End If

        'CONDITION #1
        Range(headerFind("Other Condition #1", bucketHead) & updateRow).value = Lookup("Condition_Name")(DRev_C1.value)
        Range(headerFind("Other Condition #1 Provider", bucketHead) & updateRow).value = Lookup("Condition_Provider_Name")(DRev_C1P.value)
        If Not DRev_C1.value = "None" Then
            Call addCondition( _
                clientRow:=updateRow, _
                condition:=DRev_C1.value, _
                legalStatus:=Lookup("Legal_Status_Num")(Range(hFind("Legal Status") & updateRow).value), _
                Courtroom:="PJJSC", _
                DA:=DA.value, _
                agency:=DRev_C1P.value, _
                startDate:=DateOfHearing.value, _
                re1:="N/A", _
                re2:="N/A", _
                re3:="N/A", _
                Notes:="Referred at detention")
        End If

        'CONDITION #2
        Range(headerFind("Other Condition #2", bucketHead) & updateRow).value = Lookup("Condition_Name")(DRev_C2.value)
        Range(headerFind("Other Condition #2 Provider", bucketHead) & updateRow).value = Lookup("Condition_Provider_Name")(DRev_C2P.value)
        If Not DRev_C2.value = "None" Then
            Call addCondition( _
                clientRow:=updateRow, _
                condition:=DRev_C2.value, _
                legalStatus:=Lookup("Legal_Status_Num")(Range(hFind("Legal Status") & updateRow).value), _
                Courtroom:="PJJSC", _
                DA:=DA.value, _
                agency:=DRev_C2P.value, _
                startDate:=DateOfHearing.value, _
                re1:="N/A", _
                re2:="N/A", _
                re3:="N/A", _
                Notes:="Referred at detention")
        End If

        'CONDITION #3
        Range(headerFind("Other Condition #3", bucketHead) & updateRow).value = Lookup("Condition_Name")(DRev_C3.value)
        Range(headerFind("Other Condition #3 Provider", bucketHead) & updateRow).value = Lookup("Condition_Provider_Name")(DRev_C3P.value)
        If Not DRev_C3.value = "None" Then
            Call addCondition( _
                clientRow:=updateRow, _
                condition:=DRev_C3.value, _
                legalStatus:=Lookup("Legal_Status_Num")(Range(hFind("Legal Status") & updateRow).value), _
                Courtroom:="PJJSC", _
                DA:=DA.value, _
                agency:=DRev_C3P.value, _
                startDate:=DateOfHearing.value, _
                re1:="N/A", _
                re2:="N/A", _
                re3:="N/A", _
                Notes:="Referred at detention")
        End If


        Range(headerFind("Notes on Detention", bucketHead) & updateRow).value = PJJSC_NotesOnDetentionOutcome.value
        Range(headerFind("LOS in Detention", bucketHead) & updateRow).value _
            = calcLOS(Range(hFind("Date of Initial Detention Hearing", "DETENTION") & updateRow).value, DateOfHearing.value)
        Range(headerFind("LOS from Arrest Until Hearing", bucketHead) & updateRow).value _
            = calcLOS(Range(hFind("Arrest Date") & updateRow).value, Range(hFind("Date of Initial Detention Hearing", "DETENTION") & updateRow).value)
    End If


    'Update Next Court Date in Front of Database
    Range(headerFind("Next Court Date") & updateRow) = PJJSC_NextCourtDate.value



    Call closeCallIn(DateOfHearing.value, updateRow)
    Call closeIntakeConference(DateOfHearing.value, updateRow)
    Call closeIntakeDetentions(DateOfHearing.value, updateRow)

    Call UnloadAll
    
    Call formSubmitEnd

End Sub

''''''''''''''''''''''''''''
''''''''''BUTTONS'''''''''''
''''''''''''''''''''''''''''

'''''''''''''''''''
'''ADULT_UPDATES'''
'''''''''''''''''''

Sub Adult_Legal_Status_Remain_Click()
    Call toggleSelect(Adult_Legal_Status_Remain, Adult_Return_Legal_Status, Adult_Fetch_Legal_Status)
    Adult_Legal_Status_Update.BackColor = unselectedColor
End Sub
Sub Adult_Legal_Status_Update_Click()
    'Modal_Adult_Legal_Status.Show
End Sub


Sub Adult_Reslate_Remain_Click()
    Call toggleSelect(Adult_Reslate_Remain, Adult_Return_Reslate, Adult_Fetch_Reslate)
    Adult_Reslate_Update.BackColor = unselectedColor
End Sub
Sub Adult_Reslate_Update_Click()
    Modal_Adult_Reslate.Show
End Sub


Sub Adult_Decertification_Remain_Click()
    Call toggleSelect(Adult_Decertification_Remain, Adult_Return_Decertification, Adult_Fetch_Decertification)
    Adult_Decertification_Update.BackColor = unselectedColor
End Sub
Sub Adult_Decertification_Update_Click()
    Modal_Adult_Decertification.Show
End Sub


Sub Adult_Admission_Remain_Click()
    Call toggleSelect(Adult_Admission_Remain, Adult_Return_Admission, Adult_Fetch_Admission)
    Adult_Admission_Update.BackColor = unselectedColor
End Sub
Sub Adult_Admission_Update_Click()
    Modal_Adult_Admission.Show
End Sub


Sub Adult_Adjudication_Remain_Click()
    Call toggleSelect(Adult_Adjudication_Remain, Adult_Return_Adjudication, Adult_Fetch_Adjudication)
    Adult_Adjudication_Update.BackColor = unselectedColor
End Sub
Sub Adult_Adjudication_Update_Click()
    'Modal_Adult_Adjudication.Show
End Sub

Sub Adult_Continuance_Remain_Click()
    Call toggleSelect(Adult_Continuance_Remain, Adult_Return_Continuance, "No")
    Adult_Continuance_Update.BackColor = unselectedColor
End Sub
Sub Adult_Continuance_Update_Click()
    'Modal_Adult_Continuance.Show
End Sub

Sub Adult_Supervision_Add_Click()
    'Modal_Adult_Add_Supervision.Show
End Sub
Sub Adult_Supervision_Discharge_Click()
    'Modal_Adult_Drop_Supervision.Show
End Sub
Sub Adult_Supervision_Remain_Click()
    Call toggleSelect(Adult_Supervision_Remain)
End Sub

Sub Adult_Condition_Add_Click()
    'Modal_Adult_Add_Condition.Show
End Sub
Sub Adult_Condition_Discharge_Click()
    'Modal_Adult_Drop_Condition.Show
End Sub
Sub Adult_Condition_Remain_Click()
    Call toggleSelect(Adult_Condition_Remain)
End Sub

Sub Adult_Remain_All_Click()
    If Not Adult_Legal_Status_Update.BackColor = selectedColor Then
        Call Adult_Legal_Status_Remain_Click
    End If

    If Not Adult_Reslate_Update.BackColor = selectedColor Then
        Call Adult_Reslate_Remain_Click
    End If

    If Not Adult_Decertification_Update.BackColor = selectedColor Then
        Call Adult_Decertification_Remain_Click
    End If

    If Not Adult_Admission_Update.BackColor = selectedColor Then
        Call Adult_Admission_Remain_Click
    End If

    If Not Adult_Adjudication_Update.BackColor = selectedColor Then
        Call Adult_Adjudication_Remain_Click
    End If

    If Not Adult_Continuance_Update.BackColor = selectedColor Then
        Call Adult_Continuance_Remain_Click
    End If

End Sub

'''''''''''''''''''
'STANDARD_UPDATES''
'''''''''''''''''''
Sub Standard_Legal_Status_Remain_Click()
    Call toggleSelect(Standard_Legal_Status_Remain, Standard_Return_Legal_Status, Standard_Fetch_Legal_Status)
    Standard_Legal_Status_Update.BackColor = unselectedColor
End Sub
Sub Standard_Legal_Status_Update_Click()
    Modal_Standard_Legal_Status.Show
End Sub


Sub Standard_Certification_Remain_Click()
    Call toggleSelect(Standard_Certification_Remain, Standard_Return_Certification, Standard_Fetch_Certification)
    Standard_Certification_Update.BackColor = unselectedColor
End Sub
Sub Standard_Certification_Update_Click()
    Modal_Standard_Certification.Show
End Sub


Sub Standard_Admission_Remain_Click()
    Call toggleSelect(Standard_Admission_Remain, Standard_Return_Admission, Standard_Fetch_Admission)
    Standard_Admission_Update.BackColor = unselectedColor
End Sub
Sub Standard_Admission_Update_Click()
    Modal_Standard_Admission.Show
End Sub


Sub Standard_Adjudication_Remain_Click()
    Call toggleSelect(Standard_Adjudication_Remain, Standard_Return_Adjudication, Standard_Fetch_Adjudication)
    Standard_Adjudication_Update.BackColor = unselectedColor
End Sub
Sub Standard_Adjudication_Update_Click()
    Modal_Standard_Adjudication.Show
End Sub

Sub Standard_Continuance_Remain_Click()
    Call toggleSelect(Standard_Continuance_Remain, Standard_Return_Continuance, "No")
    Standard_Continuance_Update.BackColor = unselectedColor
End Sub
Sub Standard_Continuance_Update_Click()
    Modal_Standard_Continuance.Show
End Sub

Sub Standard_Supervision_Add_Click()
    Modal_Standard_Add_Supervision.Show
End Sub
Sub Standard_Supervision_Discharge_Click()
    Modal_Standard_Drop_Supervision.Show
End Sub
Sub Standard_Supervision_Remain_Click()
    Call toggleSelect(Standard_Supervision_Remain)
End Sub

Sub Standard_Condition_Add_Click()
    Modal_Standard_Add_Condition.Show
End Sub
Sub Standard_Condition_Discharge_Click()
    Modal_Standard_Drop_Condition.Show
End Sub
Sub Standard_Condition_Remain_Click()
    Call toggleSelect(Standard_Condition_Remain)
End Sub
Private Sub Standard_Remain_All_Click()
    If Not Standard_Legal_Status_Update.BackColor = selectedColor Then
        Call Standard_Legal_Status_Remain_Click
    End If

    If Not Standard_Certification_Update.BackColor = selectedColor Then
        Call Standard_Certification_Remain_Click
    End If

    If Not Standard_Admission_Update.BackColor = selectedColor Then
        Call Standard_Admission_Remain_Click
    End If

    If Not Standard_Adjudication_Update.BackColor = selectedColor Then
        Call Standard_Adjudication_Remain_Click
    End If

    If Not Standard_Continuance_Update.BackColor = selectedColor Then
        Call Standard_Continuance_Remain_Click
    End If

End Sub

'''''''''''''''''''
'JTC_PHASE_UPDATES'
'''''''''''''''''''
Sub JTC_Accept_Click()
    Modal_JTC_Accept.Show
End Sub
Sub JTC_Reject_Click()
    Modal_JTC_Reject.Show
End Sub
Sub JTC_Phase_Stepup_Click()
    Modal_JTC_Stepup.Show
End Sub
Sub JTC_Phase_Pushback_Click()
    Modal_JTC_Phase_Pushback.Show
End Sub
Sub JTC_Discharge_Click()
    Modal_JTC_Discharge.Show
End Sub
Sub JTC_Expungement_Click()
    Modal_JTC_Expungement.Show
End Sub

Sub JTC_Phase_Remain_Click()
    'take displayed phase and stepup
    JTC_Return_Phase.Caption = JTC_Fetch_Phase.Caption
    JTC_Return_Stepup_Date.Caption = JTC_Fetch_Stepup_Date.Caption

    'hide pushback display
    JTC_Pushback_Reason_Label.Visible = False
    JTC_Pushback_Reason1.Caption = ""
    JTC_Pushback_Reason2.Caption = ""
    JTC_Pushback_Reason3.Caption = ""

    'color buttons
    JTC_Phase_Remain.BackColor = selectedColor
    JTC_Phase_Stepup.BackColor = unselectedColor
    JTC_Phase_Pushback.BackColor = unselectedColor
    JTC_Reject.BackColor = unselectedColor
    JTC_Discharge.BackColor = unselectedColor
End Sub

'''''''''''''''''''''''
'JTC_TREATMENT_UPDATES'
'''''''''''''''''''''''

Sub JTC_Treatment_Stepdown_Click()
    Modal_JTC_Stepdown.Show
End Sub

Sub JTC_Treatment_Provider_Update_Click()
    Modal_JTC_Provider.Show
End Sub

Sub JTC_Treatment_Provider_Remain_Click()
    'push provider name
    JTC_Return_Treatment_Provider.Caption = JTC_Fetch_Treatment_Provider.Caption

    'hide stepdown display
    JTC_Return_Stepdown_Date.Caption = ""
    ClientUpdateForm.JTC_Stepdown_Label.Visible = False

    'color buttons
    JTC_Treatment_Provider_Remain.BackColor = selectedColor
    JTC_Treatment_Stepdown.BackColor = unselectedColor
    JTC_Treatment_Provider_Update.BackColor = unselectedColor
    JTC_Treatment_Discharge.BackColor = unselectedColor
End Sub

Sub JTC_Treatment_Discharge_Click()
    JTC_Treatment_Provider_Remain.BackColor = unselectedColor
    JTC_Treatment_Stepdown.BackColor = unselectedColor
    JTC_Treatment_Provider_Update.BackColor = unselectedColor
    JTC_Treatment_Discharge.BackColor = selectedColor
End Sub
Sub JTC_Certification_Remain_Click()
    Call toggleSelect(JTC_Certification_Remain, JTC_Return_Certification, JTC_Fetch_Certification)
    JTC_Certification_Update.BackColor = unselectedColor
End Sub
Sub JTC_Certification_Update_Click()
    Modal_JTC_Certification.Show
End Sub


Sub JTC_Admission_Remain_Click()
    Call toggleSelect(JTC_Admission_Remain, JTC_Return_Admission, JTC_Fetch_Admission)
    JTC_Admission_Update.BackColor = unselectedColor
End Sub
Sub JTC_Admission_Update_Click()
    Modal_JTC_Admission.Show
End Sub


Sub JTC_Adjudication_Remain_Click()
    Call toggleSelect(JTC_Adjudication_Remain, JTC_Return_Adjudication, JTC_Fetch_Adjudication)
    JTC_Adjudication_Update.BackColor = unselectedColor
End Sub
Sub JTC_Adjudication_Update_Click()
    Modal_JTC_Adjudication.Show
End Sub

Sub JTC_Continuance_Remain_Click()
    Call toggleSelect(JTC_Continuance_Remain, JTC_Return_Continuance, "No")
    JTC_Continuance_Update.BackColor = unselectedColor
End Sub
Sub JTC_Continuance_Update_Click()
    Modal_JTC_Continuance.Show
End Sub
Sub JTC_Condition_Add_Click()
    Modal_JTC_Add_Condition.Show
End Sub
Sub JTC_Condition_Discharge_Click()
    Modal_JTC_Drop_Condition.Show
End Sub
Sub JTC_Condition_Remain_Click()
    Call toggleSelect(JTC_Condition_Remain)
End Sub

Sub JTC_Service_Add_Click()
    Modal_JTC_Add_Service.Show
End Sub

Sub JTC_Service_Discharge_Click()
    Modal_JTC_Drop_Service.Show
End Sub

Sub JTC_Service_Remain_Click()
    If JTC_Service_Remain.BackColor = unselectedColor Then
        JTC_Service_Remain.BackColor = selectedColor
    Else
        JTC_Service_Remain.BackColor = unselectedColor
    End If
End Sub

Private Sub JTC_FTA_Yes_Click()
    If Worksheets("Entry").Range(hFind("Active B/W?") & updateRow).value = Lookup("Generic_YNOU_Name")("Yes") Then
        JTC_FTA_Yes.BackColor = selectedColor
        JTC_FTA_No.BackColor = unselectedColor
    Else
        Modal_FTA.Show
    End If
End Sub

Private Sub JTC_FTA_No_Click()
    JTC_FTA_Yes.BackColor = unselectedColor
    JTC_FTA_No.BackColor = selectedColor
End Sub


Private Sub JTC_Lift_BW_Click()
    Modal_Lift_BW.Show
End Sub

Private Sub JTC_Remain_All_Click()
    If Not JTC_Certification_Update.BackColor = selectedColor Then
        Call JTC_Certification_Remain_Click
    End If

    If Not JTC_Admission_Update.BackColor = selectedColor Then
        Call JTC_Admission_Remain_Click
    End If

    If Not JTC_Adjudication_Update.BackColor = selectedColor Then
        Call JTC_Adjudication_Remain_Click
    End If

    If Not JTC_Continuance_Update.BackColor = selectedColor Then
        Call JTC_Continuance_Remain_Click
    End If

    If Not JTC_Treatment_Provider_Update.BackColor = selectedColor _
    And Not JTC_Treatment_Discharge.BackColor = selectedColor _
    And Not JTC_Treatment_Stepdown.BackColor = selectedColor Then
        Call JTC_Treatment_Provider_Remain_Click
    End If

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


Private Sub InitialHearingDate_Enter()
    InitialHearingDate.value = CalendarForm.GetDate(RangeOfYears:=5)
End Sub
Private Sub InitialHearingDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Set ctl = Me.InitialHearingDate
    Call DateValidation(ctl, Cancel)
End Sub


Private Sub Intake_Cancel_Click()
    Unload Me
End Sub


Private Sub Intake_Submit_Click()
    On Error GoTo err

    Call Generate_Dictionaries
    
    

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

    If DRAI_Action.value = "Follow - Hold" Or DRAI_Action.value = "Override - Hold" Then
        If DetentionFacility.value = "N/A" Then
            MsgBox "Detention facility required for call-in hold"
            Exit Sub
        End If
    End If

    
    If DHS_Status.value = "" And InConfRecord.value = "Yes" Then
        MsgBox "'DHS Status' Required"
        Exit Sub
    End If
    
     If Diagnosis1.value = "" And InConfRecord.value = "Yes" Then
        MsgBox "Diagnosis #1 Required"
        Exit Sub
    End If
    
    If Treatment1.value = "" And InConfRecord.value = "Yes" Then
        MsgBox "Treatment #1 Required"
        Exit Sub
    End If
    
    If TraumaType1.value = "" And InConfRecord.value = "Yes" Then
        MsgBox "Trauma #1 Required"
        Exit Sub
    End If
    
    
    If InitialHearingLocation.value = "Diversion" And ConfOutcome.value = "Release for Court" Then
        MsgBox "Diversion is not a valid courtroom for release"
        Exit Sub
    End If

    Call formSubmitStart(updateRow)

    ''''''''''''''''''
    'SET LEGAL STATUS'
    ''''''''''''''''''

    Dim bucketHead As String


    If InitialHearingLocation.value = "Adult" Then
        Call legalStatusStart( _
            clientRow:=updateRow, _
            statusType:="Adult", _
            Courtroom:="Adult", _
            DA:=DA.value, _
            startDate:=DateOfHearing.value)
    End If


    '''''''''''''''''''
    ''''''CALL IN''''''
    '''''''''''''''''''

    tempHead = headerFind("CALL-IN 2")

    If CallInRecord.value = "Yes" Then
        Range(headerFind("Did Youth Have Call-In?", tempHead) & updateRow).value _
                = Lookup("Generic_NYNOU_Name")("Yes")

        Range(headerFind("Date of Call-In", tempHead) & updateRow).value _
                    = CallInDate.value

        Range(headerFind("Was DRAI Administered?", tempHead) & updateRow).value _
                = Lookup("Generic_NYNOU_Name")(Was_DRAI_Administered.value)

        Range(headerFind("DRAI Score", tempHead) & updateRow).value _
                = DRAI_Score.value

        Select Case DRAI_Score.value
            Case Is < 10
                Range(hFind("DRAI Recommendation", "CALL-IN 2") & updateRow).value _
                        = Lookup("DRAI_Recommendation_Name")("Release")
            Case Is < 15
                Range(hFind("DRAI Recommendation", "CALL-IN 2") & updateRow).value _
                        = Lookup("DRAI_Recommendation_Name")("Release w/ Supervision")
            Case Is < 30
                Range(hFind("DRAI Recommendation", "CALL-IN 2") & updateRow).value _
                        = Lookup("DRAI_Recommendation_Name")("Release w/ Supervision")
            Case Else
                Range(hFind("DRAI Recommendation", "CALL-IN 2") & updateRow).value _
                        = Lookup("DRAI_Recommendation_Name")("Unknown")
        End Select


        Range(headerFind("DRAI Recommendation", tempHead) & updateRow).value _
                = Lookup("DRAI_Recommendation_Name")(DRAI_Rec.value)

        Range(headerFind("DRAI Action", tempHead) & updateRow).value _
                = Lookup("DRAI_Action_Name")(DRAI_Action.value)

        Select Case DRAI_Action.value
            Case "Override - Hold", "Follow - Hold"
                Range(headerFind("End Date", tempHead) & updateRow).value _
                        = InConfDate.value
                Range(headerFind("LOS in Detention", tempHead) & updateRow).value _
                        = calcLOS(CallInDate.value, InConfDate.value)
                Call addSupervision( _
                    clientRow:=updateRow, _
                    serviceType:="Detention (not respite)", _
                    legalStatus:="Pretrial", _
                    Courtroom:="Call-In", _
                    CourtroomOfOrder:="Call-In", _
                    DA:=DA.value, _
                    agency:="PJJSC", _
                    startDate:=CallInDate.value, _
                    endDate:=InConfDate.value, _
                    re1:="", _
                    re2:="", _
                    re3:="", _
                    Notes:="Held at call-in")

        End Select
        Range(headerFind("LOS Until Next Hearing", tempHead) & updateRow).value _
                        = calcLOS(CallInDate.value, InConfDate.value)

        Range(headerFind("Detention Facility", tempHead) & updateRow).value _
                    = Lookup("Detention_Facility_Name")(DetentionFacility.value)

        Range(headerFind("Reason #1 for Override Hold", tempHead) & updateRow).value _
                = Lookup("DRAI_Override_Reason_Name")(OverrideHoldRe1.value)
        Range(headerFind("Reason #2 for Override Hold", tempHead) & updateRow).value _
                = Lookup("DRAI_Override_Reason_Name")(OverrideHoldRe2.value)
        Range(headerFind("Reason #3 for Override Hold", tempHead) & updateRow).value _
                = Lookup("DRAI_Override_Reason_Name")(OverrideHoldRe3.value)
    Else
        Range(headerFind("Did Youth Have Call-In?", tempHead) & updateRow).value _
                = Lookup("Generic_NYNOU_Name")("Unknown")
    End If


    '''''''''''''''''''
    'Intake Conference'
    '''''''''''''''''''

    tempHead = headerFind("INTAKE CONFERENCE 2")



    If InConfRecord.value = "Yes" Then
        Range(headerFind("Did Youth Have an Intake Conference?", tempHead) & updateRow).value _
                = Lookup("Generic_NYNOU_Name")("Yes")

        Range(headerFind("Date of Intake Conference", tempHead) & updateRow).value _
                = InConfDate.value

        Range(headerFind("Intake Conference Type", tempHead) & updateRow).value _
                = Lookup("Intake_Conference_Type_Name")(InConfType.value)

        Range(headerFind("DA", tempHead) & updateRow).value _
                = Lookup("DA_Last_Name_Name")(DA.value)


        Range(headerFind("Intake Conference Outcome", tempHead) & updateRow).value _
                = Lookup("Intake_Conference_Outcome_Name")(ConfOutcome.value)
        
        Range(hFind("Status at Arrest", "DHS") & updateRow).value _
                = Lookup("DHS_Status_at_Arrest_Name")(DHS_Status.value)
        
        If DHS_Status.value = "N/A" Or DHS_Status.value = "None" Or DHS_Status.value = "Unknown" Then
            Range(hFind("Did youth have any DHS contact?", "DHS") & updateRow).value = 2 'no
        Else
            Range(hFind("Did youth have any DHS contact?", "DHS") & updateRow).value = 1 'yes
        End If

        Range(headerFind("Location of Next Event", tempHead) & updateRow).value _
                = Lookup("Courtroom_Name")(InitialHearingLocation.value)

        Range(headerFind("Next Event Date", tempHead) & updateRow).value _
                = InitialHearingDate.value

        Range(headerFind("LOS from Arrest Until Conference", tempHead) & updateRow).value _
                    = calcLOS(Range(headerFind("Arrest Date") & updateRow).value, InConfDate.value)

        tempHead = headerFind("Supervision Ordered #1", tempHead)

        Range(tempHead & updateRow).value _
                = Lookup("Supervision_Program_Name")(Supv1.value)
        Range(headerFind("Community-Based Agency #1", tempHead) & updateRow).value _
                = Lookup("Community_Based_Supervision_Provider_Name")(Supv1Pro.value)

        Range(headerFind("Reason #1 for Supervision Referral", tempHead) & updateRow).value _
                = Lookup("Supervision_Referral_Reason_Name")(Supv1Re1.value)
        Range(headerFind("Reason #2 for Supervision Referral", tempHead) & updateRow).value _
                = Lookup("Supervision_Referral_Reason_Name")(Supv1Re2.value)
        Range(headerFind("Reason #3 for Supervision Referral", tempHead) & updateRow).value _
                = Lookup("Supervision_Referral_Reason_Name")(Supv1Re3.value)

        tempHead = headerFind("Supervision Ordered #2", tempHead)

        Range(tempHead & updateRow).value _
                = Lookup("Supervision_Program_Name")(Supv2.value)
        Range(headerFind("Community-Based Agency #2", tempHead) & updateRow).value _
                = Lookup("Community_Based_Supervision_Provider_Name")(Supv2Pro.value)


        Range(headerFind("Reason #1 for Supervision Referral", tempHead) & updateRow).value _
                = Lookup("Supervision_Referral_Reason_Name")(Supv2Re1.value)
        Range(headerFind("Reason #2 for Supervision Referral", tempHead) & updateRow).value _
                = Lookup("Supervision_Referral_Reason_Name")(Supv2Re2.value)
        Range(headerFind("Reason #3 for Supervision Referral", tempHead) & updateRow).value _
                = Lookup("Supervision_Referral_Reason_Name")(Supv2Re3.value)

        Range(headerFind("Other Condition #1", tempHead) & updateRow).value _
                = Lookup("Condition_Name")(Cond1.value)
        Range(headerFind("Other Condition #1 Provider", tempHead) & updateRow).value _
                = Lookup("Condition_Provider_Name")(Cond1Pro.value)

        Range(headerFind("Other Condition #2", tempHead) & updateRow).value _
                = Lookup("Condition_Name")(Cond2.value)
        Range(headerFind("Other Condition #2 Provider", tempHead) & updateRow).value _
                = Lookup("Condition_Provider_Name")(Cond2Pro.value)

        Range(headerFind("Other Condition #3", tempHead) & updateRow).value _
                = Lookup("Condition_Name")(Cond3.value)
        Range(headerFind("Other Condition #3 Provider", tempHead) & updateRow).value _
                = Lookup("Condition_Provider_Name")(Cond3Pro.value)
                
        Range(headerFind("Diagnosis #1") & updateRow).value = Lookup("Diagnosis_Name")(Diagnosis1.value)
        Range(headerFind("Diagnosis #2") & updateRow).value = Lookup("Diagnosis_Name")(Diagnosis2.value)
        Range(headerFind("Diagnosis #3") & updateRow).value = Lookup("Diagnosis_Name")(Diagnosis3.value)
        Range(headerFind("Trauma Type #1") & updateRow).value = Lookup("Trauma_Type_Name")(TraumaType1.value)
        Range(headerFind("Trauma Type #2") & updateRow).value = Lookup("Trauma_Type_Name")(TraumaType2.value)
        Range(headerFind("Trauma Type #3") & updateRow).value = Lookup("Trauma_Type_Name")(TraumaType3.value)
        Range(headerFind("Treatment #1") & updateRow).value = Lookup("Treatment_Name")(Treatment1.value)
        Range(headerFind("Treatment #2") & updateRow).value = Lookup("Treatment_Name")(Treatment2.value)
        Range(headerFind("Treatment #3") & updateRow).value = Lookup("Treatment_Name")(Treatment3.value)
        

        Range(headerFind("Notes on Intake Conference", tempHead) & updateRow).value _
                = GeneralNotes.value

        Select Case ConfOutcome.value
            Case "Hold for Detention"
                Range(headerFind("Active Courtroom") & updateRow).value _
                         = Lookup("Courtroom_Name")("PJJSC")
                Call flagNo(Range(hFind("Did Youth Have Initial Detention Hearing?", "DETENTION") & updateRow))
                Range(hFind("Detention Facility", "DETENTION") & updateRow).value _
                         = Lookup("Detention_Facility_Name")(DetentionFacility.value)
                Call addSupervision( _
                    clientRow:=updateRow, _
                    serviceType:="Detention (not respite)", _
                    legalStatus:="Pretrial", _
                    Courtroom:="Intake Conf.", _
                    DA:=DA.value, _
                    agency:="", _
                    startDate:=InConfDate.value, _
                    re1:="", _
                    re2:="", _
                    re3:="", _
                    Notes:="Held at intake conference")
            Case "Roll to Detention Hearing"
                    Range(headerFind("Active Courtroom") & updateRow).value _
                         = Lookup("Courtroom_Name")("PJJSC")

            Case "Release for Court"
                Call ReferClientTo( _
                    referralDate:=InConfDate.value, _
                    clientRow:=updateRow, _
                    fromCR:="Intake Conf.", _
                    toCR:=InitialHearingLocation.value, _
                    DA:=DA.value _
                    )
                If InitialHearingLocation.value = "5E" Then
                    Range(hFind("Courtroom of Origin", "Crossover") & updateRow).value _
                            = Lookup("Courtroom_Name")("Intake Conf.")
                Else
                    Range(hFind("Courtroom of Origin", InitialHearingLocation.value) & updateRow).value _
                            = Lookup("Courtroom_Name")("Intake Conf.")
                End If

                'add supervisions and conditions if assigned
                If Not Supv1.value = "None" Then
                    Call addSupervision( _
                        clientRow:=updateRow, _
                        serviceType:=Supv1.value, _
                        legalStatus:="Pretrial", _
                        Courtroom:="Intake Conf.", _
                        CourtroomOfOrder:="Intake Conf.", _
                        DA:=DA.value, _
                        agency:=Supv1Pro.value, _
                        startDate:=InConfDate.value, _
                        NextCourtDate:=InitialHearingDate.value, _
                        re1:=Supv1Re1.value, _
                        re2:=Supv1Re2.value, _
                        re3:=Supv1Re3.value, _
                        Notes:="Referred at intake conference")
                End If

                If Not Supv2.value = "None" Then
                    Call addSupervision( _
                        clientRow:=updateRow, _
                        serviceType:=Supv2.value, _
                        legalStatus:="Pretrial", _
                        Courtroom:="Intake Conf.", _
                        CourtroomOfOrder:="Intake Conf.", _
                        DA:=DA.value, _
                        agency:=Supv2Pro.value, _
                        startDate:=InConfDate.value, _
                        NextCourtDate:=InitialHearingDate.value, _
                        re1:=Supv2Re1.value, _
                        re2:=Supv2Re2.value, _
                        re3:=Supv2Re3.value, _
                        Notes:="Referred at intake conference")
                End If

                If Not Cond1.value = "None" Then
                    Call addCondition( _
                        clientRow:=updateRow, _
                        condition:=Cond1.value, _
                        legalStatus:="Pretrial", _
                        Courtroom:="Intake Conf.", _
                        CourtroomOfOrder:="Intake Conf.", _
                        DA:=DA.value, _
                        agency:=Cond1Pro.value, _
                        startDate:=InConfDate.value, _
                        re1:="N/A", _
                        re2:="N/A", _
                        re3:="N/A", _
                        Notes:="Referred at intake conference")
                End If

                If Not Cond2.value = "None" Then
                    Call addCondition( _
                        clientRow:=updateRow, _
                        condition:=Cond2.value, _
                        legalStatus:="Pretrial", _
                        Courtroom:="Intake Conf.", _
                        CourtroomOfOrder:="Intake Conf.", _
                        DA:=DA.value, _
                        agency:=Cond2Pro.value, _
                        startDate:=InConfDate.value, _
                        re1:="N/A", _
                        re2:="N/A", _
                        re3:="N/A", _
                        Notes:="Referred at intake conference")
                End If

                If Not Cond3.value = "None" Then
                    Call addCondition( _
                        clientRow:=updateRow, _
                        condition:=Cond3.value, _
                        legalStatus:="Pretrial", _
                        Courtroom:="Intake Conf.", _
                        CourtroomOfOrder:="Intake Conf.", _
                        DA:=DA.value, _
                        agency:=Cond3Pro.value, _
                        startDate:=InConfDate.value, _
                        re1:="N/A", _
                        re2:="N/A", _
                        re3:="N/A", _
                        Notes:="Referred at intake conference")
                End If

            Case "Release for Diversion"

        End Select
    Else
        Range(headerFind("Did Youth Have an Intake Conference?", tempHead) & updateRow).value _
                = Lookup("Generic_NYNOU_Name")("Unknown")

        Select Case InitialHearingLocation.value
            Case "4G", "4E", "6F", "6H", "3E", "JTC", "5E", "WRAP", "Adult"
                Call ReferClientTo( _
                    referralDate:=DateOfHearing.value, _
                    clientRow:=updateRow, _
                    toCR:=InitialHearingLocation.value, _
                    DA:=DA.value _
                    )
        End Select
    End If
    

    Dim noteDate As String
    
    If Not InConfDate.value = "" Then
        noteDate = InConfDate.value
    Else
        If Not InitialHearingDate.value = "" Then
            noteDate = InitialHearingDate.value
        Else
            noteDate = "Date not available"
        End If
    End If
    
    Call addNotes( _
        Courtroom:=InitialHearingLocation.value, _
        DateOf:=noteDate, _
        userRow:=updateRow, _
        Notes:=GeneralNotes, _
        DA:=DA.value _
    )
    
    Call formSubmitEnd
    
done:
    Call UnloadAll
    Exit Sub
err:

        Call loadFromCache(2)

    'Stop 'press F8 twice to see the error point
    'Resume
    MsgBox "Something went wrong. Database has been restored to state prior to submission. " _
      & vbNewLine & vbNewLine & "Message: " & vbNewLine & err.Description _
      & vbNewLine & vbNewLine & "Source: " & vbNewLine & err.Source

    Call UnloadAll
End Sub

