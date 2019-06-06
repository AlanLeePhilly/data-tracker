VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ClientUpdateForm 
   Caption         =   "ClientUpdateForm"
   ClientHeight    =   11580
   ClientLeft      =   45
   ClientTop       =   -75
   ClientWidth     =   17910
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

    If Not JTC_Treatment_Provider_Update.BackColor = selectedColor Then
        Call JTC_Treatment_Provider_Remain_Click
    End If

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

    'define variable Long(a big integer) named emptyRow
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




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'JTC_SUBMIT_CLICK'''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Sub JTC_Submit_Click()

    On Error GoTo err

    Dim restorer As Variant

    restorer = Sheets("Entry").Range("C" & updateRow & ":" & hFind("END") & updateRow).value

    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With

    Worksheets("Entry").Activate

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
    ''EXPUNGEMENT''
    '''''''''''''''

    If JTC_Fetch_Phase = "Graduated, Awaiting Expungment" Then
        If JTC_Return_Phase = "Graduated, Record Expunged" Then
            Range(headerFind("Phase", courtHead) & updateRow) = Lookup("JTC_Phase_Name")("Graduated, Record Expunged")
            Range(headerFind("Record Expunged?", newPhaseHead) & updateRow) = Lookup("Generic_YN_Name")("Yes")
            Range(headerFind("Date of Expungement", newPhaseHead) & updateRow) = JTC_Accept_Reject_Date.Caption
            Range(headerFind("LOS (expungement)", newPhaseHead) & updateRow) _
                = DateDiff("d", Range(headerFind("Accepted Date", courtHead) & updateRow), JTC_Accept_Reject_Date.Caption)
        Else
            MsgBox "Invalid submission."
            Exit Sub
        End If
    End If

    '''''''''''''''''
    ''PHASE SECTION''
    '''''''''''''''''
    If JTC_Reject.BackColor = selectedColor Then
        Range(headerFind("Accepted (Y/N)", courtHead) & updateRow).value = 2
        Range(headerFind("Rejected Date", courtHead) & updateRow).value = DateOfHearing.value
        Range(headerFind("Next Hearing Location (if rejected)", courtHead) & updateRow).value = _
                Lookup("Courtroom_Name")(Modal_JTC_Reject.ReferredTo.value)
        Call ReferClientTo( _
                    referralDate:=DateOfHearing.value, _
                    clientRow:=updateRow, _
                    toCR:=Modal_JTC_Reject.ReferredTo.value, _
                    fromCR:="JTC", _
                    Notes:="Rejected from JTC")
        Call Cancel_Click
        Worksheets("User Entry").Activate
        Exit Sub
    End If

    If JTC_Accept.BackColor = selectedColor Then
        Range(headerFind("Phase") & updateRow).value = "1"
        Range(headerFind("Accepted (Y/N)", courtHead) & updateRow).value = 1 'Yes
        Range(headerFind("Accepted Date", courtHead) & updateRow).value = JTC_Accept_Reject_Date.Caption
        Range(headerFind("Start Date", newPhaseHead) & updateRow).value = JTC_Accept_Reject_Date.Caption
        Range(headerFind("Scheduled Step-Up Date", newPhaseHead) & updateRow) = JTC_Return_Stepup_Date.Caption

        Call endLegalStatus( _
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

        Call startLegalStatus( _
            clientRow:=updateRow, _
            statusType:="JTC", _
            Courtroom:="JTC", _
            DA:=DA.value, _
            startDate:=DateOfHearing.value, _
            Notes:="Accepted to JTC")

        Call closeOpenLegalStatuses( _
            clientRow:=updateRow, _
            dateOf:=DateOfHearing.value, _
            Courtroom:="JTC", _
            DA:=DA.value)
    End If

    'if step-up is selected
    If JTC_Phase_Stepup.BackColor = selectedColor Then
        'set DoH as old phase end date and new phase begin date
        Range(headerFind("End Date", oldPhaseHead) & updateRow) = DateOfHearing.value
        Range(headerFind("Start Date", newPhaseHead) & updateRow) = DateOfHearing.value
        Range(headerFind("LOS", oldPhaseHead) & updateRow) _
                        = DateDiff("d", Range(headerFind("Start Date", oldPhaseHead) & updateRow), DateOfHearing.value)
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

    'if discharge
    If JTC_Discharge.BackColor = selectedColor Then
        Dim endHead As String
        endHead = hFind("Petition Outcomes", "AGGREGATES")
        tempHead = headerFind("JTC OUTCOMES")
        'set current phase end date
        Range(headerFind("End Date", oldPhaseHead) & updateRow) = DateOfHearing.value
        'set "Date of Overall Discharge"
        Range(headerFind("Date of Overall Discharge", tempHead) & updateRow) = DateOfHearing.value
        'set "Active or Discharged" to discharged
        Range(headerFind("Active or Discharged", tempHead) & updateRow) = Lookup("Active_Name")("Discharged")
        'set "Nature of Discharge"

        Select Case Modal_JTC_Discharge.DetailedOutcome.value
            Case "Rearrested & Held (adult)"
                Call totalOutcome( _
                            clientRow:=updateRow, _
                            dateOf:=DateOfHearing.value, _
                            Courtroom:="JTC", _
                            DA:=DA.value, _
                            legalStatus:="JTC", _
                            Nature:="Negative", _
                            detailed:="Rearrested & Held")

                ''''''''''''''''''''''
            Case "Positive Completion"
                Call totalOutcome( _
                            clientRow:=updateRow, _
                            dateOf:=DateOfHearing.value, _
                            Courtroom:="JTC", _
                            DA:=DA.value, _
                            legalStatus:="JTC", _
                            Nature:="Positive", _
                            detailed:="Petition Closed - Positive Comp. Terms")

                If JTC_Fetch_Phase = 3 Then
                    Range(headerFind("Phase", courtHead) & updateRow) = Lookup("JTC_Phase_Name")("Graduated, Awaiting Expungment")
                    Range(headerFind("Record Expunged?", newPhaseHead) & updateRow) = Lookup("Generic_YN_Name")("No")
                    Range(headerFind("LOS", oldPhaseHead) & updateRow).value _
                                = DateDiff("d", Range(headerFind("Start Date", oldPhaseHead) & updateRow).value, DateOfHearing.value)
                End If
                Range(headerFind("LOS (discharged)") & updateRow).value _
                            = DateDiff("d", Range(headerFind("Arrest Date") & updateRow).value, DateOfHearing.value)

                '''''''''''''''''''''
            Case "Aged Out"
                Call totalOutcome( _
                            clientRow:=updateRow, _
                            dateOf:=DateOfHearing.value, _
                            Courtroom:="JTC", _
                            DA:=DA.value, _
                            legalStatus:="JTC", _
                            Nature:="Negative", _
                            detailed:="Aged Out")

                '''''''''''''''''''''
            Case "Acceptance Not Granted", "Show Cause", "Hosp. (Mental Health)", "Hosp. (Physical Health)", "Other", "Unknown"
                Call ReferClientTo( _
                            referralDate:=DateOfHearing.value, _
                            clientRow:=updateRow, _
                            toCR:=Modal_JTC_Discharge.New_CR.value, _
                            fromCR:="JTC", _
                            newLegalStatus:="Probation")

                '''''''''''''''''''''
            Case "Transfer to Dependent"
                Call ReferClientTo( _
                            referralDate:=DateOfHearing.value, _
                            clientRow:=updateRow, _
                            toCR:="5E", _
                            fromCR:="JTC")

                '''''''''''''''''''''
            Case "Transfer to Other County"
                Call totalOutcome( _
                            clientRow:=updateRow, _
                            dateOf:=DateOfHearing.value, _
                            Courtroom:="JTC", _
                            DA:=DA.value, _
                            legalStatus:="JTC", _
                            Nature:="Neutral", _
                            detailed:="Transfer to Other County")

        End Select

        If JTC_Return_Phase = "Positive Discharge" Then
            Range(headerFind("Nature of Discharge", tempHead) & updateRow) _
                            = Lookup("Nature_of_Discharge_Name")("Positive")
        End If

        If JTC_Return_Phase = "Negative Discharge" Then
            Range(headerFind("Nature of Discharge", tempHead) & updateRow) _
                            = Lookup("Nature_of_Discharge_Name")("Negative")
        End If

        If JTC_Return_Phase = "Neutral Discharge" Then
            Range(headerFind("Nature of Discharge", tempHead) & updateRow) _
                            = Lookup("Nature_of_Discharge_Name")("Neutral")
        End If

        'set detailed outcome
        Range(headerFind("Detailed Courtroom Outcome", tempHead) & updateRow) = _
                        Lookup("JTC_Outcome_Name")(Modal_JTC_Discharge.DetailedOutcome.value)
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
                        DateDiff("d", Range(headerFind("Accepted Date", courtHead) & updateRow), DateOfHearing)
        'calc LOS from Arrest
        Range(headerFind("Total LOS from Arrest", tempHead) & updateRow) = _
                        DateDiff("d", Range(headerFind("Arrest Date") & updateRow), DateOfHearing)
        '######set T/F discharge reasons
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
                                DateDiff("d", Range(headerFind("Referral Date", hFind("IOP Provider #1", "JTC")) & updateRow).value, Modal_JTC_Provider.Referral_Date.value)

            Case isEmptyOrZero(Range(hFind("IOP Provider #3", "JTC") & updateRow))
                Range(hFind("IOP Provider #3", "JTC") & updateRow) = _
                                Lookup("IOP_Provider_Name")(JTC_Return_Treatment_Provider.Caption)
                Range(hFind("Referral Date", "IOP Provider #3", "JTC") & updateRow) = Modal_JTC_Provider.Referral_Date
                Range(hFind("Discharge Date", "IOP Provider #2", "JTC") & updateRow) = DateOfHearing
                Range(hFind("LOS IOP", "IOP Provider #2", "JTC") & updateRow) = _
                                DateDiff("d", Range(headerFind("Referral Date", headerFind("IOP Provider #2", courtHead)) & updateRow).value, Modal_JTC_Provider.Referral_Date.value)

            Case Else
                Range(hFind("Discharge Date", hFind("IOP Provider #3", "JTC")) & updateRow) = DateOfHearing
                Range(hFind("LOS IOP", "IOP Provider #3", "JTC") & updateRow) = _
                                DateDiff("d", Range(hFind("Referral Date", "IOP Provider #3", "JTC") & updateRow).value, Modal_JTC_Provider.Referral_Date.value)
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
            Case Range(hFind("IOP Provider #3", "JTC") & updateRow).value = Lookup("IOP_Provider_Name")(JTC_Return_Treatment_Provider.Caption)
                Range(hFind("Discharge Date", "IOP Provider #3", "JTC") & updateRow) = DateOfHearing
                Range(hFind("LOS IOP", "IOP Provider #3", "JTC") & updateRow) _
                    = DateDiff("d", Range(hFind("Referral Date", "IOP Provider #3", "JTC") & updateRow), _
                      Range(hFind("Discharge Date", "IOP Provider #3", "JTC") & updateRow))

            Case Range(hFind("IOP Provider #2", "JTC") & updateRow).value = Lookup("IOP_Provider_Name")(JTC_Return_Treatment_Provider.Caption)
                Range(hFind("Discharge Date", "IOP Provider #2", "JTC") & updateRow) = DateOfHearing
                Range(hFind("LOS IOP", "IOP Provider #2", "JTC") & updateRow) _
                    = DateDiff("d", Range(hFind("Referral Date", "IOP Provider #2", "JTC") & updateRow), _
                      Range(hFind("Discharge Date", "IOP Provider #2", "JTC") & updateRow))

            Case Range(hFind("IOP Provider #1", "JTC") & updateRow).value = Lookup("IOP_Provider_Name")(JTC_Return_Treatment_Provider.Caption)
                Range(hFind("Discharge Date", "IOP Provider #1", "JTC") & updateRow) = DateOfHearing
                Range(hFind("LOS IOP", "IOP Provider #1", "JTC") & updateRow) _
                      = DateDiff("d", Range(hFind("Referral Date", "IOP Provider #1", "JTC") & updateRow), _
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
        Call admissionStart( _
                clientRow:=updateRow, _
                petitionNum:=Modal_JTC_Admission.PetitionBox.value, _
                statusType:="JTC", _
                Courtroom:="JTC", _
                DA:=DA.value, _
                startDate:=Modal_JTC_Admission.Admission_Date.value, _
                Result:=Modal_JTC_Admission.Result.value, _
                detailed:=Modal_JTC_Admission.Detailed_Result.value _
            )
    End If

    If JTC_Adjudication_Update.BackColor = selectedColor Then
        Call adjudicationStart( _
                clientRow:=updateRow, _
                petitionNum:=Modal_JTC_Adjudication.PetitionBox.value, _
                Courtroom:="JTC", _
                DA:=DA.value, _
                startDate:=Modal_JTC_Adjudication.Adjudication_Date.value, _
                Type_of:=Modal_JTC_Adjudication.Type_of.value, _
                Re1:=Modal_JTC_Adjudication.Reason1.value, _
                Re2:=Modal_JTC_Adjudication.Reason2.value, _
                Re3:=Modal_JTC_Adjudication.Reason3.value, _
                Re4:=Modal_JTC_Adjudication.Reason4.value, _
                Re5:=Modal_JTC_Adjudication.Reason5.value _
            )
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
            If .List(i, 4) = "New" Then 'if new service
                Call addSupervision( _
                        clientRow:=updateRow, _
                        serviceType:=.List(i, 0), _
                        legalStatus:="JTC", _
                        Courtroom:="JTC", _
                        DA:=DA.value, _
                        agency:=.List(i, 1), _
                        startDate:=.List(i, 2), _
                        Re1:=.List(i, 6), _
                        Re2:=.List(i, 7), _
                        Re3:=.List(i, 8), _
                        NextCourtDate:=NextCourtDate.value, _
                        phase:=JTC_Return_Phase.Caption, _
                        Notes:=.List(i, 9))

            Else
                If Range(headerFind("Courtroom of Order", .List(i, 4)) & updateRow).value _
                        = Lookup("Courtroom_Name")("Intake Conf.") Then

                    Call dropSupervision( _
                            clientRow:=updateRow, _
                            head:=.List(i, 4), _
                            serviceType:=.List(i, 0), _
                            startDate:=.List(i, 2), _
                            endDate:=DateOfHearing.value, _
                            Nature:="Neutral", _
                            Re1:="N/A", _
                            Re2:="N/A", _
                            Re3:="N/A", _
                            Notes:="Continued from intake conf.")

                    If Not IsDate(.List(i, 3)) Then
                        Call addSupervision( _
                                clientRow:=updateRow, _
                                serviceType:=.List(i, 0), _
                                legalStatus:="JTC", _
                                Courtroom:="JTC", _
                                DA:=DA.value, _
                                agency:=.List(i, 1), _
                                startDate:=DateOfHearing.value, _
                                NextCourtDate:=Standard_NextCourtDate.value, _
                                Re1:="N/A", _
                                Re2:="N/A", _
                                Re3:="N/A", _
                                phase:=JTC_Return_Phase.Caption)
                    End If
                Else
                    If Range(headerFind("Courtroom of Order", .List(i, 4)) & updateRow).value _
                            = Lookup("Courtroom_Name")("PJJSC") Then

                        Call dropSupervision( _
                                clientRow:=updateRow, _
                                head:=.List(i, 4), _
                                serviceType:=.List(i, 0), _
                                startDate:=.List(i, 2), _
                                endDate:=DateOfHearing.value, _
                                Nature:="Neutral", _
                                Re1:="N/A", _
                                Re2:="N/A", _
                                Re3:="N/A", _
                                Notes:="Continued from PJJSC")

                        If Not IsDate(.List(i, 3)) Then
                            Call addSupervision( _
                                    clientRow:=updateRow, _
                                    serviceType:=.List(i, 0), _
                                    legalStatus:="JTC", _
                                    Courtroom:="JTC", _
                                    DA:=DA.value, _
                                    agency:=.List(i, 1), _
                                    startDate:=DateOfHearing.value, _
                                    NextCourtDate:=Standard_NextCourtDate.value, _
                                    Re1:="N/A", _
                                    Re2:="N/A", _
                                    Re3:="N/A", _
                                    phase:=JTC_Return_Phase.Caption)
                        End If
                    Else
                        If IsDate(.List(i, 3)) Then 'if has End Date
                            Call dropSupervision( _
                                    clientRow:=updateRow, _
                                    head:=.List(i, 4), _
                                    serviceType:=.List(i, 0), _
                                    startDate:=.List(i, 2), _
                                    endDate:=.List(i, 3), _
                                    Nature:=.List(i, 5), _
                                    Re1:=.List(i, 6), _
                                    Re2:=.List(i, 7), _
                                    Re3:=.List(i, 8), _
                                    Notes:=.List(i, 9))
                        End If
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
            If .List(i, 4) = "New" Then 'if new service
                Call addCondition( _
                        clientRow:=updateRow, _
                        condition:=.List(i, 0), _
                        legalStatus:="JTC", _
                        Courtroom:="JTC", _
                        DA:=DA.value, _
                        agency:=.List(i, 1), _
                        startDate:=.List(i, 2), _
                        Re1:=.List(i, 6), _
                        Re2:=.List(i, 7), _
                        Re3:=.List(i, 8), _
                        phase:=JTC_Return_Phase.Caption, _
                        Notes:=.List(i, 9))
            Else
                If Range(headerFind("Courtroom of Order", .List(i, 4)) & updateRow).value _
                        = Lookup("Courtroom_Name")("Intake Conf.") Then
                    Call dropCondition( _
                                clientRow:=updateRow, _
                                head:=.List(i, 4), _
                                condition:=.List(i, 0), _
                                startDate:=.List(i, 2), _
                                endDate:=DateOfHearing.value, _
                                Nature:="Neutral", _
                                Re1:="N/A", _
                                Re2:="N/A", _
                                Re3:="N/A", _
                                Notes:="Continued from intake conf.")

                    If Not IsDate(.List(i, 3)) Then
                        Call addCondition( _
                                clientRow:=updateRow, _
                                condition:=.List(i, 0), _
                                legalStatus:="JTC", _
                                Courtroom:="JTC", _
                                DA:=DA.value, _
                                agency:=.List(i, 1), _
                                startDate:=DateOfHearing.value, _
                                Re1:="N/A", _
                                Re2:="N/A", _
                                Re3:="N/A", _
                                phase:=JTC_Return_Phase.Caption)
                    End If
                Else
                    If Range(headerFind("Courtroom of Order", .List(i, 4)) & updateRow).value _
                            = Lookup("Courtroom_Name")("PJJSC") Then
                        Call dropCondition( _
                                clientRow:=updateRow, _
                                head:=.List(i, 4), _
                                condition:=.List(i, 0), _
                                startDate:=.List(i, 2), _
                                endDate:=DateOfHearing.value, _
                                Nature:="Neutral", _
                                Re1:="N/A", _
                                Re2:="N/A", _
                                Re3:="N/A", _
                                Notes:="Continued from PJJSC")

                        If Not IsDate(.List(i, 3)) Then
                            Call addCondition( _
                                    clientRow:=updateRow, _
                                    condition:=.List(i, 0), _
                                    legalStatus:="JTC", _
                                    Courtroom:="JTC", _
                                    DA:=DA.value, _
                                    agency:=.List(i, 1), _
                                    startDate:=DateOfHearing.value, _
                                    Re1:="N/A", _
                                    Re2:="N/A", _
                                    Re3:="N/A", _
                                    phase:=JTC_Return_Phase.Caption)
                        End If
                    Else
                        If IsDate(.List(i, 3)) Then 'if has End Date
                            Call dropCondition( _
                                    clientRow:=updateRow, _
                                    head:=.List(i, 4), _
                                    condition:=.List(i, 0), _
                                    startDate:=.List(i, 2), _
                                    endDate:=.List(i, 3), _
                                    Nature:=.List(i, 5), _
                                    Re1:=.List(i, 6), _
                                    Re2:=.List(i, 7), _
                                    Re3:=.List(i, 8), _
                                    Notes:=.List(i, 9))
                        End If
                    End If
                End If
            End If
        Next i
    End With

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
    Call closeCallIn(DateOfHearing.value, updateRow)
    Call closeIntakeConference(DateOfHearing.value, updateRow)
    Call addNotes("JTC", DateOfHearing.value, updateRow, JTC_Notes.value, "JTC")
    Call Save_Countdown
    Call UnloadAll

    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
    End With
    Worksheets("User Entry").Activate

done:


    Exit Sub
err:
    Sheets("Entry").Range("C" & updateRow & ":" & hFind("END") & updateRow).value = restorer

    Stop   'press F8 twice to see the error point
    Resume
    MsgBox "Something went wrong. Database has been restored to state prior to submission. " _
      & vbNewLine & vbNewLine & "Message: " & vbNewLine & err.Description _
      & vbNewLine & vbNewLine & "Source: " & vbNewLine & err.Source
    Call UnloadAll
End Sub

Sub Standard_Submit_Click()
    On Error GoTo err

    Worksheets("Entry").Activate
    Dim restorer As Variant
    restorer = Sheets("Entry").Range("C" & updateRow & ":" & hFind("END") & updateRow).value

    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With

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

    'ASSUMPTION: NO COURTROOM CHANGE

    Dim oldLegalHead As String
    Dim newLegalHead As String
    Dim bucketHead As String
    Dim i As Long



    ''''''''''''''
    'LEGAL STATUS'
    ''''''''''''''
    If Standard_Legal_Status_Update.BackColor = selectedColor Then
        With Modal_Standard_Legal_Status
            Call endLegalStatus( _
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

            Call startLegalStatus( _
                    clientRow:=updateRow, _
                    statusType:=.New_Legal_Status, _
                    Courtroom:=newCourtroom, _
                    DA:=DA.value, _
                    startDate:=.New_Start_Date, _
                    Notes:=.New_Notes)
        End With
    Else
        Call startLegalStatus( _
                clientRow:=updateRow, _
                statusType:=Standard_Return_Legal_Status.Caption, _
                Courtroom:=newCourtroom, _
                DA:=DA.value, _
                startDate:=DateOfHearing.value, _
                Notes:="Continued from prior courtroom")
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
        Call admissionStart( _
                clientRow:=updateRow, _
                petitionNum:=Modal_Standard_Admission.PetitionBox.value, _
                statusType:=Lookup("Legal_Status_Num")(Range(headerFind("Legal Status") & updateRow).value), _
                Courtroom:=newCourtroom, _
                DA:=DA.value, _
                startDate:=Modal_Standard_Admission.Admission_Date.value, _
                Result:=Modal_Standard_Admission.Result.value, _
                detailed:=Modal_Standard_Admission.Detailed_Result.value _
            )
    End If

    ''''''''''''''
    'Adjudication'
    ''''''''''''''
    If Standard_Adjudication_Update.BackColor = selectedColor Then
        Call adjudicationStart( _
                clientRow:=updateRow, _
                petitionNum:=Modal_Standard_Adjudication.PetitionBox.value, _
                Courtroom:=newCourtroom, _
                DA:=DA.value, _
                startDate:=Modal_Standard_Adjudication.Adjudication_Date.value, _
                Type_of:=Modal_Standard_Adjudication.Type_of.value, _
                Re1:=Modal_Standard_Adjudication.Reason1.value, _
                Re2:=Modal_Standard_Adjudication.Reason2.value, _
                Re3:=Modal_Standard_Adjudication.Reason3.value, _
                Re4:=Modal_Standard_Adjudication.Reason4.value, _
                Re5:=Modal_Standard_Adjudication.Reason5.value _
            )
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
            If .List(i, 4) = "New" Then 'if new service
                Call addSupervision( _
                        clientRow:=updateRow, _
                        serviceType:=.List(i, 0), _
                        legalStatus:=Standard_Return_Legal_Status.Caption, _
                        Courtroom:=oldCourtroom, _
                        DA:=DA.value, _
                        agency:=.List(i, 1), _
                        startDate:=.List(i, 2), _
                        NextCourtDate:=Standard_NextCourtDate.value, _
                        Re1:=.List(i, 6), _
                        Re2:=.List(i, 7), _
                        Re3:=.List(i, 8), _
                        Notes:=.List(i, 9))
            Else
                If Range(headerFind("Courtroom of Order", .List(i, 4)) & updateRow).value _
                        = Lookup("Courtroom_Name")("Intake Conf.") Then

                    Call dropSupervision( _
                            clientRow:=updateRow, _
                            head:=.List(i, 4), _
                            serviceType:=.List(i, 0), _
                            startDate:=.List(i, 2), _
                            endDate:=DateOfHearing.value, _
                            Nature:="Neutral", _
                            Re1:="N/A", _
                            Re2:="N/A", _
                            Re3:="N/A", _
                            Notes:="from intake conf.")

                    If Not IsDate(.List(i, 3)) Then
                        Call addSupervision( _
                                clientRow:=updateRow, _
                                serviceType:=.List(i, 0), _
                                legalStatus:=Standard_Return_Legal_Status.Caption, _
                                Courtroom:=oldCourtroom, _
                                DA:=DA.value, _
                                agency:=.List(i, 1), _
                                startDate:=DateOfHearing.value, _
                                NextCourtDate:=Standard_NextCourtDate.value, _
                                Re1:="N/A", _
                                Re2:="N/A", _
                                Re3:="N/A")
                    End If
                Else
                    If Range(headerFind("Courtroom of Order", .List(i, 4)) & updateRow).value _
                            = Lookup("Courtroom_Name")("PJJSC") Then

                        Call dropSupervision( _
                                clientRow:=updateRow, _
                                head:=.List(i, 4), _
                                serviceType:=.List(i, 0), _
                                startDate:=.List(i, 2), _
                                endDate:=DateOfHearing.value, _
                                Nature:="Neutral", _
                                Re1:="N/A", _
                                Re2:="N/A", _
                                Re3:="N/A", _
                                Notes:="from PJJSC")

                        If Not IsDate(.List(i, 3)) Then
                            Call addSupervision( _
                                    clientRow:=updateRow, _
                                    serviceType:=.List(i, 0), _
                                    legalStatus:=Standard_Return_Legal_Status.Caption, _
                                    Courtroom:=oldCourtroom, _
                                    DA:=DA.value, _
                                    agency:=.List(i, 1), _
                                    startDate:=DateOfHearing.value, _
                                    NextCourtDate:=Standard_NextCourtDate.value, _
                                    Re1:="N/A", _
                                    Re2:="N/A", _
                                    Re3:="N/A")
                        End If
                    Else
                        If IsDate(.List(i, 3)) Then 'if has End Date
                            Call dropSupervision( _
                                    clientRow:=updateRow, _
                                    head:=.List(i, 4), _
                                    serviceType:=.List(i, 0), _
                                    startDate:=.List(i, 2), _
                                    endDate:=.List(i, 3), _
                                    Nature:=.List(i, 5), _
                                    Re1:=.List(i, 6), _
                                    Re2:=.List(i, 7), _
                                    Re3:=.List(i, 8), _
                                    Notes:=.List(i, 9))
                        End If
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
            If .List(i, 4) = "New" Then 'if new service
                Call addCondition( _
                        clientRow:=updateRow, _
                        condition:=.List(i, 0), _
                        legalStatus:=Standard_Return_Legal_Status.Caption, _
                        Courtroom:=oldCourtroom, _
                        DA:=DA.value, _
                        agency:=.List(i, 1), _
                        startDate:=.List(i, 2), _
                        Re1:=.List(i, 6), _
                        Re2:=.List(i, 7), _
                        Re3:=.List(i, 8), _
                        Notes:=.List(i, 9))
            Else
                If Range(headerFind("Courtroom of Order", .List(i, 4)) & updateRow).value _
                        = Lookup("Courtroom_Name")("Intake Conf.") Then

                    Call dropCondition( _
                            clientRow:=updateRow, _
                            head:=.List(i, 4), _
                            condition:=.List(i, 0), _
                            startDate:=.List(i, 2), _
                            endDate:=DateOfHearing.value, _
                            Nature:="Neutral", _
                            Re1:="N/A", _
                            Re2:="N/A", _
                            Re3:="N/A", _
                            Notes:="Continued from intake conf.")

                    If Not IsDate(.List(i, 3)) Then
                        Call addCondition( _
                                clientRow:=updateRow, _
                                condition:=.List(i, 0), _
                                legalStatus:=Standard_Return_Legal_Status.Caption, _
                                Courtroom:=oldCourtroom, _
                                DA:=DA.value, _
                                agency:=.List(i, 1), _
                                startDate:=DateOfHearing.value, _
                                Re1:="N/A", _
                                Re2:="N/A", _
                                Re3:="N/A")
                    End If
                Else
                    If Range(headerFind("Courtroom of Order", .List(i, 4)) & updateRow).value _
                            = Lookup("Courtroom_Name")("PJJSC") Then
                        Call dropCondition( _
                                clientRow:=updateRow, _
                                head:=.List(i, 4), _
                                condition:=.List(i, 0), _
                                startDate:=.List(i, 2), _
                                endDate:=DateOfHearing.value, _
                                Nature:="Neutral", _
                                Re1:="N/A", _
                                Re2:="N/A", _
                                Re3:="N/A", _
                                Notes:="Continued from PJJSC")

                        If Not IsDate(.List(i, 3)) Then
                            Call addCondition( _
                                    clientRow:=updateRow, _
                                    condition:=.List(i, 0), _
                                    legalStatus:=Standard_Return_Legal_Status.Caption, _
                                    Courtroom:=oldCourtroom, _
                                    DA:=DA.value, _
                                    agency:=.List(i, 1), _
                                    startDate:=DateOfHearing.value, _
                                    Re1:="N/A", _
                                    Re2:="N/A", _
                                    Re3:="N/A")
                        End If
                    Else
                        If IsDate(.List(i, 3)) Then 'if has End Date
                            Call dropCondition( _
                                    clientRow:=updateRow, _
                                    head:=.List(i, 4), _
                                    condition:=.List(i, 0), _
                                    startDate:=.List(i, 2), _
                                    endDate:=.List(i, 3), _
                                    Nature:=.List(i, 5), _
                                    Re1:=.List(i, 6), _
                                    Re2:=.List(i, 7), _
                                    Re3:=.List(i, 8), _
                                    Notes:=.List(i, 9))
                        End If
                    End If
                End If
            End If
        Next i
    End With

    If Standard_Court_Transfer.BackColor = selectedColor Then
        Dim outcomeHead As String
        outcomeHead = headerFind("OUTCOMES", oldCourtHead)

        Range(headerFind("Date of Overall Discharge", outcomeHead) & updateRow).value = DateOfHearing.value
        Range(headerFind("Courtroom of Discharge", outcomeHead) & updateRow).value = Lookup("Courtroom_Name")(oldCourtroom)
        Range(headerFind("Legal Status of Discharge", outcomeHead) & updateRow).value = Lookup("Legal_Status_Name")(Standard_Fetch_Legal_Status.Caption)
        Range(headerFind("DA", outcomeHead) & updateRow).value = Lookup("DA_Last_Name_Name")(DA.value)
        Range(headerFind("Active or Discharged", outcomeHead) & updateRow).value = 2 'discharged
        Range(headerFind("Detailed Courtroom Outcome", outcomeHead) & updateRow).value _
                = Lookup("Detailed_Courtroom_Outcome_Name")(Modal_Standard_Court_Transfer.Detailed_Outcome.value)
        Call ReferClientTo( _
                referralDate:=DateOfHearing.value, _
                clientRow:=updateRow, _
                toCR:=Modal_Standard_Court_Transfer.Courtroom.value, _
                fromCR:=oldCourtroom)
    End If

    Call closeCallIn(DateOfHearing.value, updateRow)
    Call closeIntakeConference(DateOfHearing.value, updateRow)
    Call UnloadAll

    Call addNotes(oldCourtroom, DateOfHearing.value, updateRow, Standard_Notes, Standard_Fetch_Legal_Status.Caption)
    Call Save_Countdown

    Call CheckForConcurrency(updateRow, DateOfHearing.value)

    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
    End With
    Worksheets("User Entry").Activate


done:


    Exit Sub
err:
    Sheets("Entry").Range("C" & updateRow & ":" & hFind("END") & updateRow).value = restorer


    MsgBox "Something went wrong. Database has been restored to state prior to submission. " _
      & vbNewLine & vbNewLine & "Message: " & vbNewLine & err.Description _
      & vbNewLine & vbNewLine & "Source: " & vbNewLine & err.Source
    Call UnloadAll

    Stop   'press F8 twice to see the error point
    Resume
End Sub



Private Sub PJJSC_Submit_Click()
    Dim Num As Integer
    Dim bucketHead As String
    Dim detentionHead As String

    Worksheets("Entry").Activate
    Call addNotes(DRev_Facility.value, DateOfHearing.value, updateRow, "", Range(hFind("Legal Status") & updateRow).value)

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

    If DRev_DetentionDecision.value = "Held" _
        And DRev_HoldAndTransfer.value = "Yes" Then
        Call ReferClientTo( _
            referralDate:=DateOfHearing.value, _
            clientRow:=updateRow, _
            fromCR:="PJJSC", _
            toCR:=DRev_NextHearingLocation.value)

        Call addSupervision( _
            clientRow:=updateRow, _
            serviceType:="Detention (not respite)", _
            legalStatus:=Lookup("Legal_Status_Num")(Range(hFind("Legal Status") & updateRow).value), _
            Courtroom:="PJJSC", _
            DA:=DA.value, _
            agency:=DRev_Facility.value, _
            startDate:=DateOfHearing.value, _
            NextCourtDate:=PJJSC_NextCourtDate.value, _
            Re1:=ReasonForDetentionCommit1.value, _
            Re2:=ReasonForDetentionCommit2.value, _
            Re3:=ReasonForDetentionCommit3.value)
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
            toCR:=DRev_ReferredTo.value _
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
                Re1:=DRevSup1_Re1.value, _
                Re2:=DRevSup1_Re2.value, _
                Re3:=DRevSup1_Re3.value, _
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
                Re1:=DRevSup2_Re1.value, _
                Re2:=DRevSup2_Re2.value, _
                Re3:=DRevSup2_Re3.value, _
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
                Re1:="N/A", _
                Re2:="N/A", _
                Re3:="N/A", _
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
                Re1:="N/A", _
                Re2:="N/A", _
                Re3:="N/A", _
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
                Re1:="N/A", _
                Re2:="N/A", _
                Re3:="N/A", _
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

    Worksheets("User Entry").Activate
    Call Save_Countdown
    Call UnloadAll

End Sub

''''''''''''''''''''''''''''
''''''''''BUTTONS'''''''''''
''''''''''''''''''''''''''''

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

