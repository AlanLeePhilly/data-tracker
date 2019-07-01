VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DiversionUpdateForm 
   Caption         =   "DetentionReferral"
   ClientHeight    =   11700
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16440
   OleObjectBlob   =   "DiversionUpdateForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DiversionUpdateForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit




Private Sub FollowupResult_Change()
    Select Case FollowupResult
        Case "Rearrest"
            FollowupMP.value = 0
        Case "FTA - Breach"
            FollowupMP.value = 1
        Case "Recommended to Court"
            FollowupMP.value = 2
        Case Else
            FollowupMP.value = 0
    End Select
End Sub

Private Sub ReviewOpt_Click()
    FollowupMP.value = 1
End Sub

Private Sub ExitOpt_Click()
    FollowupMP.value = 2
End Sub

''''''''''''''''
'INITIALIZATION'
''''''''''''''''

Private Sub UserForm_Initialize()
    Call RefreshNamedRanges
    Call Generate_Dictionaries
    HearingMP.value = 0
    FirstHearingMP.value = 0
    FollowupMP.value = 0
    Me.ScrollTop = 0
End Sub

'''''''''''''
'VALIDATIONS'
'''''''''''''
Private Sub DateOfHearing_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Set ctl = Me.DateOfHearing

    Call DateValidation(ctl, Cancel)
End Sub

Private Sub HearingDateFill_Click()
    DateOfContract.value = DateOfHearing.value
End Sub
Private Sub FirstHearingNextCourtDate_Enter()
    FirstHearingNextCourtDate.value = CalendarForm.GetDate(RangeOfYears:=5)
End Sub

Private Sub FirstHearingNextCourtDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Set ctl = Me.FirstHearingNextCourtDate

    Call DateValidation(ctl, Cancel)
End Sub

Private Sub FollowupNextCourtDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Set ctl = Me.FollowupNextCourtDate

    Call DateValidation(ctl, Cancel)
End Sub

Private Sub DateOfContract_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Set ctl = Me.DateOfContract

    Call DateValidation(ctl, Cancel)
End Sub

Private Sub FirstTerm1_Change()
    Dim i As Integer

    If FirstTerm1.value = "None" Then
        FirstTerm1Provider.value = ""
        FirstTerm1Provider.Enabled = False
        FirstTerm2.value = ""
        FirstTerm2.Enabled = False
        FirstTerm2Provider.value = ""
        FirstTerm2Provider.Enabled = False
        FirstTerm3.value = ""
        FirstTerm3.Enabled = False
        FirstTerm3Provider.value = ""
        FirstTerm3Provider.Enabled = False
        FirstTerm4.value = ""
        FirstTerm4.Enabled = False
        FirstTerm4Provider.value = ""
        FirstTerm4Provider.Enabled = False
        FirstTerm5.value = ""
        FirstTerm5.Enabled = False
        FirstTerm5Provider.value = ""
        FirstTerm5Provider.Enabled = False
    Else
        FirstTerm1Provider.Enabled = True
        FirstTerm2.Enabled = True
    End If

End Sub

Private Sub FirstTerm2_Change()
    Dim i As Integer

    If FirstTerm2.value = "None" Then
        FirstTerm2Provider.value = ""
        FirstTerm2Provider.Enabled = False
        FirstTerm3.value = ""
        FirstTerm3.Enabled = False
        FirstTerm3Provider.value = ""
        FirstTerm3Provider.Enabled = False
        FirstTerm4.value = ""
        FirstTerm4.Enabled = False
        FirstTerm4Provider.value = ""
        FirstTerm4Provider.Enabled = False
        FirstTerm5.value = ""
        FirstTerm5.Enabled = False
        FirstTerm5Provider.value = ""
        FirstTerm5Provider.Enabled = False
    Else
        FirstTerm2Provider.Enabled = True
        FirstTerm3.Enabled = True
    End If
End Sub

Private Sub FirstTerm3_Change()
    Dim i As Integer

    If FirstTerm3.value = "None" Then
        FirstTerm3Provider.value = ""
        FirstTerm3Provider.Enabled = False
        FirstTerm4.value = ""
        FirstTerm4.Enabled = False
        FirstTerm4Provider.value = ""
        FirstTerm4Provider.Enabled = False
        FirstTerm5.value = ""
        FirstTerm5.Enabled = False
        FirstTerm5Provider.value = ""
        FirstTerm5Provider.Enabled = False
    Else
        FirstTerm3Provider.Enabled = True
        FirstTerm4.Enabled = True
    End If
End Sub

Private Sub FirstTerm4_Change()
    Dim i As Integer

    If FirstTerm4.value = "None" Then
        FirstTerm4Provider.value = ""
        FirstTerm4Provider.Enabled = False
        FirstTerm5.value = ""
        FirstTerm5.Enabled = False
        FirstTerm5Provider.value = ""
        FirstTerm5Provider.Enabled = False
    Else
        FirstTerm4Provider.Enabled = True
        FirstTerm5.Enabled = True
    End If
End Sub

Private Sub FirstTerm5_Change()
    Dim i As Integer

    If FirstTerm5.value = "None" Then
        FirstTerm5Provider.value = ""
        FirstTerm5Provider.Enabled = False
    Else
        FirstTerm5Provider.Enabled = True
    End If
End Sub


Private Sub SearchButton_Click()
    On Error Resume Next

    'define variable Long(a big integer) named emptyRow
    Dim lastRow As Long
    Dim Query As String
    Dim lookCell As String
    Dim lookRow As Long

    'activate the spreadsheet as default selector
    With Worksheets("Entry")

        'define variable of search query in UPPERCASE named 'query'
        Query = UCase(SearchTextBox.value)

        SearchResultsBox.Clear

        lastRow = .Range("C" & Rows.count).End(xlUp).row

        For lookRow = 3 To lastRow

            lookCell = UCase(.Range(headerFind(Search_Type.value) & lookRow))

            If InStr(1, lookCell, Query) > 0 Then
                With SearchResultsBox
                    .ColumnCount = 4
                    .AddItem lookRow
                    .List(SearchResultsBox.ListCount - 1, 1) = Worksheets("Entry").Range(headerFind("First Name") & lookRow)
                    .List(SearchResultsBox.ListCount - 1, 2) = Worksheets("Entry").Range(headerFind("Last Name") & lookRow)
                    .List(SearchResultsBox.ListCount - 1, 3) = Worksheets("Entry").Range(headerFind("Arrest Date") & lookRow)
                End With
            End If
        Next lookRow
    End With
End Sub
Private Sub SearchResultsBox_Click()
    updateRow = SearchResultsBox.value
End Sub
Private Sub LockSearch_Click()
    SearchResultsBox.Locked = True
    SearchResultsBox.Enabled = False
    SearchButton.Enabled = False
    SearchTextBox.Enabled = False
End Sub
Private Sub UnlockSearch_Click()
    SearchResultsBox.Locked = False
    SearchResultsBox.Enabled = True
    SearchButton.Enabled = True
    SearchTextBox.Enabled = True
    SearchResultsBox.value = ""
End Sub
Private Sub Cancel_Click()
    Call Clear_Click
    Unload Me
End Sub
Private Sub Yesterday_Click()
    DateOfHearing.value = DateAdd("d", -1, Date)
End Sub

Private Sub Today_Click()
    DateOfHearing.value = Date
End Sub
Private Sub Clear_Click()
    Call Clear_Form(Me)
    Call UnlockSearch_Click
End Sub

Private Sub LookupData_Click()
    'Dim courtHead As String
    Dim diversionHead As String
    Worksheets("Entry").Activate
    
    'courtHead = headerFind(ReferredTo.value)
    diversionHead = headerFind("DIVERSION")

    If DateOfHearing.value = "" Then
        MsgBox "Date of Hearing Required for Lookup"
        Exit Sub
    End If
    
    FetchTerms.Clear
    ReturnTerms.Clear

    Referral_Date.Caption _
            = Range(headerFind("Referral Date", diversionHead) & updateRow).value
    
    Referral_Source.Caption _
        = Lookup("Diversion_Referral_Source_Num")(Range(headerFind("Referral Source", diversionHead) & updateRow).value)

    Select Case True
        Case isEmptyOrZero(Range(headerFind("Date of First Hearing", diversionHead) & updateRow))
            Status.Caption = "Referred"
            HearingMP.value = 1

        Case Range(headerFind("Outcomes of First Hearing", diversionHead) & updateRow).value _
                = Lookup("YAP_First_Hearing_Outcome_Name")("FTA - Continue")
            Status.Caption = "Referred"
            HearingMP.value = 1

        Case Range(headerFind("Outcomes of First Hearing", diversionHead) & updateRow).value _
                = Lookup("YAP_First_Hearing_Outcome_Name")("Contract Received")
            Status.Caption = "Contract Granted"
            HearingMP.value = 2

        Case Range(headerFind("Outcomes of First Hearing", diversionHead) & updateRow).value = 13
            Status.Caption = "Recommended to Court"
        Case Range(headerFind("Outcomes of First Hearing", diversionHead) & updateRow).value = 98
            Status.Caption = "Other"
        Case Range(headerFind("Outcomes of First Hearing", diversionHead) & updateRow).value = 99
            Status.Caption = "Unknown"
    End Select

    FetchMonitorFirst = Range(headerFind("Monitor First Name", diversionHead) & updateRow)
    FetchMonitorLast = Range(headerFind("Monitor Last Name", diversionHead) & updateRow)
    FetchVictimFirst = Range(headerFind("Victim First Name", diversionHead) & updateRow)
    FetchVictimLast = Range(headerFind("Victim Last Name", diversionHead) & updateRow)
    FetchYAPPanel.Caption = Lookup("Police_District_Num")(Range(headerFind("YAP Panel District #", diversionHead) & updateRow).value)

    If Status.Caption = "Contract Granted" Then
        Dim i As Long
        Dim j As String
        Dim k As String
        Dim l As Integer
        Dim m As String
        For i = 1 To 5
            k = Lookup("Condition_Num")(getContractTerm(updateRow, i))
            j = Lookup("Condition_Provider_Num")(getContractTermProvider(updateRow, i))
            m = getContractTermDate(updateRow, i)
            l = DateDiff("d", getContractTermDate(updateRow, i), Date)
            If Not k = "None" Then
                With FetchTerms
                    .ColumnCount = 4
                    .ColumnWidths = "0;70;70;30;"
                    .AddItem i
                    .List(FetchTerms.ListCount - 1, 1) = Lookup("Condition_Num")(getContractTerm(updateRow, i))
                    .List(FetchTerms.ListCount - 1, 2) = Lookup("Condition_Provider_Num")(getContractTermProvider(updateRow, i))
                    .List(FetchTerms.ListCount - 1, 3) = DateDiff("d", getContractTermDate(updateRow, i), DateOfHearing)
                End With
                With ReturnTerms
                    .ColumnCount = 5
                    .ColumnWidths = "0;70;70;0;0" 'term number, term, provider, LOS, edit flag
                    .AddItem i
                    .List(ReturnTerms.ListCount - 1, 1) = Lookup("Condition_Num")(getContractTerm(updateRow, i))
                    .List(ReturnTerms.ListCount - 1, 2) = Lookup("Condition_Provider_Num")(getContractTermProvider(updateRow, i))
                    .List(ReturnTerms.ListCount - 1, 3) = DateDiff("d", getContractTermDate(updateRow, i), DateOfHearing)
                End With
            End If
        Next i
    End If
    
    Worksheets("User Entry").Activate
End Sub

Private Sub FirstHearingResult_Change()
    Select Case FirstHearingResult.value
        Case "Contract Received"
            FirstHearingMP.value = 1
        Case "FTA - Breach"
            FirstHearingMP.value = 3
        Case "FTA - Continue"
            FirstHearingMP.value = 0
        Case "Recommended to Court"
            FirstHearingMP.value = 2
        Case "Unknown"
            FirstHearingMP.value = 0
        Case "Other"
            FirstHearingMP.value = 0
    End Select
End Sub

Private Sub EditTerms_Click()
    Dim i As Integer

    For i = 0 To ReturnTerms.ListCount - 1
        With Modal_Diversion_Term.EditTerms
            .ColumnCount = 4
            .ColumnWidths = "10;70;70;0;"
            .AddItem
            .List(i, 0) = ReturnTerms.List(i, 0)
            .List(i, 1) = ReturnTerms.List(i, 1)
            .List(i, 2) = ReturnTerms.List(i, 2)
            .List(i, 3) = ReturnTerms.List(i, 3)
        End With
    Next i
    Modal_Diversion_Term.Show
End Sub

Private Sub FirstHearingSubmit_Click()

    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With

    Worksheets("Entry").Activate
    'Dim courtHead As String
    Dim diversionHead As String
    'courtHead = headerFind(ReferredTo.value)
    diversionHead = headerFind("DIVERSION")

    'ADD FRONT NEXT COURT DATE & BUMP TO PREVIOUS

    Range(headerFind("Next Court Date") & updateRow).value _
        = FirstHearingNextCourtDate.value

    Range(headerFind("Outcomes of First Hearing", diversionHead) & updateRow).value _
        = Lookup("YAP_First_Hearing_Outcome_Name")(FirstHearingResult.value)

    If FirstHearingResult.value = "Contract Received" Then
        Range(headerFind("Date of Contract", diversionHead) & updateRow).value = DateOfContract.value
        Range(headerFind("Date of First Hearing", diversionHead) & updateRow).value = DateOfHearing.value
        Range(headerFind("Projected Completion Date", diversionHead) & updateRow).value = FirstHearingNextCourtDate.value

        Range(headerFind("Monitor First Name", diversionHead) & updateRow).value = MonitorFirstName.value
        Range(headerFind("Monitor Last Name", diversionHead) & updateRow).value = MonitorLastName.value

        Range(headerFind("Contract Term #1", diversionHead) & updateRow).value = Lookup("Condition_Name")(FirstTerm1.value)
        Range(headerFind("Contract Term #1 Provider", diversionHead) & updateRow).value = Lookup("Condition_Provider_Name")(FirstTerm1Provider.value)
        Range(headerFind("Contract Term #2", diversionHead) & updateRow).value = Lookup("Condition_Name")(FirstTerm2.value)
        Range(headerFind("Contract Term #2 Provider", diversionHead) & updateRow).value = Lookup("Condition_Provider_Name")(FirstTerm2Provider.value)
        Range(headerFind("Contract Term #3", diversionHead) & updateRow).value = Lookup("Condition_Name")(FirstTerm3.value)
        Range(headerFind("Contract Term #3 Provider", diversionHead) & updateRow).value = Lookup("Condition_Provider_Name")(FirstTerm3Provider.value)
        Range(headerFind("Contract Term #4", diversionHead) & updateRow).value = Lookup("Condition_Name")(FirstTerm4.value)
        Range(headerFind("Contract Term #4 Provider", diversionHead) & updateRow).value = Lookup("Condition_Provider_Name")(FirstTerm4Provider.value)
        Range(headerFind("Contract Term #5", diversionHead) & updateRow).value = Lookup("Condition_Name")(FirstTerm5.value)
        Range(headerFind("Contract Term #5 Provider", diversionHead) & updateRow).value = Lookup("Condition_Provider_Name")(FirstTerm5Provider.value)
    End If

    If FirstHearingResult.value = "FTA - Breach" Then
        Range(headerFind("Reason #1 Recommended to Court", diversionHead) & updateRow).value _
            = Lookup("Diversion_Court_Recommendation_Reason_Name")("FTA")
        Range(headerFind("Date of First Hearing", diversionHead) & updateRow).value = DateOfHearing.value
        Range(headerFind("Courtroom of Transfer", diversionHead) & updateRow).value _
            = Lookup("Courtroom_Name")(BreachCourtroom.value)

        Range(headerFind("Nature of Discharge", diversionHead) & updateRow).value = 2 'negative
        Range(headerFind("Discharge Date", diversionHead) & updateRow).value = DateOfHearing.value
        Range(headerFind("LOS Diversion", diversionHead) & updateRow).value _
            = DateDiff("d", Range(headerFind("Arrest Date") & updateRow).value, DateOfHearing.value)
        Range(headerFind("Detailed YAP Outcome", diversionHead) & updateRow).value = 2 'FTA
        Call ReferClientTo( _
            referralDate:=DateOfHearing.value, _
            clientRow:=updateRow, _
            toCR:=BreachCourtroom.value _
            )
        Call addFTA( _
            updateRow, _
            DateOfHearing.value, _
            Lookup("Courtroom_Num")(Range(hFind("Active Courtroom") & updateRow).value), _
            "Diversion")
    End If

    If FirstHearingResult.value = "FTA - Continue" Then
        Call addFTA( _
            updateRow, _
            DateOfHearing.value, _
            Lookup("Courtroom_Num")(Range(hFind("Active Courtroom") & updateRow).value), _
            "Diversion")
    End If

    If FirstHearingResult.value = "Recommended to Court" Then
        Range(headerFind("Reason #1 Recommended to Court", diversionHead) & updateRow).value _
            = Lookup("Diversion_Court_Recommendation_Reason_Name")(RecToCourtReason1.value)
        Range(headerFind("Reason #2 Recommended to Court", diversionHead) & updateRow).value _
            = Lookup("Diversion_Court_Recommendation_Reason_Name")(RecToCourtReason2.value)
        Range(headerFind("Reason #3 Recommended to Court", diversionHead) & updateRow).value _
            = Lookup("Diversion_Court_Recommendation_Reason_Name")(RecToCourtReason3.value)
        Range(headerFind("Reason #4 Recommended to Court", diversionHead) & updateRow).value _
            = Lookup("Diversion_Court_Recommendation_Reason_Name")(RecToCourtReason4.value)
        Range(headerFind("Reason #5 Recommended to Court", diversionHead) & updateRow).value _
            = Lookup("Diversion_Court_Recommendation_Reason_Name")(RecToCourtReason5.value)

        Range(headerFind("Courtroom of Transfer", diversionHead) & updateRow).value _
            = Lookup("Courtroom")(CourtroomReferredTo.value)
        Range(headerFind("Nature of Discharge", diversionHead) & updateRow).value = 2 'negative
        Range(headerFind("Discharge Date", diversionHead) & updateRow).value = DateOfHearing.value
        Range(headerFind("LOS Diversion", diversionHead) & updateRow).value _
            = DateDiff("d", Range(headerFind("Arrest Date") & updateRow).value, DateOfHearing.value)
        Range(headerFind("Detailed YAP Outcome", diversionHead) & updateRow).value = 14 'Admittance Not Granted
        'TODO ReferTo
    End If


    Unload DiversionUpdateForm


    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
    End With

End Sub
Private Sub FollowupSubmit_Click()
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With

    Worksheets("Entry").Activate
    'Dim courtHead As String
    Dim diversionHead As String
    Dim hearingType As String
    'courtHead = headerFind(ReferredTo.value)
    diversionHead = headerFind("DIVERSION")

    'ADD FRONT NEXT COURT DATE & BUMP TO PREVIOUS
    Range(headerFind("Next Court Date") & updateRow).value _
            = FirstHearingNextCourtDate.value

    If ReviewOpt = False Then
        If ExitOpt = False Then
            MsgBox "Please Select Type of Hearing"
            Exit Sub
        Else
            hearingType = "Exit"
            Range(headerFind("Did Youth Receive an Exit Hearing?", diversionHead) & updateRow).value = 1
        End If
    Else
        hearingType = "Review"
        Range(headerFind("Did Youth Receive a Review Hearing?", diversionHead) & updateRow).value = 1
    End If

    Dim i As Integer

    If IsEmpty(Range(headerFind("Date of " & hearingType & " Hearing #1", diversionHead) & updateRow).value) Then
        i = 1
    Else
        If IsEmpty(Range(headerFind("Date of " & hearingType & " Hearing #2", diversionHead) & updateRow).value) Then
            i = 2
        Else
            If IsEmpty(Range(headerFind("Date of " & hearingType & " Hearing #3", diversionHead) & updateRow).value) Then
                i = 3
            Else
                MsgBox "Already 3 " & hearingType & " hearings on record. See administrator for more information."
                Exit Sub
            End If
        End If
    End If
    '2

    Range(headerFind("Date of " & hearingType & " Hearing #" & i, diversionHead) & updateRow).value = DateOfHearing.value
    Range(headerFind("Result of " & hearingType & " Hearing #" & i, diversionHead) & updateRow).value = Lookup("Diversion_Review_Hearing_Outcome_Name")(FollowupResult.value)

    Dim j As Integer

    For j = 0 To FetchTerms.ListCount - 1
        If FetchTerms.List(j, 1) = ReturnTerms.List(j, 1) And FetchTerms.List(j, 2) = ReturnTerms.List(j, 2) Then
        Else
            If getOpenTermEdit(updateRow) = 0 Then
                MsgBox "Already 5 term edits. See administrator for more information"
                Exit Sub
            Else
                tempHead = headerFind("Which Term Was Updated #" & getOpenTermEdit(updateRow), diversionHead)
                Range(headerFind("Which Term Was Updated #" & getOpenTermEdit(updateRow), diversionHead) & updateRow) = FetchTerms.List(j, 0)
                Range(headerFind("Date of Update", tempHead) & updateRow) = DateOfHearing
                Range(headerFind("Previous Term", tempHead) & updateRow) = Lookup("Condition_Name")(FetchTerms.List(j, 1))
                Range(headerFind("New Term", tempHead) & updateRow) = Lookup("Condition_Name")(ReturnTerms.List(j, 1))
                Range(headerFind("New Term Provider", tempHead) & updateRow) = Lookup("Condition_Provider_Name")(ReturnTerms.List(j, 2))
            End If
        End If
    Next j

    Select Case FollowupResult
        Case "Rearrest"
            'TODO Rearrest
        Case "FTA - Breach"
            Range(headerFind("Reason #1 Recommended to Court", diversionHead) & updateRow).value _
                    = Lookup("Diversion_Court_Recommendation_Reason_Name")("FTA")
            Range(headerFind("Courtroom of Transfer", diversionHead) & updateRow).value _
                    = Lookup("Courtroom_Name")(FollowupBreachCourt.value)

            Range(headerFind("Nature of Discharge", diversionHead) & updateRow).value = 2 'negative
            Range(headerFind("Discharge Date", diversionHead) & updateRow).value = DateOfHearing.value
            Range(headerFind("LOS Diversion", diversionHead) & updateRow).value _
                    = DateDiff("d", Range(headerFind("Arrest Date") & updateRow).value, DateOfHearing.value)
            Range(headerFind("Detailed YAP Outcome", diversionHead) & updateRow).value = 2 'FTA
            Call ReferClientTo( _
                    referralDate:=DateOfHearing.value, _
                    clientRow:=updateRow, _
                    toCR:=FollowupBreachCourt.value _
                    )
            Call addFTA( _
                    updateRow, _
                    DateOfHearing.value, _
                    Lookup("Courtroom_Num")(Range(hFind("Active Courtroom") & updateRow).value), _
                    "Diversion")
        Case "FTA - Continue"
            Call addFTA( _
                    updateRow, _
                    DateOfHearing.value, _
                    Lookup("Courtroom_Num")(Range(hFind("Active Courtroom") & updateRow).value), _
                    "Diversion")
        Case "Recommended to Court"
            Range(headerFind("Reason #1 Recommended to Court", diversionHead) & updateRow).value _
                    = Lookup("Diversion_Court_Recommendation_Reason_Name")(FollowupRecReason.value)
            Range(headerFind("Reasons Recommended to Court #" & i, headerFind(hearingType, diversionHead)) & updateRow).value _
                    = Lookup("Diversion_Court_Recommendation_Reason_Name")(FollowupRecReason.value)
            Range(headerFind("Courtroom of Transfer", diversionHead) & updateRow).value _
                    = Lookup("Courtroom_Name")(FollowupCourtroom.value)
            Call ReferClientTo( _
                    referralDate:=DateOfHearing.value, _
                    clientRow:=updateRow, _
                    toCR:=FollowupCourtroom.value _
                    )
            Range(headerFind("Nature of Discharge", diversionHead) & updateRow).value = 2 'negative
            Range(headerFind("Discharge Date", diversionHead) & updateRow).value = DateOfHearing.value
            Range(headerFind("LOS Diversion", diversionHead) & updateRow).value _
                    = DateDiff("d", Range(headerFind("Arrest Date") & updateRow).value, DateOfHearing.value) / 365
            Range(headerFind("Detailed YAP Outcome", diversionHead) & updateRow).value = 13 'Contract Breach

        Case "Pending Completion - Orginal Timeline"
        Case "Pending Completion - Extension"
            Range(headerFind("Projected Completion Date", diversionHead) & updateRow).value = DateOfHearing.value
        Case "Positive Completion"
            Range(headerFind("Nature of Discharge", diversionHead) & updateRow).value = 1 'positive
            Range(headerFind("Discharge Date", diversionHead) & updateRow).value = DateOfHearing.value
            Range(headerFind("LOS Diversion", diversionHead) & updateRow).value _
                    = DateDiff("d", Range(headerFind("Arrest Date") & updateRow).value, DateOfHearing.value) / 365
            Range(headerFind("Detailed YAP Outcome", diversionHead) & updateRow).value = 6 ' positive completion
            Call totalOutcome( _
                    updateRow, _
                    DateOfHearing.value, _
                    Lookup("Courtroom_Num")(Range(hFind("Active Courtroom") & updateRow).value), _
                    FollowupDA.value, _
                    "Diversion", _
                    "Positive", _
                    "Petition Diverted & Withdrawn", _
                    FollowupNotes.value)


    End Select

    'zerofill?


    Unload DiversionUpdateForm


    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
    End With
End Sub


