VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CourtroomReferral 
   Caption         =   "UserForm1"
   ClientHeight    =   8835.001
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   9408.001
   OleObjectBlob   =   "CourtroomReferral.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CourtroomReferral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


''''''''''''''''
'INITIALIZATION'
''''''''''''''''

Private Sub UserForm_Initialize()
    Call Generate_Dictionaries
End Sub

'''''''''''''
'VALIDATIONS'
'''''''''''''

Private Sub DateOfReferral_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Dim ctl As Control
    Set ctl = Me.DateOfReferral

    Call DateValidation(ctl, Cancel)
End Sub

Private Sub DateOfNextHearing_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Dim ctl As Control
    Set ctl = Me.DateOfNextHearing

    Call DateValidation(ctl, Cancel)
End Sub

Private Sub SearchButton_Click()
    On Error Resume Next

    'define variable Long(a big integer) named emptyRow
    Dim lastRow As Long
    Dim Query As String
    Dim lookCell As String
    Dim lookRow As Long

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
                .ColumnCount = 5
                .AddItem lookRow
                .List(SearchResultsBox.ListCount - 1, 1) = _
                        Range(headerFind("First Name") & lookRow)
                .List(SearchResultsBox.ListCount - 1, 2) = _
                        Range(headerFind("Last Name") & lookRow)
                .List(SearchResultsBox.ListCount - 1, 3) = _
                        Range(headerFind("Arrest Date") & lookRow)
                .List(SearchResultsBox.ListCount - 1, 4) = _
                        Lookup("Courtroom_Num")(Range(headerFind("Active Courtroom") & lookRow).value)
            End With
        End If
    Next lookRow
End Sub
Private Sub SearchResultsBox_Click()
    updateRow = SearchResultsBox.value
    ReferredFrom.value = SearchResultsBox.List(SearchResultsBox.listIndex, 4)
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
Private Sub Yesterday_Click()
    DateOfReferral.value = DateAdd("d", -1, Date)
End Sub

Private Sub Today_Click()
    DateOfReferral.value = Date
End Sub

Private Sub Cancel_Click()
    Call Clear_Click
    Unload Me
End Sub

Private Sub Clear_Click()
    Call Clear_Form(Me)
    Call UnlockSearch_Click
End Sub

Private Sub Submit_Click()
    Dim restorer As Variant

    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With

    Worksheets("Entry").Activate
    Dim fromHead As String
    Dim toHead As String

    If IsEmpty(ReferredTo.value) Then
        MsgBox "Courtroom Referred To is required"
        Exit Sub
    End If

    If IsEmpty(DateOfReferral.value) Then
        MsgBox "Date of Referral Required"
        Exit Sub
    End If

    restorer = Range("C" & updateRow & ":" & hFind("END") & updateRow).value
    On Error GoTo err



    'universal
    Range(headerFind("Next Court Date") & updateRow).value _
            = DateOfNextHearing.value



    Call prepend(Range(headerFind("Previous Court Dates") & updateRow), DateOfNextHearing.value)

    Call ReferClientTo( _
            referralDate:=DateOfReferral.value, _
            clientRow:=updateRow, _
            toCR:=ReferredTo.value, _
            fromCR:=ReferredFrom.value, _
            Notes:=Notes.value _
            )

done:
    Worksheets("User Entry").Activate
    Unload Me


    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
    End With
    Exit Sub

err:
    Range("C" & updateRow & ":" & hFind("END") & updateRow).value = restorer
    MsgBox "Something went wrong. Database has been restored to state prior to submission. " _
      & vbNewLine & vbNewLine & "Message: " & vbNewLine & err.Description _
      & vbNewLine & vbNewLine & "Source: " & vbNewLine & err.Source

    Unload Me
End Sub
