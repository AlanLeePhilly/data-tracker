VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ClientEdit 
   Caption         =   "ClientEdit"
   ClientHeight    =   11100
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14520
   OleObjectBlob   =   "ClientEdit.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ClientEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub DB_Section_Change()
    DB_Subsection.value = "All"
    
    Select Case DB_Section.value
        Case "Demographics", "Petition", "DRAI", "Detention"
            DB_Subsection.RowSource = "DB_Subsection_All"
        Case "Detention (VOP)"
            DB_Subsection.RowSource = "DB_Subsection_VOP"
            DB_Subsection.Enabled = True
        Case "Diversion"
            DB_Subsection.RowSource = "DB_Subsection_Diversion"
            DB_Subsection.Enabled = True
        Case "4G", "4E", "6F", "6H", "3E", "Crossover", "WRAP"
            DB_Subsection.RowSource = "DB_Subsection_Courtroom"
            DB_Subsection.Enabled = True
        Case "JTC"
            DB_Subsection.RowSource = "DB_Subsection_JTC"
            DB_Subsection.Enabled = True
        Case "Adult"
            DB_Subsection.RowSource = "DB_Subsection_Adult"
            DB_Subsection.Enabled = True
        Case "Aggregates"
            DB_Subsection.RowSource = "DB_Subsection_Aggregate"
            DB_Subsection.Enabled = True
        Case Else
            MsgBox "The section list has changed and the code was not updated to reflect that. Contact system admin."
    End Select
End Sub

Private Sub Edit_Click()
    Dim ColAlpha As String
    Dim HeaderName As String
    Dim CellVal As String
    
    ColAlpha = ClientEdit.LookupBox.List(ClientEdit.LookupBox.listIndex, 1)
    HeaderName = ClientEdit.LookupBox.List(ClientEdit.LookupBox.listIndex, 2)
    CellVal = ClientEdit.LookupBox.List(ClientEdit.LookupBox.listIndex, 3)
    
    Modal_Client_Edit.Column_Name.Caption = HeaderName
    Modal_Client_Edit.Column_Value.Caption = CellVal
    If (Not IsEmpty(Range(ColAlpha + "1"))) Then
        
        Modal_Client_Edit.New_Value_Box.RowSource = Range(ColAlpha + "1")
        Modal_Client_Edit.New_Value_Text.Visible = False
        Modal_Client_Edit.New_Value_Box.Visible = True
    Else
        
        Modal_Client_Edit.New_Value_Text.Visible = True
        Modal_Client_Edit.New_Value_Box.Visible = False
    End If
        
    Modal_Client_Edit.Show
End Sub

Private Sub Remove_Click()
    If UpdateBox.listIndex > -1 Then
        UpdateBox.RemoveItem (UpdateBox.listIndex)
    End If
End Sub

Private Sub SearchResultsBox_Click()
    updateRow = SearchResultsBox.value
End Sub

Private Sub Lookup_Button_Click()
    Call Generate_Dictionaries
    Dim startCol As Long
    Dim endCol As Long
    Dim count As Long
    
    LookupBox.Clear
    
    Select Case DB_Section
        Case "Demographics"
            startCol = alphaToNum(headerFind("DEMOGRAPHICS")) + 1
            endCol = alphaToNum(headerFind("PETITION")) - 1
        Case "Petition"
            startCol = alphaToNum(headerFind("PETITION")) + 1
            endCol = alphaToNum(headerFind("INTAKE CONFERENCE")) - 1
        Case "Intake Conference"
            startCol = alphaToNum(headerFind("INTAKE CONFERENCE")) + 1
            endCol = alphaToNum(headerFind("DETENTION")) - 1
        Case "Detention"
            startCol = alphaToNum(headerFind("DETENTION")) + 1
            endCol = alphaToNum(headerFind("DETENTION (VOP)")) - 1
        Case "Detention (VOP)"
            courtHead = headerFind("DETENTION (VOP)")
            Select Case DB_Subsection
                Case "All"
                    startCol = alphaToNum(courtHead) + 1
                    endCol = alphaToNum(headerFind("DIVERSION", courtHead)) - 1
                Case "Hearing #1"
                    startCol = alphaToNum(courtHead) + 1
                    endCol = alphaToNum(headerFind("Date of New Detention Hearing #2", courtHead)) - 1
                Case "Hearing #2"
                    startCol = alphaToNum(headerFind("Date of New Detention Hearing #2", courtHead))
                    endCol = alphaToNum(headerFind("Date of New Detention Hearing #3", courtHead)) - 1
                Case "Hearing #3"
                    startCol = alphaToNum(headerFind("Date of New Detention Hearing #3", courtHead))
                    endCol = alphaToNum(headerFind("DIVERSION", courtHead)) - 1
                Case Else
                    MsgBox "The section list has changed and the code was not updated to reflect that." _
                        + vbNewLine _
                        + "You were looking for:" _
                        + "Section: " + DB_Section.value _
                        + "Sub-Section: " + DB_Subsection.value _
                        + "Contact system admin."
            End Select
        Case "Diversion"
            courtHead = headerFind("DIVERSION")
            Select Case DB_Subsection
                Case "All"
                    startCol = alphaToNum(courtHead) + 1
                    endCol = alphaToNum(headerFind("4G", courtHead)) - 1
                Case "General"
                    startCol = alphaToNum(courtHead) + 1
                    endCol = alphaToNum(headerFind("YAP", courtHead)) - 1
                Case "YAP"
                    startCol = alphaToNum(headerFind("YAP", courtHead)) + 1
                    endCol = alphaToNum(headerFind("YAP Hearings", courtHead)) - 1
                Case "Review Hearings"
                    startCol = alphaToNum(headerFind("Review", courtHead))
                    endCol = alphaToNum(headerFind("Exit", courtHead)) - 1
                Case "Exit Hearings"
                    startCol = alphaToNum(headerFind("Exit", courtHead))
                    endCol = alphaToNum(headerFind("YAP CONTRACT EDITS", courtHead)) - 1
                Case "Contract Edits"
                    startCol = alphaToNum(headerFind("YAP CONTRACT EDITS", courtHead))
                    endCol = alphaToNum(headerFind("OUTCOMES", courtHead)) - 1
                Case "Outcome"
                    startCol = alphaToNum(headerFind("OUTCOMES", courtHead))
                    endCol = alphaToNum(headerFind("4G", courtHead)) - 1
                Case Else
                    MsgBox "The section list has changed and the code was not updated to reflect that." _
                        + vbNewLine _
                        + "You were looking for:" _
                        + "Section: " + DB_Section.value _
                        + "Sub-Section: " + DB_Subsection.value _
                        + "Contact system admin."
            End Select
        Case "4G", "4E", "6F", "6H", "3E", "Crossover", "WRAP"
            courtHead = headerFind(DB_Section.value)
            Select Case DB_Subsection.value
                Case "All"
                    startCol = alphaToNum(courtHead) + 1
                    endCol = alphaToNum(headerFind("JTC", courtHead)) - 1
                Case "General"
                    startCol = alphaToNum(courtHead) + 1
                    endCol = alphaToNum(headerFind("LEGAL STATUS", courtHead)) - 1
                Case "Legal Status"
                    startCol = alphaToNum(headerFind("LEGAL STATUS", courtHead)) + 1
                    endCol = alphaToNum(headerFind("OUTCOMES", courtHead)) - 1
                Case "Outcomes"
                    startCol = alphaToNum(headerFind("OUTCOMES", courtHead)) + 1
                    endCol = alphaToNum(headerFind("COURT PROCEEDINGS", courtHead)) - 1
                Case "Court Proceedings"
                    startCol = alphaToNum(headerFind("COURT PROCEEDINGS", courtHead)) + 1
                    endCol = alphaToNum(headerFind("Admissions", courtHead)) - 1
                Case "Admissions"
                    startCol = alphaToNum(headerFind("Admissions", courtHead)) + 1
                    endCol = alphaToNum(headerFind("Adjudications", courtHead)) - 1
                Case "Adjudications"
                    startCol = alphaToNum(headerFind("Adjudications", courtHead)) + 1
                    endCol = alphaToNum(headerFind("Continuances", courtHead)) - 1
                Case "Continuances"
                    startCol = alphaToNum(headerFind("Continuances", courtHead)) + 1
                    endCol = alphaToNum(headerFind("Placements", courtHead)) - 1
                Case "Placements"
                    startCol = alphaToNum(headerFind("Placements", courtHead)) + 1
                    endCol = alphaToNum(headerFind("Supervision Programs", courtHead)) - 1
                Case "Supervision Programs"
                    startCol = alphaToNum(headerFind("Supervision Programs", courtHead)) + 1
                    endCol = alphaToNum(headerFind("Conditions", courtHead)) - 1
                Case "Conditions"
                    startCol = alphaToNum(headerFind("Conditions", courtHead)) + 1
                    endCol = alphaToNum(headerFind("JTC", courtHead)) - 1
                Case Else
                    MsgBox "The section list has changed and the code was not updated to reflect that." _
                        + vbNewLine _
                        + "You were looking for:" _
                        + "Section: " + DB_Section.value _
                        + "Sub-Section: " + DB_Subsection.value _
                        + "Contact system admin."
            End Select
        Case "JTC"
            courtHead = headerFind("JTC")
            
            Select Case DB_Subsection.value
                Case "All"
                    startCol = alphaToNum(courtHead) + 1
                    endCol = alphaToNum(headerFind("ADULT", courtHead)) - 1
                Case "General"
                    startCol = alphaToNum(courtHead) + 1
                    endCol = alphaToNum(headerFind("PHASE 1", courtHead)) - 1
                Case "Phase #1"
                    startCol = alphaToNum(headerFind("PHASE 1", courtHead)) + 1
                    endCol = alphaToNum(headerFind("PHASE 2", courtHead)) - 1
                Case "Phase #2"
                    startCol = alphaToNum(headerFind("PHASE 2", courtHead)) + 1
                    endCol = alphaToNum(headerFind("PHASE 3", courtHead)) - 1
                Case "Phase #3"
                    startCol = alphaToNum(headerFind("PHASE 3", courtHead)) + 1
                    endCol = alphaToNum(headerFind("Placements", courtHead)) - 1
                Case "Placement"
                    startCol = alphaToNum(headerFind("Placements", courtHead)) + 1
                    endCol = alphaToNum(headerFind("JTC OUTCOMES", courtHead)) - 1
                Case "Outcome"
                    startCol = alphaToNum(headerFind("JTC OUTCOMES", courtHead)) + 1
                    endCol = alphaToNum(headerFind("Conditions", courtHead)) - 1
                Case "Supervision Programs"
                    startCol = alphaToNum(headerFind("Supervision Programs", courtHead)) + 1
                    endCol = alphaToNum(headerFind("Conditions", courtHead)) - 1
                Case "Conditions"
                    startCol = alphaToNum(headerFind("Conditions", courtHead)) + 1
                    endCol = alphaToNum(headerFind("ADULT", courtHead)) - 1
                Case Else
                    MsgBox "The section list has changed and the code was not updated to reflect that." _
                        + vbNewLine _
                        + "You were looking for:" _
                        + "Section: " + DB_Section.value _
                        + "Sub-Section: " + DB_Subsection.value _
                        + "Contact system admin."
            End Select
        Case "Adult"
            courtHead = headerFind("ADULT")
            
            Select Case DB_Subsection.value
                Case "All"
                    startCol = alphaToNum(courtHead) + 1
                    endCol = alphaToNum(headerFind("AGGREGATES", courtHead)) - 1
                Case "General"
                    startCol = alphaToNum(courtHead) + 1
                    endCol = alphaToNum(headerFind("OUTCOMES", courtHead)) - 1
                Case "Outcomes"
                    startCol = alphaToNum(headerFind("OUTCOMES", courtHead)) + 1
                    endCol = alphaToNum(headerFind("COURT PROCEEDINGS", courtHead)) - 1
                Case "Certification"
                    startCol = alphaToNum(headerFind("Certification", courtHead)) + 1
                    endCol = alphaToNum(headerFind("Admissions", courtHead)) - 1
                Case "Admissions"
                    startCol = alphaToNum(headerFind("Admissions", courtHead)) + 1
                    endCol = alphaToNum(headerFind("Adjudications", courtHead)) - 1
                Case "Adjudications"
                    startCol = alphaToNum(headerFind("Adjudications", courtHead)) + 1
                    endCol = alphaToNum(headerFind("Continuances", courtHead)) - 1
                Case "Continuances"
                    startCol = alphaToNum(headerFind("Continuances", courtHead)) + 1
                    endCol = alphaToNum(headerFind("Supervision Programs", courtHead)) - 1
                Case "Supervision Programs"
                    startCol = alphaToNum(headerFind("Supervision Programs", courtHead)) + 1
                    endCol = alphaToNum(headerFind("Conditions", courtHead)) - 1
                Case "Conditions"
                    startCol = alphaToNum(headerFind("Conditions", courtHead)) + 1
                    endCol = alphaToNum(headerFind("AGGREGATES", courtHead)) - 1
                Case Else
                    MsgBox "The section list has changed and the code was not updated to reflect that." _
                        + vbNewLine _
                        + "You were looking for:" _
                        + "Section: " + DB_Section.value _
                        + "Sub-Section: " + DB_Subsection.value _
                        + "Contact system admin."
            End Select
        Case "Aggregates"
            courtHead = headerFind("AGGREGATES")
            
            Select Case DB_Subsection.value
                Case "All"
                    startCol = alphaToNum(courtHead) + 1
                    endCol = alphaToNum(headerFind("ARREST GRAPH", courtHead)) - 1
                Case "Legal Status"
                    startCol = alphaToNum(headerFind("LEGAL STATUS", courtHead)) + 1
                    endCol = alphaToNum(headerFind("Petition Outcomes", courtHead)) - 1
                Case "Petition Outcomes"
                    startCol = alphaToNum(headerFind("Petition Outcomes", courtHead)) + 1
                    endCol = alphaToNum(headerFind("COURT PROCEEDINGS", courtHead)) - 1
                Case "Certification"
                    startCol = alphaToNum(headerFind("Certification", courtHead)) + 1
                    endCol = alphaToNum(headerFind("Admissions", courtHead)) - 1
                Case "Admissions"
                    startCol = alphaToNum(headerFind("Admissions", courtHead)) + 1
                    endCol = alphaToNum(headerFind("Adjudications", courtHead)) - 1
                Case "Adjudications"
                    startCol = alphaToNum(headerFind("Adjudications", courtHead)) + 1
                    endCol = alphaToNum(headerFind("Continuances", courtHead)) - 1
                Case "Continuances"
                    startCol = alphaToNum(headerFind("Continuances", courtHead)) + 1
                    endCol = alphaToNum(headerFind("PLACEMENTS", courtHead)) - 1
                Case "Placements"
                    startCol = alphaToNum(headerFind("PLACEMENTS", courtHead)) + 1
                    endCol = alphaToNum(headerFind("Supervision Programs", courtHead)) - 1
                Case "Supervision Programs"
                    startCol = alphaToNum(headerFind("Supervision Programs", courtHead)) + 1
                    endCol = alphaToNum(headerFind("Conditions", courtHead)) - 1
                Case "Conditions"
                    startCol = alphaToNum(headerFind("Conditions", courtHead)) + 1
                    endCol = alphaToNum(headerFind("Restitution & Costs", courtHead)) - 1
                Case "Restitution"
                    startCol = alphaToNum(headerFind("Restitution & Costs", courtHead)) + 1
                    endCol = alphaToNum(headerFind("EXPUNGEMENTS", courtHead)) - 1
                Case "Expungment"
                    startCol = alphaToNum(headerFind("EXPUNGEMENTS", courtHead)) + 1
                    endCol = alphaToNum(headerFind("ARREST GRAPH", courtHead)) - 1
                Case Else
                    MsgBox "The section list has changed and the code was not updated to reflect that." _
                        + vbNewLine _
                        + "You were looking for:" _
                        + "Section: " + DB_Section.value _
                        + "Sub-Section: " + DB_Subsection.value _
                        + "Contact system admin."
            End Select
        Case Else
    End Select
    
    For count = startCol To endCol
        With LookupBox
                .ColumnCount = 4
                .ColumnWidths = "0;20;130;130;"
                .AddItem count
                    .List(LookupBox.ListCount - 1, 1) = numToAlpha(count)
                    .List(LookupBox.ListCount - 1, 2) = Cells(2, count)
                    If (Not IsEmpty(Cells(1, count))) Then
                        Dim word As String
                        Dim word2 As Long
                        Dim word3 As String
                        
                        word = Cells(1, count).value + "_Num"
                        word2 = Cells(updateRow, count).value
                        word3 = Lookup(word)(word2)
                        .List(LookupBox.ListCount - 1, 3) = word3
                    Else
                        .List(LookupBox.ListCount - 1, 3) = Cells(updateRow, count)
                    End If
                    
        End With
    Next count
End Sub




Private Sub SearchButton_Click()
    On Error Resume Next
    
    'define variable Long(a big integer) named emptyRow
    Dim lastRow As Long
    Dim Query As String
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
                .ColumnCount = 8
                .ColumnWidths = "30;70;80;80;70;70;70;70;"
                .AddItem lookRow
                    .List(SearchResultsBox.ListCount - 1, 1) = Range(headerFind("First Name") & lookRow)
                    .List(SearchResultsBox.ListCount - 1, 2) = Range(headerFind("Last Name") & lookRow)
                    .List(SearchResultsBox.ListCount - 1, 3) = Range(headerFind("DOB") & lookRow)
                    .List(SearchResultsBox.ListCount - 1, 4) = Range(headerFind("Arrest Date") & lookRow)
                    .List(SearchResultsBox.ListCount - 1, 5) = Range(headerFind("Petition #1") & lookRow)
                    .List(SearchResultsBox.ListCount - 1, 6) = Lookup("Courtroom_Num")(Range(headerFind("Active Courtroom") & lookRow).value)
                    .List(SearchResultsBox.ListCount - 1, 7) = Lookup("Legal_Status_Num")(Range(headerFind("Legal Status") & lookRow).value)
            End With
        End If
    Next lookRow
End Sub

Private Sub Submit_Click()
    Dim count As Long
    
    For count = 0 To (UpdateBox.ListCount - 1)
        If Not IsEmpty(Range(UpdateBox.List(count, 1) & "1")) Then
            Range(UpdateBox.List(count, 1) & updateRow).value _
                = Lookup(Range(UpdateBox.List(count, 1) & "1") + "_Name")(UpdateBox.List(count, 3))
        Else
            Range(UpdateBox.List(count, 1) & updateRow).value _
                = UpdateBox.List(count, 3)
        End If
    Next count
    
    Unload Me
End Sub

