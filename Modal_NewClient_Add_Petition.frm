VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Modal_NewClient_Add_Petition 
   Caption         =   "Add Petition"
   ClientHeight    =   11580
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14715
   OleObjectBlob   =   "Modal_NewClient_Add_Petition.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Modal_NewClient_Add_Petition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
                        '''''''''''''
                        'VALIDATIONS'
                        '''''''''''''
Private Sub DateFiled_Enter()
    DateFiled.value = CalendarForm.GetDate(RangeOfYears:=5)
End Sub
Private Sub DateFiled_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Set ctl = Modal_NewClient_Add_Petition.DateFiled

    Call DateValidation(ctl, Cancel)
End Sub

                        ''''''''''''''''''
                        '''''BUTTONS''''''
                        ''''''''''''''''''
    
Private Sub InsertDoH_Click()
    DateFiled = NewClientForm.InitialHearingDate
End Sub

Private Sub Cancel_Click()
    Unload Modal_NewClient_Add_Petition
End Sub




                         
                        '''''''''''''''''''''''
                        '''''SUBMIT LOGIC''''''
                        '''''''''''''''''''''''

Private Sub Continue_Click()
    'VALIDATIONS
    If Not HasContent(Num) Then
        MsgBox "'Petition #' Required"
        Exit Sub
    End If

    If Not HasContent(DateFiled) Then
        MsgBox "'Date Filed' Required"
        Exit Sub
    End If
    
    If SelectedCharge1.Caption = "" Then
        MsgBox "'Selected Charge' Required"
    End If
    
    If ChargeGrade1.value = "" Then
        MsgBox "Charge Grade Required"
    End If
    
    If ChargeGroup1.value = "" Then
        MsgBox "Charge Group Required"
    End If

    Dim pBox, cBox
    
    If headline.Caption = "Re-Arrest" Then
        Set pBox = Modal_New_Arrest.PetitionBox
        Set cBox = Modal_New_Arrest.ChargeBox
    Else
        If headline.Caption = "New Client" Then
            Set pBox = NewClientForm.PetitionBox
            Set cBox = NewClientForm.ChargeBox
        Else
            MsgBox "Something went wrong. The form that called this modal did not properly identify itself"
            Exit Sub
            
        End If
    End If
    
    With pBox
        .ColumnCount = 7
        .ColumnWidths = "50;50;30;50;65;50;0"
        ' 0 Date Filed
        ' 1 Petition Number
        ' 2 Charge Grade
        ' 3 Charge Group
        ' 4 Charge Code
        ' 5 Charge Name
        ' 6 Was Petition from other county?
        .AddItem DateFiled
                .List(.ListCount - 1, 0) = DateFiled
                .List(.ListCount - 1, 1) = Num
                .List(.ListCount - 1, 2) = ChargeGrade1
                .List(.ListCount - 1, 3) = ChargeGroup1
                .List(.ListCount - 1, 4) = ChargeList1.List(ChargeList1.listIndex, 0)
                .List(.ListCount - 1, 5) = ChargeList1.List(ChargeList1.listIndex, 1)
                .List(.ListCount - 1, 6) = isTransferred
    End With
    
    If Not ChargeList2.value = "" Then
        With cBox
            .ColumnCount = 5
            .ColumnWidths = "50;50;30;50;65;"
            ' 0 Petition Number
            ' 1 Charge Grade
            ' 2 Charge Group (specific)
            ' 3 Charge Code
            ' 4 Charge Name
            .AddItem Num
                    .List(.ListCount - 1, 0) = Num
                    .List(.ListCount - 1, 1) = ChargeGrade2
                    .List(.ListCount - 1, 2) = ChargeGroup2
                    .List(.ListCount - 1, 3) = ChargeList2.List(ChargeList2.listIndex, 0)
                    .List(.ListCount - 1, 4) = ChargeList2.List(ChargeList2.listIndex, 1)
        End With
    End If

    If Not ChargeList3.value = "" Then
        With cBox
            .ColumnCount = 5
            .ColumnWidths = "50;50;30;50;65;"
            ' 0 Petition Number
            ' 1 Charge Grade (specific)
            ' 2 Charge Group (category)
            ' 3 Charge Code
            ' 4 Charge Name
            .AddItem Num
                    .List(.ListCount - 1, 0) = Num
                    .List(.ListCount - 1, 1) = ChargeGrade3
                    .List(.ListCount - 1, 2) = ChargeGroup3
                    .List(.ListCount - 1, 3) = ChargeList3.List(ChargeList3.listIndex, 0)
                    .List(.ListCount - 1, 4) = ChargeList3.List(ChargeList3.listIndex, 1)
        End With
    End If
    
    If Not ChargeList4.value = "" Then
        With cBox
            .ColumnCount = 5
            .ColumnWidths = "50;50;30;50;65;"
            ' 0 Petition Number
            ' 1 Charge Grade
            ' 2 Charge Group (specific)
            ' 3 Charge Code
            ' 4 Charge Name
            .AddItem Num
                    .List(.ListCount - 1, 0) = Num
                    .List(.ListCount - 1, 1) = ChargeGrade4
                    .List(.ListCount - 1, 2) = ChargeGroup4
                    .List(.ListCount - 1, 3) = ChargeList4.List(ChargeList4.listIndex, 0)
                    .List(.ListCount - 1, 4) = ChargeList4.List(ChargeList4.listIndex, 1)
        End With
    End If
    
    If Not ChargeList5.value = "" Then
        With cBox
            .ColumnCount = 5
            .ColumnWidths = "50;50;30;50;65;"
            ' 0 Petition Number
            ' 1 Charge Grade
            ' 2 Charge Group (specific)
            ' 3 Charge Code
            ' 4 Charge Name
            .AddItem Num
                    .List(.ListCount - 1, 0) = Num
                    .List(.ListCount - 1, 1) = ChargeGrade5
                    .List(.ListCount - 1, 2) = ChargeGroup5
                    .List(.ListCount - 1, 3) = ChargeList5.List(ChargeList5.listIndex, 0)
                    .List(.ListCount - 1, 4) = ChargeList5.List(ChargeList5.listIndex, 1)
        End With
    End If
    
    Unload Modal_NewClient_Add_Petition
End Sub



                        ''''''''''''''
                        'SEARCH STUFF'
                        ''''''''''''''
Private Sub ChargeSearchButton1_Click()
    'define variable Long(a big integer) named emptyRow
    Dim lastRow As Long
    Dim Query As String
    Dim lookRow As Long
    Dim lookCell As String
    Dim lookCell2 As String
    
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With
    
        
        'define variable of search query in UPPERCASE named 'query'
        Query = UCase(ChargeSearch1.value)
        
        ChargeList1.Clear
        
        lastRow = Sheets("CrimeCodes").Range("A" & Sheets("CrimeCodes").Rows.count).End(xlUp).row
        
        For lookRow = 2 To lastRow
    
            lookCell = UCase(Sheets("CrimeCodes").Range("A" & lookRow))
            lookCell2 = UCase(Sheets("CrimeCodes").Range("B" & lookRow))
            
            If InStr(1, lookCell, Query) > 0 Or InStr(1, lookCell2, Query) Then
                With ChargeList1
                    .ColumnCount = 2
                    .ColumnWidths = "85;400;"
                    .AddItem Sheets("CrimeCodes").Range("B" & lookRow)
                        .List(ChargeList1.ListCount - 1, 1) = Sheets("CrimeCodes").Range("A" & lookRow)
                End With
            End If
        Next lookRow
    
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
    End With
End Sub

Private Sub ChargeSearchButton2_Click()
    'define variable Long(a big integer) named emptyRow
    Dim lastRow As Long
    Dim Query As String
    Dim lookRow As Long
    Dim lookCell As String
    Dim lookCell2 As String
    
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With
    
        
        'define variable of search query in UPPERCASE named 'query'
        Query = UCase(ChargeSearch2.value)
        
        ChargeList2.Clear
        
        lastRow = Sheets("CrimeCodes").Range("A" & Sheets("CrimeCodes").Rows.count).End(xlUp).row
        
        For lookRow = 2 To lastRow
    
            lookCell = UCase(Sheets("CrimeCodes").Range("A" & lookRow))
            lookCell2 = UCase(Sheets("CrimeCodes").Range("B" & lookRow))
            
            If InStr(1, lookCell, Query) > 0 Or InStr(1, lookCell2, Query) Then
                With ChargeList2
                    .ColumnCount = 2
                    .ColumnWidths = "85;400;"
                    .AddItem Sheets("CrimeCodes").Range("B" & lookRow)
                        .List(ChargeList2.ListCount - 1, 1) = Sheets("CrimeCodes").Range("A" & lookRow)
                End With
            End If
        Next lookRow
    
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
    End With
End Sub

Private Sub ChargeSearchButton3_Click()
    'define variable Long(a big integer) named emptyRow
    Dim lastRow As Long
    Dim Query As String
    Dim lookRow As Long
    Dim lookCell As String
    Dim lookCell2 As String
    
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With
    
        
        'define variable of search query in UPPERCASE named 'query'
        Query = UCase(ChargeSearch3.value)
        
        ChargeList3.Clear
        
        lastRow = Sheets("CrimeCodes").Range("A" & Sheets("CrimeCodes").Rows.count).End(xlUp).row
        
        For lookRow = 2 To lastRow
    
            lookCell = UCase(Sheets("CrimeCodes").Range("A" & lookRow))
            lookCell2 = UCase(Sheets("CrimeCodes").Range("B" & lookRow))
            
            If InStr(1, lookCell, Query) > 0 Or InStr(1, lookCell2, Query) Then
                With ChargeList3
                    .ColumnCount = 2
                    .ColumnWidths = "85;400;"
                    .AddItem Sheets("CrimeCodes").Range("B" & lookRow)
                        .List(ChargeList3.ListCount - 1, 1) = Sheets("CrimeCodes").Range("A" & lookRow)
                End With
            End If
        Next lookRow
    
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
    End With
End Sub

Private Sub ChargeSearchButton4_Click()
    'define variable Long(a big integer) named emptyRow
    Dim lastRow As Long
    Dim Query As String
    Dim lookRow As Long
    Dim lookCell As String
    Dim lookCell2 As String
    
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With
    
        
        'define variable of search query in UPPERCASE named 'query'
        Query = UCase(ChargeSearch4.value)
        
        ChargeList4.Clear
        
        lastRow = Sheets("CrimeCodes").Range("A" & Sheets("CrimeCodes").Rows.count).End(xlUp).row
        
        For lookRow = 2 To lastRow
    
            lookCell = UCase(Sheets("CrimeCodes").Range("A" & lookRow))
            lookCell2 = UCase(Sheets("CrimeCodes").Range("B" & lookRow))
            
            If InStr(1, lookCell, Query) > 0 Or InStr(1, lookCell2, Query) Then
                With ChargeList4
                    .ColumnCount = 2
                    .ColumnWidths = "85;400;"
                    .AddItem Sheets("CrimeCodes").Range("B" & lookRow)
                        .List(ChargeList4.ListCount - 1, 1) = Sheets("CrimeCodes").Range("A" & lookRow)
                End With
            End If
        Next lookRow
    
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
    End With
End Sub

Private Sub ChargeSearchButton5_Click()
    'define variable Long(a big integer) named emptyRow
    Dim lastRow As Long
    Dim Query As String
    Dim lookRow As Long
    Dim lookCell As String
    Dim lookCell2 As String
    
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With
    
        
        'define variable of search query in UPPERCASE named 'query'
        Query = UCase(ChargeSearch5.value)
        
        ChargeList5.Clear
        
        lastRow = Sheets("CrimeCodes").Range("A" & Sheets("CrimeCodes").Rows.count).End(xlUp).row
        
        For lookRow = 2 To lastRow
    
            lookCell = UCase(Sheets("CrimeCodes").Range("A" & lookRow))
            lookCell2 = UCase(Sheets("CrimeCodes").Range("B" & lookRow))
            
            If InStr(1, lookCell, Query) > 0 Or InStr(1, lookCell2, Query) Then
                With ChargeList5
                    .ColumnCount = 2
                    .ColumnWidths = "85;400;"
                    .AddItem Sheets("CrimeCodes").Range("B" & lookRow)
                        .List(ChargeList5.ListCount - 1, 1) = Sheets("CrimeCodes").Range("A" & lookRow)
                End With
            End If
        Next lookRow
    
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
    End With
End Sub


Private Sub ChargeList1_click()
    SelectedCharge1.Caption = ChargeList1.List(ChargeList1.listIndex, 1)
End Sub

Private Sub ChargeList2_click()
    SelectedCharge2.Caption = ChargeList2.List(ChargeList2.listIndex, 1)
End Sub

Private Sub ChargeList3_click()
    SelectedCharge3.Caption = ChargeList3.List(ChargeList3.listIndex, 1)
End Sub

Private Sub ChargeList4_click()
    SelectedCharge4.Caption = ChargeList4.List(ChargeList4.listIndex, 1)
End Sub

Private Sub ChargeList5_click()
    SelectedCharge5.Caption = ChargeList5.List(ChargeList5.listIndex, 1)
End Sub

