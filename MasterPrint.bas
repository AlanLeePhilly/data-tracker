Attribute VB_Name = "MasterPrint"
Sub PrintRecent()

    Dim Origbook As Workbook
    Set Origbook = ThisWorkbook
    
    Dim YouthSheet As Worksheet
    Dim DataSheet As Worksheet
    Dim EntrySheet As Worksheet
    Set YouthSheet = Origbook.Worksheets("Youth Search")
    Set DataSheet = Origbook.Worksheets("Run Sheet")
    Set EntrySheet = Origbook.Worksheets("Entry")

    Application.SheetsInNewWorkbook = 3
    Dim wb As Workbook

    Dim x As Integer
    Dim r As Range
    
    Application.ScreenUpdating = False
    
    'Find start of Petition #1 column on run sheet
    Set r = DataSheet.Cells.Find("Petition #1")
    DataSheet.Activate
    r.Select
    r.Activate
    
    'set number of rows of data
    NumRows = DataSheet.Range(r, ActiveCell.End(xlDown)).Rows.count
    
    For x = 1 To NumRows
        'Get Petition #1 for current kid
        DataSheet.Activate
        r.Select
        r.Activate
        ActiveCell.Offset(x, 0).Select
        
        'Load data into Youth Search
        Dim rfind As Range
        Set rfind = EntrySheet.Cells.Find("Petition #1")
        YouthSheet.Range("J5").value = EntrySheet.Cells.Find(ActiveCell.value, After:=rfind, SearchOrder:=xlColumns).row
        YouthSheet.Activate
        Run ("YouthSearchPrint0")
        
        'Create temp workbook
        Set wb = Workbooks.Add
        
        Call format(wb, x)
        
        wb.PrintOut From:=1, To:=2
        
        'Print Supervision, Condition, and Court Listing History
        If YouthSheet.Cells(270, 4).value = "None" And YouthSheet.Cells(270, 19).value = "None" Then
            wb.Worksheets(2).PageSetup.PrintArea = "A1:Z100"
            If IsEmpty(YouthSheet.Cells(543, 11).value) Then
                wb.PrintOut From:=3, To:=3
            ElseIf IsEmpty(YouthSheet.Cells(615, 11).value) Then
                If IsEmpty(YouthSheet.Cells(551, 11).value) Or IsEmpty(YouthSheet.Cells(559, 11).value) Then
                    wb.PrintOut From:=3, To:=3
                    wb.PrintOut From:=4, To:=4
                Else
                    wb.PrintOut From:=3, To:=4
                End If
            ElseIf IsEmpty(YouthSheet.Cells(687, 11).value) Then
                If IsEmpty(YouthSheet.Cells(623, 11).value) Or IsEmpty(YouthSheet.Cells(631, 11).value) Then
                    wb.PrintOut From:=3, To:=5
                Else
                    wb.Worksheets(3).PageSetup.PrintArea = "A76:M322"
                    wb.PrintOut From:=3, To:=4
                End If
            ElseIf IsEmpty(YouthSheet.Cells(759, 11).value) Then
                If IsEmpty(YouthSheet.Cells(695, 11).value) Or IsEmpty(YouthSheet.Cells(703, 11).value) Then
                    wb.Worksheets(3).PageSetup.PrintArea = "A76:M322"
                    wb.PrintOut From:=3, To:=5
                Else
                    wb.Worksheets(3).PageSetup.PrintArea = "A148:M322"
                    wb.PrintOut From:=3, To:=4
                End If
            Else
                If IsEmpty(YouthSheet.Cells(767, 11).value) Or IsEmpty(YouthSheet.Cells(775, 11).value) Then
                    wb.Worksheets(3).PageSetup.PrintArea = "A148:M322"
                    wb.PrintOut From:=3, To:=5
                Else
                    wb.Worksheets(3).PageSetup.PrintArea = "A220:M322"
                    wb.PrintOut From:=3, To:=4
                End If
            End If
        
        ElseIf YouthSheet.Cells(368, 4).value = "None" And YouthSheet.Cells(368, 19).value = "None" Then
            wb.Worksheets(2).PageSetup.PrintArea = "A1:Z198"
            If IsEmpty(YouthSheet.Cells(543, 11).value) Then
                wb.PrintOut From:=3, To:=4
            ElseIf IsEmpty(YouthSheet.Cells(615, 11).value) Then
                wb.PrintOut From:=3, To:=5
            ElseIf IsEmpty(YouthSheet.Cells(687, 11).value) Then
                If IsEmpty(YouthSheet.Cells(623, 11).value) Or IsEmpty(YouthSheet.Cells(631, 11).value) Then
                    wb.PrintOut From:=3, To:=5
                    wb.PrintOut From:=6, To:=6
                Else
                    wb.Worksheets(3).PageSetup.PrintArea = "A76:M322"
                    wb.PrintOut From:=3, To:=5
                End If
            ElseIf IsEmpty(YouthSheet.Cells(759, 11).value) Then
                If IsEmpty(YouthSheet.Cells(695, 11).value) Or IsEmpty(YouthSheet.Cells(703, 11).value) Then
                    wb.Worksheets(3).PageSetup.PrintArea = "A76:M322"
                    wb.PrintOut From:=3, To:=5
                    wb.PrintOut From:=6, To:=6
                Else
                    wb.Worksheets(3).PageSetup.PrintArea = "A148:M322"
                    wb.PrintOut From:=3, To:=5
                End If
            Else
                If IsEmpty(YouthSheet.Cells(767, 11).value) Or IsEmpty(YouthSheet.Cells(775, 11).value) Then
                    wb.Worksheets(3).PageSetup.PrintArea = "A148:M322"
                    wb.PrintOut From:=3, To:=5
                    wb.PrintOut From:=6, To:=6
                Else
                    wb.Worksheets(3).PageSetup.PrintArea = "A220:M322"
                    wb.PrintOut From:=3, To:=5
                End If
            End If
            
        Else
            If YouthSheet.Cells(382, 4).value = "None" And YouthSheet.Cells(382, 19).value = "None" Then
                wb.PrintOut From:=3, To:=3
                If IsEmpty(YouthSheet.Cells(543, 11).value) Then
                    wb.PrintOut From:=4, To:=5
                ElseIf IsEmpty(YouthSheet.Cells(615, 11).value) Then
                    wb.PrintOut From:=4, To:=6
                ElseIf IsEmpty(YouthSheet.Cells(687, 11).value) Then
                    If IsEmpty(YouthSheet.Cells(623, 11).value) Or IsEmpty(YouthSheet.Cells(631, 11).value) Then
                        wb.PrintOut From:=4, To:=6
                        wb.PrintOut From:=7, To:=7
                    Else
                        wb.Worksheets(3).PageSetup.PrintArea = "A76:M322"
                        wb.PrintOut From:=4, To:=6
                    End If
                ElseIf IsEmpty(YouthSheet.Cells(759, 11).value) Then
                    If IsEmpty(YouthSheet.Cells(695, 11).value) Or IsEmpty(YouthSheet.Cells(703, 11).value) Then
                        wb.Worksheets(3).PageSetup.PrintArea = "A76:M322"
                        wb.PrintOut From:=4, To:=6
                        wb.PrintOut From:=7, To:=7
                    Else
                        wb.Worksheets(3).PageSetup.PrintArea = "A148:M322"
                        wb.PrintOut From:=4, To:=6
                    End If
                Else
                    If IsEmpty(YouthSheet.Cells(767, 11).value) Or IsEmpty(YouthSheet.Cells(775, 11).value) Then
                        wb.Worksheets(3).PageSetup.PrintArea = "A148:M322"
                        wb.PrintOut From:=4, To:=6
                        wb.PrintOut From:=7, To:=7
                    Else
                        wb.Worksheets(3).PageSetup.PrintArea = "A220:M322"
                        wb.PrintOut From:=4, To:=6
                    End If
                End If
            Else
                wb.PrintOut From:=3, To:=4
                If IsEmpty(YouthSheet.Cells(543, 11).value) Then
                    wb.PrintOut From:=5, To:=5
                ElseIf IsEmpty(YouthSheet.Cells(615, 11).value) Then
                    wb.PrintOut From:=5, To:=6
                ElseIf IsEmpty(YouthSheet.Cells(687, 11).value) Then
                    If IsEmpty(YouthSheet.Cells(623, 11).value) Or IsEmpty(YouthSheet.Cells(631, 11).value) Then
                        wb.PrintOut From:=5, To:=7
                    Else
                        wb.Worksheets(3).PageSetup.PrintArea = "A76:M322"
                        wb.PrintOut From:=5, To:=6
                    End If
                ElseIf IsEmpty(YouthSheet.Cells(759, 11).value) Then
                    If IsEmpty(YouthSheet.Cells(695, 11).value) Or IsEmpty(YouthSheet.Cells(703, 11).value) Then
                        wb.Worksheets(3).PageSetup.PrintArea = "A76:M322"
                        wb.PrintOut From:=5, To:=7
                    Else
                        wb.Worksheets(3).PageSetup.PrintArea = "A148:M322"
                        wb.PrintOut From:=5, To:=6
                    End If
                Else
                    If IsEmpty(YouthSheet.Cells(767, 11).value) Or IsEmpty(YouthSheet.Cells(775, 11).value) Then
                        wb.Worksheets(3).PageSetup.PrintArea = "A148:M322"
                        wb.PrintOut From:=5, To:=7
                    Else
                        wb.Worksheets(3).PageSetup.PrintArea = "A220:M322"
                        wb.PrintOut From:=5, To:=6
                    End If
                End If
            End If
        End If
        
        'close temp workbook memory
        wb.Close SaveChanges:=False
    
    Next x
    
    'clear temp workbook from memory
    Set wb = Nothing
    
    'reset SheetsInNewWorkbook
    Application.SheetsInNewWorkbook = 1
        
    Application.ScreenUpdating = True

End Sub
Sub PrintFull()

    Dim Origbook As Workbook
    Set Origbook = ThisWorkbook
    
    Dim YouthSheet As Worksheet
    Dim DataSheet As Worksheet
    Dim EntrySheet As Worksheet
    Set YouthSheet = Origbook.Worksheets("Youth Search")
    Set DataSheet = Origbook.Worksheets("Run Sheet")
    Set EntrySheet = Origbook.Worksheets("Entry")

    Application.SheetsInNewWorkbook = 3
    Dim wb As Workbook

    Dim x As Integer
    Dim r As Range
    
    Application.ScreenUpdating = False
    
    'Find start of Petition #1 column on run sheet
    Set r = DataSheet.Cells.Find("Petition #1")
    DataSheet.Activate
    r.Select
    r.Activate
    
    'set number of rows of data
    NumRows = DataSheet.Range(r, ActiveCell.End(xlDown)).Rows.count
    
    For x = 1 To NumRows
        'Get Petition #1 for current kid
        DataSheet.Activate
        r.Select
        r.Activate
        ActiveCell.Offset(x, 0).Select
        
        'Load data into Youth Search
        Dim rfind As Range
        Set rfind = EntrySheet.Cells.Find("Petition #1")
        YouthSheet.Range("J5").value = EntrySheet.Cells.Find(ActiveCell.value, After:=rfind, SearchOrder:=xlColumns).row
        YouthSheet.Activate
        Run ("YouthSearchPrint0")
        
        'Create temp workbook
        Set wb = Workbooks.Add
        
        Call format(wb, x)
            
        'Print Full History
        wb.PrintOut From:=1, To:=2
        
        If YouthSheet.Cells(270, 4).value = "None" And YouthSheet.Cells(270, 19).value = "None" Then
            wb.Worksheets(2).PageSetup.PrintArea = "A1:Z100"
            If IsEmpty(YouthSheet.Cells(543, 11).value) Then
                wb.PrintOut From:=3, To:=3
            ElseIf IsEmpty(YouthSheet.Cells(615, 11).value) Then
                If IsEmpty(YouthSheet.Cells(551, 11).value) Or IsEmpty(YouthSheet.Cells(559, 11).value) Then
                    wb.PrintOut From:=3, To:=3
                    wb.PrintOut From:=4, To:=4
                Else
                    wb.PrintOut From:=3, To:=4
                End If
            ElseIf IsEmpty(YouthSheet.Cells(687, 11).value) Then
                wb.PrintOut From:=3, To:=5
            ElseIf IsEmpty(YouthSheet.Cells(759, 11).value) Then
                If IsEmpty(YouthSheet.Cells(695, 11).value) Or IsEmpty(YouthSheet.Cells(703, 11).value) Then
                    wb.PrintOut From:=3, To:=5
                    wb.PrintOut From:=6, To:=6
                Else
                    wb.PrintOut From:=3, To:=6
                End If
            Else
                wb.PrintOut From:=3, To:=7
            End If
            
        ElseIf YouthSheet.Cells(368, 4).value = "None" And YouthSheet.Cells(368, 19).value = "None" Then
            wb.Worksheets(2).PageSetup.PrintArea = "A1:Z198"
            If IsEmpty(YouthSheet.Cells(543, 11).value) Then
                wb.PrintOut From:=3, To:=4
            ElseIf IsEmpty(YouthSheet.Cells(615, 11).value) Then
                wb.PrintOut From:=3, To:=5
            ElseIf IsEmpty(YouthSheet.Cells(687, 11).value) Then
                If IsEmpty(YouthSheet.Cells(623, 11).value) Or IsEmpty(YouthSheet.Cells(631, 11).value) Then
                    wb.PrintOut From:=3, To:=5
                    wb.PrintOut From:=6, To:=6
                Else
                    wb.PrintOut From:=3, To:=6
                End If
            ElseIf IsEmpty(YouthSheet.Cells(759, 11).value) Then
                wb.PrintOut From:=3, To:=7
            Else
                If IsEmpty(YouthSheet.Cells(767, 11).value) Or IsEmpty(YouthSheet.Cells(775, 11).value) Then
                    wb.PrintOut From:=3, To:=7
                    wb.PrintOut From:=8, To:=8
                Else
                    wb.PrintOut From:=3, To:=8
                End If
            End If

        Else
            If YouthSheet.Cells(382, 4).value = "None" And YouthSheet.Cells(382, 19).value = "None" Then
                wb.PrintOut From:=3, To:=3
                If IsEmpty(YouthSheet.Cells(543, 11).value) Then
                    wb.PrintOut From:=4, To:=5
                ElseIf IsEmpty(YouthSheet.Cells(615, 11).value) Then
                    wb.PrintOut From:=4, To:=6
                ElseIf IsEmpty(YouthSheet.Cells(687, 11).value) Then
                    If IsEmpty(YouthSheet.Cells(623, 11).value) Or IsEmpty(YouthSheet.Cells(631, 11).value) Then
                        wb.PrintOut From:=4, To:=6
                        wb.PrintOut From:=7, To:=7
                    Else
                        wb.PrintOut From:=4, To:=7
                    End If
                ElseIf IsEmpty(YouthSheet.Cells(759, 11).value) Then
                    wb.PrintOut From:=4, To:=8
                Else
                    If IsEmpty(YouthSheet.Cells(767, 11).value) Or IsEmpty(YouthSheet.Cells(775, 11).value) Then
                        wb.PrintOut From:=4, To:=8
                        wb.PrintOut From:=9, To:=9
                    Else
                        wb.PrintOut From:=4, To:=9
                    End If
                End If
            Else
                wb.PrintOut From:=3, To:=4
                If IsEmpty(YouthSheet.Cells(543, 11).value) Then
                    wb.PrintOut From:=5, To:=5
                ElseIf IsEmpty(YouthSheet.Cells(615, 11).value) Then
                    If IsEmpty(YouthSheet.Cells(551, 11).value) Or IsEmpty(YouthSheet.Cells(559, 11).value) Then
                        wb.PrintOut From:=5, To:=5
                        wb.PrintOut From:=6, To:=6
                    Else
                        wb.PrintOut From:=5, To:=6
                    End If
                ElseIf IsEmpty(YouthSheet.Cells(687, 11).value) Then
                    wb.PrintOut From:=5, To:=7
                ElseIf IsEmpty(YouthSheet.Cells(759, 11).value) Then
                    If IsEmpty(YouthSheet.Cells(695, 11).value) Or IsEmpty(YouthSheet.Cells(703, 11).value) Then
                        wb.PrintOut From:=5, To:=7
                        wb.PrintOut From:=8, To:=8
                    Else
                        wb.PrintOut From:=5, To:=8
                    End If
                Else
                    wb.PrintOut From:=5, To:=9
                End If
            End If
        End If
        
        'close temp workbook memory
        wb.Close SaveChanges:=False
    
    Next x
    
    'clear temp workbook from memory
    Set wb = Nothing
    
    'reset SheetsInNewWorkbook
    Application.SheetsInNewWorkbook = 1
        
    Application.ScreenUpdating = True
    
End Sub
Sub ExportPDF()

    Dim Origbook As Workbook
    Set Origbook = ThisWorkbook
    
    Dim YouthSheet As Worksheet
    Dim DataSheet As Worksheet
    Dim EntrySheet As Worksheet
    Set YouthSheet = Origbook.Worksheets("Youth Search")
    Set DataSheet = Origbook.Worksheets("Run Sheet")
    Set EntrySheet = Origbook.Worksheets("Entry")

    Application.SheetsInNewWorkbook = 3
    Dim wb As Workbook

    Dim x As Integer
    Dim r As Range
    
    Application.ScreenUpdating = False
    
    'Find start of Petition #1 column on run sheet
    Set r = DataSheet.Cells.Find("Petition #1")
    DataSheet.Activate
    r.Select
    r.Activate
    
    Dim rName As Range
    
    'Find start of Last Name column on run sheet
    Set rName = DataSheet.Cells.Find("Last Name")
    
    'set number of rows of data
    NumRows = DataSheet.Range(r, ActiveCell.End(xlDown)).Rows.count
    
    For x = 1 To NumRows
        'Get Petition #1 for current kid
        DataSheet.Activate
        r.Select
        r.Activate
        ActiveCell.Offset(x, 0).Select
        
        'Find row in Entry
        Dim rfind As Range
        Set rfind = EntrySheet.Cells.Find("Petition #1")
        YouthSheet.Range("J5").value = EntrySheet.Cells.Find(ActiveCell.value, After:=rfind, SearchOrder:=xlColumns).row
        
        'Hold Petition # in string for pdf naming
        Dim petnum As String
        petnum = ActiveCell.value
        
        'Load data into Youth Search
        YouthSheet.Activate
        Run ("YouthSearchPrint0")
        
        'Create temp workbook
        Set wb = Workbooks.Add
        
        Call format(wb, x)
        
        'Get last name of current kid
        DataSheet.Activate
        rName.Select
        rName.Activate
        ActiveCell.Offset(x, 0).Select
        
        'Export Full History
        If YouthSheet.Cells(270, 4).value = "None" And YouthSheet.Cells(270, 19).value = "None" Then
            wb.Worksheets(2).PageSetup.PrintArea = "A1:Z100"
            If IsEmpty(YouthSheet.Cells(543, 11).value) Then
                wb.Worksheets(3).PageSetup.PrintArea = "A1:M74"
            ElseIf IsEmpty(YouthSheet.Cells(615, 11).value) Then
                wb.Worksheets(3).PageSetup.PrintArea = "A1:M146"
            ElseIf IsEmpty(YouthSheet.Cells(687, 11).value) Then
                wb.Worksheets(3).PageSetup.PrintArea = "A1:M218"
            ElseIf IsEmpty(YouthSheet.Cells(759, 11).value) Then
                wb.Worksheets(3).PageSetup.PrintArea = "A1:M290"
            End If
            
        ElseIf YouthSheet.Cells(368, 4).value = "None" And YouthSheet.Cells(368, 19).value = "None" Then
            wb.Worksheets(2).PageSetup.PrintArea = "A1:Z198"
            If IsEmpty(YouthSheet.Cells(543, 11).value) Then
                wb.Worksheets(3).PageSetup.PrintArea = "A1:M74"
            ElseIf IsEmpty(YouthSheet.Cells(615, 11).value) Then
                wb.Worksheets(3).PageSetup.PrintArea = "A1:M146"
            ElseIf IsEmpty(YouthSheet.Cells(687, 11).value) Then
                wb.Worksheets(3).PageSetup.PrintArea = "A1:M218"
            ElseIf IsEmpty(YouthSheet.Cells(759, 11).value) Then
                wb.Worksheets(3).PageSetup.PrintArea = "A1:M290"
            End If
        
        Else
            If IsEmpty(YouthSheet.Cells(543, 11).value) Then
                wb.Worksheets(3).PageSetup.PrintArea = "A1:M74"
            ElseIf IsEmpty(YouthSheet.Cells(615, 11).value) Then
                wb.Worksheets(3).PageSetup.PrintArea = "A1:M146"
            ElseIf IsEmpty(YouthSheet.Cells(687, 11).value) Then
                wb.Worksheets(3).PageSetup.PrintArea = "A1:M218"
            ElseIf IsEmpty(YouthSheet.Cells(759, 11).value) Then
                wb.Worksheets(3).PageSetup.PrintArea = "A1:M290"
            End If
        End If
        
        'Export current kid to pdf
        wb.ExportAsFixedFormat Type:=xlTypePDF, Filename:="C:\Users\piggej\Desktop\" & ActiveCell.value & Mid(petnum, 9, 8) & ".pdf"
        
        'close temp workbook memory
        wb.Close SaveChanges:=False
    
    Next x
    
    'clear temp workbook from memory
    Set wb = Nothing
    
    'reset SheetsInNewWorkbook
    Application.SheetsInNewWorkbook = 1
        
    Application.ScreenUpdating = True
    
End Sub
Sub format(wb As Workbook, x As Integer)

    Dim Origbook As Workbook
    Set Origbook = ThisWorkbook
    
    Dim YouthSheet As Worksheet
    Dim DataSheet As Worksheet
    Dim EntrySheet As Worksheet
    Set YouthSheet = Origbook.Worksheets("Youth Search")
    Set DataSheet = Origbook.Worksheets("Run Sheet")
    Set EntrySheet = Origbook.Worksheets("Entry")
    
    'Create temp worksheet for Petition Info
    YouthSheet.Range("B9:J31").Copy Destination:=wb.Worksheets(1).Range("A1")
    YouthSheet.Range("B34:J44").Copy Destination:=wb.Worksheets(1).Range("A24")
    YouthSheet.Range("B45:J75").Copy Destination:=wb.Worksheets(1).Range("K1")
    YouthSheet.Range("M9:AC56").Copy Destination:=wb.Worksheets(1).Range("B37")
    
    With wb.Worksheets(1)
        .Rows(24).RowHeight = 15
        .Rows(25).RowHeight = 15
        
        .Columns("J").EntireColumn.ColumnWidth = 4.5
        
        .Range("C4:C18, H4:H18, D20:D22, C27:C33, F27:F31, I27:I31, D40, E42:E48, I40:I79, D67:D77, M4:M30, P4:P30, S4:S30").Font.Size = 18
        .Range("C51").Font.Size = 14
        
        .Range("C70").Cut Destination:=.Range("C69")
        .Range("C73").Cut Destination:=.Range("C71")
        .Range("C75").Cut Destination:=.Range("C72")
        .Range("C77").Cut Destination:=.Range("C74")
        .Range("C79").Cut Destination:=.Range("C75")
        .Range("C81").Cut Destination:=.Range("C77")
        .Range("C83").Cut Destination:=.Range("C78")
        
        .Rows(84).Delete
        .Rows(83).Delete
        .Rows(82).Delete
        
        .Range("G37:R37").UnMerge
        .Range("G37").Cut Destination:=.Range("G38")
        .Rows(37).Delete
        .Range("G37:R37").Merge
        
        .Columns("C").EntireColumn.ColumnWidth = 21.14
        .Columns("H").EntireColumn.ColumnWidth = 14
        .Columns("M").EntireColumn.ColumnWidth = 18.43
        .Columns("S").EntireColumn.ColumnWidth = 12.71
        .Columns("A").EntireColumn.ColumnWidth = 5.86
        .Columns("E").EntireColumn.ColumnWidth = 12.29
        .Columns("I").EntireColumn.ColumnWidth = 11.57
        .Columns("N").EntireColumn.ColumnWidth = 10.43
        
        .Range("J1:J31, J32:S36, A35:I36, A37:A84, S37:S84").Interior.Color = RGB(255, 255, 255)
        .Range("A24:I24").Interior.Color = RGB(252, 228, 214)
        
        .Range("A8:I8").Borders.LineStyle = xlDouble
        .Range("I16:I22").Borders(xlEdgeRight).LineStyle = xlDouble
        .Range("K1:K15, K19:K30").Borders(xlEdgeLeft).LineStyle = xlDouble
        .Range("B37:R37").Borders(xlEdgeTop).LineStyle = xlDouble
        .Range("B80:R80, G37:R37").Borders(xlEdgeBottom).LineStyle = xlDouble
        
        .Range("C39").IndentLevel = 0
        .Range("C39").HorizontalAlignment = xlRight
        
        .Range("D27").Cut Destination:=.Range("E27")
        .Range("E27").HorizontalAlignment = xlRight
        .Range("E27").IndentLevel = 0
        .Range("G27").Cut Destination:=.Range("H27")
        .Range("H27").HorizontalAlignment = xlRight
        .Range("H27").IndentLevel = 0
        
        'Find start of Name columns on run sheet
        Dim LastName As Range
        Dim FirstName As Range
        Dim name As String
        Set LastName = DataSheet.Cells.Find("Last Name")
        Set FirstName = DataSheet.Cells.Find("First Name")
        'Get names of current kids
        DataSheet.Activate
        LastName.Select
        LastName.Activate
        name = ActiveCell.Offset(x, 0).value & ", "
        FirstName.Select
        FirstName.Activate
        name = name & ActiveCell.Offset(x, 0).value
        .Range("C4").value = name
    End With
        
    With wb.Worksheets(1).PageSetup
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesTall = 1
        .FitToPagesWide = 1
        .LeftMargin = Application.InchesToPoints(0)
        .RightMargin = Application.InchesToPoints(0)
        .TopMargin = Application.InchesToPoints(0)
        .BottomMargin = Application.InchesToPoints(0)
        .HeaderMargin = Application.InchesToPoints(0)
        .FooterMargin = Application.InchesToPoints(0)
        .CenterHorizontally = True
        .CenterVertically = True
        .PrintArea = "A1:S80"
        .PrintGridlines = True
    End With
    
    'Create temp worksheet for Courtroom and Legal Status History
    'If you add this back in, you must create the worksheet with 4 tabs instead of 3, update which tab is which, and update the printing logic
    'The easiest way to do this would be to make this the fourth tab and print it at the end
'    YouthSheet.Range("A78:AD192").Copy Destination:=wb.Worksheets(2).Cells
'
'    Dim shp As Shape
'    For Each shp In wb.Worksheets(2).Shapes
'        shp.Delete
'    Next shp
'
'    With wb.Worksheets(2)
'        .Columns("L").EntireColumn.Delete
'
'        .Columns("I").EntireColumn.AutoFit
'        .Columns("AB").EntireColumn.AutoFit
'
'        .Rows(3).Delete
'        .Rows(2).Delete
'
'        Application.DisplayAlerts = False
'        .Range("A1:AC1").UnMerge
'        .Range("A1").Cut Destination:=.Range("B1")
'        .Columns("A").EntireColumn.Delete
'        .Columns("AC").EntireColumn.Delete
'        .Columns("AB").EntireColumn.Delete
'        .Range("A1:AA1").Merge
'
'        .Range("V5:W5, V17:W17, V29:W29, V41:W41, V53:W53, V65:W65, V77:W77").UnMerge
'        .Range("V5").Cut Destination:=.Range("W5")
'        .Range("V17").Cut Destination:=.Range("W17")
'        .Range("V29").Cut Destination:=.Range("W29")
'        .Range("V41").Cut Destination:=.Range("W41")
'        .Range("V53").Cut Destination:=.Range("W53")
'        .Range("V65").Cut Destination:=.Range("W65")
'        .Range("V77").Cut Destination:=.Range("W77")
'        Application.DisplayAlerts = True
'        .Range("W5, W17, W29, W41, W53, W65, W77").HorizontalAlignment = xlRight
'        .Range("W5, W17, W29, W41, W53, W65, W77").IndentLevel = 0
'
'        .Columns("N").EntireColumn.ColumnWidth = 5
'        .Rows(113).Delete
'
'        .Range("N2:N112, O87:AA112").Interior.Color = RGB(255, 255, 255)
'        .Range("A1:AA1").Borders(xlEdgeBottom).LineStyle = xlDouble
'        .Range("A1:AA1").Interior.Color = RGB(252, 228, 214)
'    End With
'
'    With wb.Worksheets(2).PageSetup
'        .Orientation = xlPortrait
'        .Zoom = False
'        .FitToPagesTall = 1
'        .FitToPagesWide = 1
'        .LeftMargin = Application.InchesToPoints(0)
'        .RightMargin = Application.InchesToPoints(0)
'        .TopMargin = Application.InchesToPoints(0)
'        .BottomMargin = Application.InchesToPoints(0)
'        .HeaderMargin = Application.InchesToPoints(0)
'        .FooterMargin = Application.InchesToPoints(0)
'        .CenterHorizontally = True
'        .CenterVertically = True
'        .PrintArea = "A1:AA112"
'        .PrintGridlines = True
'    End With
    
    'Create temp worksheet for Supervision and Condition History
    YouthSheet.Range("A194:AD478").Copy Destination:=wb.Worksheets(2).Cells
       
    For Each shp In wb.Worksheets(2).Shapes
        shp.Delete
    Next shp
    
    With wb.Worksheets(2)
        .Columns("L").EntireColumn.Delete
    
        .Columns("D").EntireColumn.ColumnWidth = 23.29
    
        .Rows(3).Delete
        .Rows(2).Delete
        .Rows(283).Delete
        
        .Columns("A").EntireColumn.Delete
        .Columns("AB").EntireColumn.Delete
        
        .Range("A1").value = "SUPERVISION & CONDITIONS HISTORY"
        .Range("A1:Z1").Borders(xlEdgeBottom).LineStyle = xlDouble
        .Range("A1:Z1").Interior.Color = RGB(252, 228, 214)
        
        .Columns("N").EntireColumn.ColumnWidth = 5
        .Columns("V").EntireColumn.ColumnWidth = 4.71
        .Columns("W").EntireColumn.ColumnWidth = 4.86
        .Columns("O").EntireColumn.ColumnWidth = 6
        .Columns("B").EntireColumn.ColumnWidth = 7
        .Columns("Z").EntireColumn.ColumnWidth = 8
        .Columns("A").EntireColumn.ColumnWidth = 6
        .Columns("G").EntireColumn.ColumnWidth = 5.14
        .Columns("H").EntireColumn.ColumnWidth = 3.86
        
        .Columns("AA").EntireColumn.Delete
        
        .Columns("N").EntireColumn.Interior.Color = RGB(255, 255, 255)
        
        .Rows(101).PageBreak = xlPageBreakManual
        .Rows(199).PageBreak = xlPageBreakManual
        
        'Create note space for additional supervision
        i = 200
        j = 5
        Do Until (YouthSheet.Cells(i, 4).value = "None" Or i > 452)
            i = i + 14
            j = j + 14
        Loop

        If i = 284 Or i = 382 Then
            .Cells(j, 2).Cut Destination:=.Cells(j + 14, 2)
            .Range(.Cells(j - 1, 1), .Cells(j + 10, 12)).value = ""
            .Range(.Cells(j - 1, 1), .Cells(j + 10, 12)).Interior.Color = RGB(255, 255, 255)
            .Range(.Cells(j - 1, 1), .Cells(j + 10, 12)).Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone
            
            .Range(.Cells(j - 2, 1), .Cells(j + 11, 13)).Merge
            .Cells(j - 2, 1).value = "See next page to record notes."
            .Cells(j - 2, 1).HorizontalAlignment = xlCenter
            .Cells(j - 2, 1).VerticalAlignment = xlCenter
            .Cells(j - 2, 1).Font.Color = RGB(0, 0, 0)
            .Cells(j - 2, 1).Font.Bold = True
            .Cells(j - 2, 1).Font.Size = 14
            
            j = j + 14
        End If
        
        If Not (i > 452) Then
            .Cells(j, 3).value = ""
            .Cells(j, 3).Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Range(.Cells(j - 1, 8), .Cells(j + 4, 12)).value = ""
            .Cells(j + 4, 5).value = ""
            .Cells(j + 2, 5).value = "Start Date:"
            .Cells(j + 2, 8).value = "End Date:"
            Dim rect As Shape
            Set rect = .Shapes.AddShape(msoShapeRectangle, .Cells(j + 2, 6).Left, .Cells(j + 2, 6).Top, 15, 15)
            rect.Fill.ForeColor.RGB = RGB(255, 255, 255)
            rect.Line.ForeColor.RGB = RGB(0, 0, 0)
            Dim rect2 As Shape
            Set rect2 = .Shapes.AddShape(msoShapeRectangle, .Cells(j + 2, 10).Left, .Cells(j + 2, 10).Top, 15, 15)
            rect2.Fill.ForeColor.RGB = RGB(255, 255, 255)
            rect2.Line.ForeColor.RGB = RGB(0, 0, 0)
            .Cells(j, 5).value = "Provider:"
            .Range(.Cells(j, 6), .Cells(j, 9)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Range(.Cells(j + 6, 2), .Cells(j + 6, 4)).value = ""
            
            .Range(.Cells(j + 5, 4), .Cells(j + 10, 9)).Interior.Color = RGB(255, 255, 255)
            .Range(.Cells(j + 6, 4), .Cells(j + 10, 9)).Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone
            
            .Range(.Cells(j + 12, 1), .Cells(j + 12, 13)).Borders(xlEdgeTop).LineStyle = xlNone
            .Range(.Cells(j + 12, 1), .Cells(j + 25, 13)).value = ""
            .Range(.Cells(j + 20, 4), .Cells(j + 24, 9)).Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone
            .Range(.Cells(j + 20, 4), .Cells(j + 24, 9)).Interior.Color = RGB(255, 255, 255)
            
            .Range(.Cells(j + 13, 12), .Cells(j + 14, 12)).UnMerge
            .Range(.Cells(j + 4, 2), .Cells(j + 24, 12)).Interior.Color = RGB(255, 255, 235)
            .Range(.Cells(j + 4, 2), .Cells(j + 24, 12)).BorderAround LineStyle:=xlContinuous
            .Cells(j + 5, 2).value = "Reason for:"
            .Cells(j + 5, 2).Font.Size = 14
            .Cells(j + 5, 2).Font.Bold = True
            .Cells(j + 16, 2).value = "Notes:"
            .Cells(j + 16, 2).HorizontalAlignment = xlLeft
            
            .Range(.Cells(j + 27, 1), .Cells(282, 12)).value = ""
            .Range(.Cells(j + 27, 1), .Cells(282, 12)).Interior.Color = RGB(255, 255, 255)
            .Range(.Cells(j + 27, 1), .Cells(282, 13)).Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone
            .Range(.Cells(j + 27, 13), .Cells(282, 13)).Borders(xlEdgeRight).LineStyle = Excel.XlLineStyle.xlDouble
            
        End If
        
        'Create note space for additional condition
        i = 200
        j = 5
        Do Until (YouthSheet.Cells(i, 19).value = "None" Or i > 452)
            i = i + 14
            j = j + 14
        Loop
        
        If i = 284 Or i = 382 Then
            .Cells(j, 16).Cut Destination:=.Cells(j + 14, 16)
            .Range(.Cells(j - 1, 15), .Cells(j + 10, 25)).value = ""
            .Range(.Cells(j - 1, 15), .Cells(j + 10, 25)).Interior.Color = RGB(255, 255, 255)
            .Range(.Cells(j - 1, 16), .Cells(j + 10, 25)).Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone
            
            .Range(.Cells(j - 1, 15), .Cells(j + 11, 26)).Merge
            .Cells(j - 1, 15).value = "See next page to record notes."
            .Cells(j - 1, 15).HorizontalAlignment = xlCenter
            .Cells(j - 1, 15).VerticalAlignment = xlCenter
            .Cells(j - 1, 15).Font.Color = RGB(0, 0, 0)
            .Cells(j - 1, 15).Font.Bold = True
            .Cells(j - 1, 15).Font.Size = 14
            
            j = j + 14
        End If

        If Not (i > 452) Then
            .Cells(j, 17).value = ""
            .Range(.Cells(j, 17), .Cells(j, 18)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Range(.Cells(j, 20), .Cells(j + 4, 25)).value = ""
            .Cells(j + 4, 20).value = ""
            .Cells(j + 2, 20).value = "Start Date:"
            .Cells(j + 2, 23).value = "End Date:"
            Dim rect3 As Shape
            Set rect3 = .Shapes.AddShape(msoShapeRectangle, .Cells(j + 2, 21).Left, .Cells(j + 2, 21).Top, 15, 15)
            rect3.Fill.ForeColor.RGB = RGB(255, 255, 255)
            rect3.Line.ForeColor.RGB = RGB(0, 0, 0)
            Dim rect4 As Shape
            Set rect4 = .Shapes.AddShape(msoShapeRectangle, .Cells(j + 2, 25).Left, .Cells(j + 2, 25).Top, 15, 15)
            rect4.Fill.ForeColor.RGB = RGB(255, 255, 255)
            rect4.Line.ForeColor.RGB = RGB(0, 0, 0)
            .Cells(j, 20).value = "Provider:"
            .Range(.Cells(j, 21), .Cells(j, 24)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Range(.Cells(j + 6, 17), .Cells(j + 6, 19)).value = ""
            
            .Range(.Cells(j + 5, 19), .Cells(j + 10, 24)).Interior.Color = RGB(255, 255, 255)
            .Range(.Cells(j + 6, 19), .Cells(j + 10, 24)).Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone
            
            .Range(.Cells(j + 12, 15), .Cells(j + 12, 26)).Borders(xlEdgeTop).LineStyle = xlNone
            .Range(.Cells(j + 12, 15), .Cells(j + 25, 26)).value = ""
            .Range(.Cells(j + 20, 19), .Cells(j + 24, 24)).Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone
            .Range(.Cells(j + 20, 19), .Cells(j + 24, 24)).Interior.Color = RGB(255, 255, 255)
            
            .Range(.Cells(j + 13, 25), .Cells(j + 14, 25)).UnMerge
            .Range(.Cells(j + 4, 16), .Cells(j + 24, 25)).Interior.Color = RGB(255, 255, 235)
            .Range(.Cells(j + 4, 16), .Cells(j + 24, 25)).BorderAround LineStyle:=xlContinuous
            .Cells(j + 5, 16).value = "Reason for:"
            .Cells(j + 5, 16).Font.Size = 14
            .Cells(j + 5, 16).Font.Bold = True
            .Cells(j + 16, 16).value = "Notes:"
            .Cells(j + 16, 16).HorizontalAlignment = xlLeft
            
            .Range(.Cells(j + 27, 15), .Cells(282, 25)).value = ""
            .Range(.Cells(j + 27, 15), .Cells(282, 25)).Interior.Color = RGB(255, 255, 255)
            .Range(.Cells(j + 27, 15), .Cells(282, 26)).Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone
        End If
        
    End With
    
    With wb.Worksheets(2).PageSetup
        .Orientation = xlPortrait
        .Zoom = 44
        .LeftMargin = Application.InchesToPoints(0)
        .RightMargin = Application.InchesToPoints(0)
        .TopMargin = Application.InchesToPoints(0)
        .BottomMargin = Application.InchesToPoints(0)
        .HeaderMargin = Application.InchesToPoints(0)
        .FooterMargin = Application.InchesToPoints(0)
        .CenterHorizontally = True
        .PrintArea = "A1:Z282"
        .PrintGridlines = True
    End With
    
    'Create temp worksheet for Court Listing History
    YouthSheet.Range("H482:W813").Copy Destination:=wb.Worksheets(3).Cells
    
    For Each shp In wb.Worksheets(3).Shapes
        shp.Delete
    Next shp
    
    With wb.Worksheets(3)
        .Columns("L").EntireColumn.Delete
        
        .Rows(1).Delete
        .Range("A323:A332").EntireRow.Delete

        .Columns("A").EntireColumn.Delete
        .Columns("O").EntireColumn.Delete
        .Columns("N").EntireColumn.Delete
        
        .Range("A1").value = "COURT LISTINGS"
        .Range("A1").Interior.Color = RGB(252, 228, 214)
        
        .Rows(75).PageBreak = xlPageBreakManual
        .Rows(147).PageBreak = xlPageBreakManual
        .Rows(219).PageBreak = xlPageBreakManual
        .Rows(291).PageBreak = xlPageBreakManual
        
        'Create note space for new listing
        i = 489
        j = 3
        Do Until (YouthSheet.Cells(i, 11).value = "N/A" Or i > 793)
            i = i + 8
            j = j + 8
        Loop
        
        If i = 5532 Or i = 625 Or i = 697 Or i = 769 Then
            .Cells(j + 2, 2).Cut Destination:=.Cells(j + 10, 2)
            .Range(.Cells(j, 1), .Cells(j + 7, 13)).value = ""
            .Range(.Cells(j, 1), .Cells(j + 7, 13)).Interior.Color = RGB(255, 255, 255)
            .Range(.Cells(j, 1), .Cells(j + 7, 13)).Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone
            
            .Range(.Cells(j, 1), .Cells(j + 7, 13)).Merge
            .Cells(j, 1).value = "See next page to record notes."
            .Cells(j, 1).HorizontalAlignment = xlCenter
            .Cells(j, 1).VerticalAlignment = xlCenter
            .Cells(j, 1).Font.Color = RGB(0, 0, 0)
            .Cells(j, 1).Font.Bold = True
            .Cells(j, 1).Font.Size = 14
            
            j = j + 8
        End If

        If Not (i > 793) Then
            .Range(.Cells(j, 1), .Cells(j + 15, 13)).value = ""
            .Range(.Cells(j, 1), .Cells(j + 15, 13)).UnMerge
            .Range(.Cells(j + 2, 8), .Cells(j + 14, 12)).Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone
            
            .Range(.Cells(j + 8, 1), .Cells(j + 8, 13)).Borders(xlEdgeTop).LineStyle = xlNone
            
            .Cells(j + 2, 1).value = "Admission?"
            .Cells(j + 6, 1).value = "Adjudication?"
            .Cells(j + 9, 1).value = "Change in"
            .Cells(j + 10, 1).value = "Legal Status?"
            .Range(.Cells(j + 2, 1), .Cells(j + 10, 1)).Font.Size = 14
            .Range(.Cells(j + 2, 1), .Cells(j + 10, 1)).Font.Bold = True
            Dim rect5 As Shape
            Set rect5 = .Shapes.AddShape(msoShapeRectangle, .Cells(j + 2, 3).Left + 10, .Cells(j + 2, 3).Top, 15, 15)
            rect5.Fill.ForeColor.RGB = RGB(255, 255, 255)
            rect5.Line.ForeColor.RGB = RGB(0, 0, 0)
            Dim rect6 As Shape
            Set rect6 = .Shapes.AddShape(msoShapeRectangle, .Cells(j + 6, 3).Left + 10, .Cells(j + 6, 3).Top, 15, 15)
            rect6.Fill.ForeColor.RGB = RGB(255, 255, 255)
            rect6.Line.ForeColor.RGB = RGB(0, 0, 0)
            Dim rect7 As Shape
            Set rect7 = .Shapes.AddShape(msoShapeRectangle, .Cells(j + 10, 3).Left + 10, .Cells(j + 10, 3).Top, 15, 15)
            rect7.Fill.ForeColor.RGB = RGB(255, 255, 255)
            rect7.Line.ForeColor.RGB = RGB(0, 0, 0)
            
            .Range(.Cells(j + 1, 4), .Cells(j + 14, 12)).Interior.Color = RGB(255, 255, 235)
            .Range(.Cells(j + 1, 4), .Cells(j + 14, 12)).BorderAround LineStyle:=xlContinuous
            
            .Range(.Cells(j + 15, 1), .Cells(322, 13)).value = ""
            .Range(.Cells(j + 15, 1), .Cells(322, 13)).Interior.Color = RGB(255, 255, 255)
            .Range(.Cells(j + 15, 1), .Cells(322, 13)).Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone
            .Rows(j + 13).RowHeight = 1
        End If
    End With
    
    With wb.Worksheets(3).PageSetup
        .Orientation = xlPortrait
        .Zoom = 63
        .LeftMargin = Application.InchesToPoints(0)
        .RightMargin = Application.InchesToPoints(0)
        .TopMargin = Application.InchesToPoints(0)
        .BottomMargin = Application.InchesToPoints(0)
        .HeaderMargin = Application.InchesToPoints(0)
        .FooterMargin = Application.InchesToPoints(0)
        .CenterHorizontally = True
        .PrintArea = "A1:M322"
        .PrintGridlines = True
        .PrintTitleRows = "1:3"
    End With
    
End Sub
