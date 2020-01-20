Attribute VB_Name = "Module08_Save_Functions"

Public Function headerFindTwo(ByVal Query As String, Optional ByVal Start As String = "A") As String
    Dim Result As Variant
    If Start = "" Then
        Start = "A"
    End If

    Set Result = Range(Start & "2:XCM2").Find(Query, LookAt:=xlWhole)


    If Result Is Nothing Then
        err.Raise 9999, "headerFindTwo", _
        "Error: Tried to find column header " & Chr(34) & Query & Chr(34) _
        & " starting at column " & Start & " but could not find it."
        Exit Function
    Else
        headerFindTwo = numToAlpha(Result.Column)
        Exit Function
    End If

End Function



'Public Function HOLDheaderFindTwo(ByVal Query As String, Optional ByVal Start As String = "A") As String
'    headerFindTwo = numToAlpha(Range(Start & "1:AAA1").Find(Query).Column)
'End Function




'Public Function HOLDheaderFind(ByVal Query As String, Optional ByVal Start As String = "A") As String
'    headerFind = numToAlpha(Range(Start & "1:AAA1").Find(Query).Column)
'End Function

Function hFindTwo(ParamArray myArgs() As Variant) As String
    ' PASS THIS FUNCTION JUST A LIST OF STRING ARGUMENTS IN ASCENDING SPECIFICITY
    ' EXAMPLE: Call hFind("Start Date", "Pretrial", "4E")

    Dim Result As Variant
    Dim Start As String
    Dim length As Integer
    Dim i As Integer

    length = UBound(myArgs) - LBound(myArgs) + 1
    For i = UBound(myArgs) To LBound(myArgs) Step -1
        If length = 1 Then
            hFindTwo = headerFindTwo(myArgs(i))
            Debug.Print hFind
            Exit Function
        End If
        Start = headerFindTwo(myArgs(i), Start)
    Next i
    hFindTwo = Start
    Debug.Print Start

End Function

Function hFindNumTwo(ParamArray myArgs() As Variant) As Long
    ' PASS THIS FUNCTION JUST A LIST OF STRING ARGUMENTS IN ASCENDING SPECIFICITY
    ' EXAMPLE: Call hFind("Start Date", "Pretrial", "4E")

    Dim Result As Variant
    Dim Start As String
    Dim length As Integer
    Dim i As Integer

    length = UBound(myArgs) - LBound(myArgs) + 1
    For i = UBound(myArgs) To LBound(myArgs) Step -1
        If length = 1 Then
            hFindNumTwo = headerFindNumTwo(myArgs(i))
            Debug.Print hFindNumTwo
            Exit Function
        End If
        Start = headerFindTwo(myArgs(i), Start)
    Next i
    hFindNumTwo = alphaToNum(Start)
    Debug.Print hFindNumTwo

End Function

Public Function headerFindNumTwo(ByVal Query As String, Optional ByVal Start As String = "A") As Long
    Dim Result As Variant
    If Start = "" Then
        Start = "A"
    End If

    Set Result = Range(Start & "1:XCM1").Find(Query, LookAt:=xlWhole)


    If Result Is Nothing Then
        err.Raise 9999, "headerFind", _
        "Error: Tried to find column header " & Chr(34) & Query & Chr(34) _
        & " starting at column " & Start & " but could not find it."
        Exit Function
    Else
        headerFindNumTwo = Result.Column
        Exit Function
    End If

End Function
Sub Save_Countdown()
    saveCounter = saveCounter - 1

    If saveCounter < 1 Then
        ThisWorkbook.Save
        MsgBox "Nice Autosave!"
        saveCounter = 1
    End If
End Sub

Sub SaveAs_Countdown()
    'first subtract from the counter
    saveAsCounter = saveAsCounter - 1

    'declare the string variables we will need for the file name and file path
    Dim newFileName As String
    Dim newFilePath As String

    'used to access a function that will get us the current name of this workbook
    Dim fso As New Scripting.FileSystemObject

    'here you can set the name of the directory where you want the file to go.
    'note that the folder location must exist prior to saving it there
    newFilePath = "H:\SJS Archives\Archives\"

    'Here the new file name is composed of:
    'the filepath specified above
    'the current name of the workbook
    '(space)
    'this second in datetime in a specific format
    'file name extention .xlsm
    newFileName = _
        newFilePath _
        & fso.GetBaseName(ThisWorkbook.name) _
        & " " _
        & VBA.format(Now(), "yyyy-MM-dd hh.mm.ss") _
        & ".xlsm"


    If saveAsCounter < 1 Then
        ThisWorkbook.SaveCopyAs _
            Filename:=newFileName

        MsgBox _
            "Nice AutoSaveAs!" _
            & vbNewLine _
            & "Archived file is " _
            & Chr(34) _
            & newFileName _
            & Chr(34)
        saveAsCounter = 50
    End If
End Sub

Sub Archive()

    'declare the string variables we will need for the file name and file path
    Dim newFileName As String
    Dim newFilePath As String

    'used to access a function that will get us the current name of this workbook
    Dim fso As New Scripting.FileSystemObject

    'here you can set the name of the directory where you want the file to go.
    'note that the folder location must exist prior to saving it there
    newFilePath = "H:\SJS Archives\Archives\"

    'Here the new file name is composed of:
    'the filepath specified above
    'the current name of the workbook
    '(space)
    'this second in datetime in a specific format
    'file name extention .xlsm
    newFileName = _
        newFilePath _
        & fso.GetBaseName(ThisWorkbook.name) _
        & " " _
        & VBA.format(Now(), "yyyy-MM-dd hh.mm.ss") _
        & ".xlsm"


    ThisWorkbook.SaveCopyAs _
            Filename:=newFileName

    MsgBox _
            "Nice Archive!" _
            & vbNewLine _
            & "Archived file is " _
            & Chr(34) _
            & newFileName _
            & Chr(34)


End Sub




Sub ExportDataFile()
    
    Dim rowNum As Long
    Dim lastRow As Long
    
    lastRow = Worksheets("Entry").Range("C" & Rows.count).End(xlUp).row

    For rowNum = 3 To lastRow
        Call AggAggSupervisionsAndConditions(rowNum)
    Next rowNum
    
    
    'ARCHIVE BEFORE EXPORTING
    'declare the string variables we will need for the file name and file path
    Dim newFileName As String
    Dim newFilePath As String

    'used to access a function that will get us the current name of this workbook
    Dim fso As New Scripting.FileSystemObject

    'here you can set the name of the directory where you want the file to go.
    'note that the folder location must exist prior to saving it there
    newFilePath = "H:\SJS Entry\Archives\"

    'Here the new file name is composed of:
    'the filepath specified above
    'the current name of the workbook
    '(space)
    'this second in datetime in a specific format
    'file name extention .xlsm
    newFileName = _
        newFilePath _
        & fso.GetBaseName(ThisWorkbook.name) _
        & " " _
        & VBA.format(Now(), "yyyy-MM-dd hh.mm.ss") _
        & ".xlsm"


    ThisWorkbook.SaveCopyAs _
            Filename:=newFileName

    'EXPORTING COPIES THAT NEED FILTERING FOR ACCESS FIRST

    'DIVERSION TEAM

    'JEANMARIE EXPORT (Master Diversion Export - will trigger all other saveas options)

    'declare the string variables we will need for the file name and file path


    'used to access a function that will get us the current name of this workbook


    'here you can set the name of the directory where you want the file to go.
    'note that the folder location must exist prior to saving it there
    'newFilePath = "H:\SJS Analysis\JeanMarie\Diversion Data Set\"

    'Here the new file name is composed of:
    'the filepath specified above
    'the current name of the workbook
    '(space)
    'this second in datetime in a specific format
    'file name extention .xlsm

    'newFileName = _
        'newFilePath _
        '& "Diversion Data Set.xlsm"
    'If Len(Dir(newFileName)) Then Kill newFileName 'identifies if file name already exists and deletes as SaveCopyAs doesn't allow overwrite


    'ThisWorkbook.SaveCopyAs _
            'Filename:=newFileName
    'OverwriteExisting = True

    'Open copied workbook to begin filtering

    'Workbooks.Open Filename:="H:\SJS Analysis\JeanMarie\Diversion Data Set\Diversion Data Set.xlsm", Password:="DeathStar_911"
    'Application.ScreenUpdating = False
    'Application.Calculation = xlCalculationManual
    'Application.DisplayAlerts = False
    'Worksheets("Entry").Activate

    'Static arrest data paste (for diversion comparison)

    'Range(hFindTwo("Arrest Date", "PETITION") & 3, hFindTwo("Arrest Date", "PETITION") & 25000).Select
    'Selection.Copy
    'Range(hFindTwo("Total Arrests", "STATIC DATA FOR PASTE") & 3, hFindTwo("Total Arrests", "STATIC DATA FOR PASTE") & 3).Select
    'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    'Range(hFindTwo("Sex", "DEMOGRAPHICS") & 3, hFindTwo("RACE", "DEMOGRAPHICS") & 25000).Select
    'Selection.Copy
    'Range(hFindTwo("Total Gender", "STATIC DATA FOR PASTE") & 3, hFindTwo("Total Gender", "STATIC DATA FOR PASTE") & 3).Select
    'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        ':=False, Transpose:=False
    'Range(hFindTwo("Charge Grade (specific)", "PETITION") & 3, hFindTwo("Charge Grade (broad)", "PETITION") & 25000).Select
    'Selection.Copy
    'Range(hFindTwo("Charge Grade Specific", "STATIC DATA FOR PASTE") & 3, hFindTwo("Charge Grade Specific", "STATIC DATA FOR PASTE") & 3).Select
    'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        ':=False, Transpose:=False
    'Range(hFindTwo("Charge Category", "PETITION") & 3, hFindTwo("Charge Category", "PETITION") & 25000).Select
    'Selection.Copy
    'Range(hFindTwo("Charge Group", "STATIC DATA FOR PASTE") & 3, hFindTwo("Charge Group", "STATIC DATA FOR PASTE") & 3).Select
    'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        ':=False, Transpose:=False
    'Range(hFindTwo("Arresting District", "PETITION") & 3, hFindTwo("Arresting District", "PETITION") & 25000).Select
    'Selection.Copy
    'Range(hFindTwo("Arresting District", "STATIC DATA FOR PASTE") & 3, hFindTwo("Arresting District", "STATIC DATA FOR PASTE") & 3).Select
    'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        ':=False, Transpose:=False

    'Filter and clear non-diversion data (while leaving aggregate, static arrest data for comparison)
    'Range((headerFindTwo("Navigation Tools") & 2), hFindTwo("END") & 25000).AutoFilter Field:=(hFindNumTwo("Referred to Diversion?", "DESCRIPTIVES", "DIVERSION")), Criteria1:="2"
    'Range((headerFindTwo("Navigation Tools") & 3), headerFindTwo("END") & 25000).Select
    'Selection.ClearContents
    'Range((headerFindTwo("Navigation Tools") & 2), hFindTwo("END") & 25000).AutoFilter
    'Range("A2").Select

    'Hides spreadsheet and locks workbook structure for editing

    'Sheets("Entry").Visible = False
    'Sheets("User Entry").Activate
    'Range("N11").Select
    'Workbooks("Diversion Data Set.xlsm").Protect Structure:=True, Password:="capstone125", Windows:=False


    'Saves file and password protectes entry

    'ActiveWorkbook.SaveAs Filename:="H:\SJS Analysis\JeanMarie\Diversion Data Set\Diversion Data Set.xlsm", Password:="capstone121"
    'OverwriteExisting = True

    'Saves file to rest of diversion team with different passwords then close
    'ActiveWorkbook.SaveAs Filename:="H:\SJS Analysis\Faith\Diversion Data Set\Diversion Data Set.xlsm", Password:="capstone121"
    'OverwriteExisting = True




    'ActiveWorkbook.Close


    'MASTER COPIES

    'OREN EXPORT

    Dim newFilePath2 As String
    Dim newFileName2 As String


    newFilePath2 = "H:\SJS Analysis\Oren\Aggregate Data Set\"
    newFileName2 = _
        newFilePath2 _
        & "Aggregate Data Set.xlsm"
    If Len(Dir(newFileName2)) Then Kill newFileName2 'identifies if file name already exists and deletes as SaveCopyAs doesn't allow overwrite

    ThisWorkbook.SaveCopyAs _
            Filename:=newFileName2
    OverwriteExisting = True

    Workbooks.Open Filename:="H:\SJS Analysis\Oren\Aggregate Data Set\Aggregate Data Set.xlsm", Password:="DeathStar_911"
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:="H:\SJS Analysis\Oren\Aggregate Data Set\Aggregate Data Set.xlsm", Password:="DeathStar_911"
    OverwriteExisting = True
    ActiveWorkbook.Close

    'Adam EXPORT

    Dim newFilePath3 As String
    Dim newFileName3 As String


    newFilePath3 = "H:\SJS Analysis\Adam\Aggregate Data Set\"
    newFileName3 = _
        newFilePath3 _
        & "Aggregate Data Set.xlsm"
    If Len(Dir(newFileName3)) Then Kill newFileName3 'identifies if file name already exists and deletes as SaveCopyAs doesn't allow overwrite

    ThisWorkbook.SaveCopyAs _
            Filename:=newFileName3
    OverwriteExisting = True

    Workbooks.Open Filename:="H:\SJS Analysis\Adam\Aggregate Data Set\Aggregate Data Set.xlsm", Password:="DeathStar_911"
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:="H:\SJS Analysis\Adam\Aggregate Data Set\Aggregate Data Set.xlsm", Password:="DeathStar_911"
    OverwriteExisting = True
    ActiveWorkbook.Close



    'Karli EXPORT

    Dim newFilePath4 As String
    Dim newFileName4 As String


    newFilePath4 = "H:\SJS Analysis\Karli\Aggregate Data Set\"
    newFileName4 = _
        newFilePath4 _
        & "Aggregate Data Set.xlsm"
    If Len(Dir(newFileName4)) Then Kill newFileName4 'identifies if file name already exists and deletes as SaveCopyAs doesn't allow overwrite

    ThisWorkbook.SaveCopyAs _
            Filename:=newFileName4
    OverwriteExisting = True

    Workbooks.Open Filename:="H:\SJS Analysis\Karli\Aggregate Data Set\Aggregate Data Set.xlsm", Password:="DeathStar_911"
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:="H:\SJS Analysis\Karli\Aggregate Data Set\Aggregate Data Set.xlsm", Password:="DeathStar_911"
    OverwriteExisting = True
    ActiveWorkbook.Close



    'Mike EXPORT

    Dim newFilePath5 As String
    Dim newFileName5 As String


    newFilePath5 = "H:\SJS Analysis\Mike\Aggregate Data Set\"
    newFileName5 = _
        newFilePath5 _
        & "Aggregate Data Set.xlsm"
    If Len(Dir(newFileName5)) Then Kill newFileName5 'identifies if file name already exists and deletes as SaveCopyAs doesn't allow overwrite

    ThisWorkbook.SaveCopyAs _
            Filename:=newFileName5
    OverwriteExisting = True

    Workbooks.Open Filename:="H:\SJS Analysis\Mike\Aggregate Data Set\Aggregate Data Set.xlsm", Password:="DeathStar_911"
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:="H:\SJS Analysis\Mike\Aggregate Data Set\Aggregate Data Set.xlsm", Password:="DeathStar_911"
    OverwriteExisting = True
    ActiveWorkbook.Close



    'Moshe EXPORT

    Dim newFilePath6 As String
    Dim newFileName6 As String


    newFilePath6 = "H:\SJS Analysis\Moshe\Aggregate Data Set\"
    newFileName6 = _
        newFilePath6 _
        & "Aggregate Data Set.xlsm"
    If Len(Dir(newFileName6)) Then Kill newFileName6 'identifies if file name already exists and deletes as SaveCopyAs doesn't allow overwrite

    ThisWorkbook.SaveCopyAs _
            Filename:=newFileName6
    OverwriteExisting = True

    Workbooks.Open Filename:="H:\SJS Analysis\Moshe\Aggregate Data Set\Aggregate Data Set.xlsm", Password:="DeathStar_911"
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:="H:\SJS Analysis\Moshe\Aggregate Data Set\Aggregate Data Set.xlsm", Password:="DeathStar_911"
    OverwriteExisting = True
    ActiveWorkbook.Close


    'Ebony EXPORT

    Dim newFilePath7 As String
    Dim newFileName7 As String


    newFilePath7 = "H:\SJS Analysis\Ebony\Aggregate Data Set\"
    newFileName7 = _
        newFilePath7 _
        & "Aggregate Data Set.xlsm"
    If Len(Dir(newFileName7)) Then Kill newFileName7 'identifies if file name already exists and deletes as SaveCopyAs doesn't allow overwrite

    ThisWorkbook.SaveCopyAs _
            Filename:=newFileName7
    OverwriteExisting = True

    Workbooks.Open Filename:="H:\SJS Analysis\Ebony\Aggregate Data Set\Aggregate Data Set.xlsm", Password:="DeathStar_911"
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:="H:\SJS Analysis\Ebony\Aggregate Data Set\Aggregate Data Set.xlsm", Password:="DeathStar_911"
    OverwriteExisting = True
    ActiveWorkbook.Close

    'Madeline EXPORT

    Dim newFilePath8 As String
    Dim newFileName8 As String


    newFilePath8 = "H:\SJS Analysis\Madeline\Aggregate Data Set\"
    newFileName8 = _
        newFilePath8 _
        & "Aggregate Data Set.xlsm"
    If Len(Dir(newFileName8)) Then Kill newFileName8 'identifies if file name already exists and deletes as SaveCopyAs doesn't allow overwrite

    ThisWorkbook.SaveCopyAs _
            Filename:=newFileName8
    OverwriteExisting = True

    Workbooks.Open Filename:="H:\SJS Analysis\Madeline\Aggregate Data Set\Aggregate Data Set.xlsm", Password:="DeathStar_911"
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:="H:\SJS Analysis\Madeline\Aggregate Data Set\Aggregate Data Set.xlsm", Password:="DeathStar_911"
    OverwriteExisting = True
    ActiveWorkbook.Close


    'Wes EXPORT

    Dim newFilePath9 As String
    Dim newFileName9 As String


    newFilePath9 = "H:\SJS Analysis\Wes\Aggregate Data Set\"
    newFileName9 = _
        newFilePath9 _
        & "Aggregate Data Set.xlsm"
    If Len(Dir(newFileName9)) Then Kill newFileName9 'identifies if file name already exists and deletes as SaveCopyAs doesn't allow overwrite

    ThisWorkbook.SaveCopyAs _
            Filename:=newFileName9
    OverwriteExisting = True

    Workbooks.Open Filename:="H:\SJS Analysis\Wes\Aggregate Data Set\Aggregate Data Set.xlsm", Password:="DeathStar_911"
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:="H:\SJS Analysis\Wes\Aggregate Data Set\Aggregate Data Set.xlsm", Password:="DeathStar_911"
    OverwriteExisting = True
    ActiveWorkbook.Close

'Aditi EXPORT

    Dim newFilePath10 As String
    Dim newFileName10 As String


    newFilePath10 = "H:\SJS Analysis\Aditi\Aggregate Data Set\"
    newFileName10 = _
        newFilePath10 _
        & "Aggregate Data Set.xlsm"
    If Len(Dir(newFileName10)) Then Kill newFileName10 'identifies if file name already exists and deletes as SaveCopyAs doesn't allow overwrite

    ThisWorkbook.SaveCopyAs _
            Filename:=newFileName10
    OverwriteExisting = True

    Workbooks.Open Filename:="H:\SJS Analysis\Aditi\Aggregate Data Set\Aggregate Data Set.xlsm", Password:="DeathStar_911"
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:="H:\SJS Analysis\Aditi\Aggregate Data Set\Aggregate Data Set.xlsm", Password:="DeathStar_911"
    OverwriteExisting = True
    ActiveWorkbook.Close

'Christian EXPORT

    Dim newFilePath11 As String
    Dim newFileName11 As String


    newFilePath11 = "H:\SJS Analysis\Christian\Aggregate Data Set\"
    newFileName11 = _
        newFilePath11 _
        & "Aggregate Data Set.xlsm"
    If Len(Dir(newFileName11)) Then Kill newFileName11 'identifies if file name already exists and deletes as SaveCopyAs doesn't allow overwrite

    ThisWorkbook.SaveCopyAs _
            Filename:=newFileName11
    OverwriteExisting = True

    Workbooks.Open Filename:="H:\SJS Analysis\Christian\Aggregate Data Set\Aggregate Data Set.xlsm", Password:="DeathStar_911"
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:="H:\SJS Analysis\Christian\Aggregate Data Set\Aggregate Data Set.xlsm", Password:="DeathStar_911"
    OverwriteExisting = True
    ActiveWorkbook.Close



End Sub

Sub ExportDataQuickFile()
    
    'Dim rowNum As Long
    'Dim lastRow As Long
    
    'lastRow = Worksheets("Entry").Range("C" & Rows.count).End(xlUp).row

    'For rowNum = 3 To lastRow
        'Call AggAggSupervisionsAndConditions(rowNum)
    'Next rowNum
    
    
    'ARCHIVE BEFORE EXPORTING
    'declare the string variables we will need for the file name and file path
    Dim newFileName As String
    Dim newFilePath As String

    'used to access a function that will get us the current name of this workbook
    Dim fso As New Scripting.FileSystemObject

    'here you can set the name of the directory where you want the file to go.
    'note that the folder location must exist prior to saving it there
    newFilePath = "H:\SJS Entry\Archives\"

    'Here the new file name is composed of:
    'the filepath specified above
    'the current name of the workbook
    '(space)
    'this second in datetime in a specific format
    'file name extention .xlsm
    newFileName = _
        newFilePath _
        & fso.GetBaseName(ThisWorkbook.name) _
        & " " _
        & VBA.format(Now(), "yyyy-MM-dd hh.mm.ss") _
        & ".xlsm"


    ThisWorkbook.SaveCopyAs _
            Filename:=newFileName

    'EXPORTING COPIES THAT NEED FILTERING FOR ACCESS FIRST

    'DIVERSION TEAM

    'JEANMARIE EXPORT (Master Diversion Export - will trigger all other saveas options)

    'declare the string variables we will need for the file name and file path


    'used to access a function that will get us the current name of this workbook


    'here you can set the name of the directory where you want the file to go.
    'note that the folder location must exist prior to saving it there
    'newFilePath = "H:\SJS Analysis\JeanMarie\Diversion Data Set\"

    'Here the new file name is composed of:
    'the filepath specified above
    'the current name of the workbook
    '(space)
    'this second in datetime in a specific format
    'file name extention .xlsm

    'newFileName = _
        'newFilePath _
        '& "Diversion Data Set.xlsm"
    'If Len(Dir(newFileName)) Then Kill newFileName 'identifies if file name already exists and deletes as SaveCopyAs doesn't allow overwrite


    'ThisWorkbook.SaveCopyAs _
            'Filename:=newFileName
    'OverwriteExisting = True

    'Open copied workbook to begin filtering

    'Workbooks.Open Filename:="H:\SJS Analysis\JeanMarie\Diversion Data Set\Diversion Data Set.xlsm", Password:="DeathStar_911"
    'Application.ScreenUpdating = False
    'Application.Calculation = xlCalculationManual
    'Application.DisplayAlerts = False
    'Worksheets("Entry").Activate

    'Static arrest data paste (for diversion comparison)

    'Range(hFindTwo("Arrest Date", "PETITION") & 3, hFindTwo("Arrest Date", "PETITION") & 25000).Select
    'Selection.Copy
    'Range(hFindTwo("Total Arrests", "STATIC DATA FOR PASTE") & 3, hFindTwo("Total Arrests", "STATIC DATA FOR PASTE") & 3).Select
    'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    'Range(hFindTwo("Sex", "DEMOGRAPHICS") & 3, hFindTwo("RACE", "DEMOGRAPHICS") & 25000).Select
    'Selection.Copy
    'Range(hFindTwo("Total Gender", "STATIC DATA FOR PASTE") & 3, hFindTwo("Total Gender", "STATIC DATA FOR PASTE") & 3).Select
    'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        ':=False, Transpose:=False
    'Range(hFindTwo("Charge Grade (specific)", "PETITION") & 3, hFindTwo("Charge Grade (broad)", "PETITION") & 25000).Select
    'Selection.Copy
    'Range(hFindTwo("Charge Grade Specific", "STATIC DATA FOR PASTE") & 3, hFindTwo("Charge Grade Specific", "STATIC DATA FOR PASTE") & 3).Select
    'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        ':=False, Transpose:=False
    'Range(hFindTwo("Charge Category", "PETITION") & 3, hFindTwo("Charge Category", "PETITION") & 25000).Select
    'Selection.Copy
    'Range(hFindTwo("Charge Group", "STATIC DATA FOR PASTE") & 3, hFindTwo("Charge Group", "STATIC DATA FOR PASTE") & 3).Select
    'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        ':=False, Transpose:=False
    'Range(hFindTwo("Arresting District", "PETITION") & 3, hFindTwo("Arresting District", "PETITION") & 25000).Select
    'Selection.Copy
    'Range(hFindTwo("Arresting District", "STATIC DATA FOR PASTE") & 3, hFindTwo("Arresting District", "STATIC DATA FOR PASTE") & 3).Select
    'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        ':=False, Transpose:=False

    'Filter and clear non-diversion data (while leaving aggregate, static arrest data for comparison)
    'Range((headerFindTwo("Navigation Tools") & 2), hFindTwo("END") & 25000).AutoFilter Field:=(hFindNumTwo("Referred to Diversion?", "DESCRIPTIVES", "DIVERSION")), Criteria1:="2"
    'Range((headerFindTwo("Navigation Tools") & 3), headerFindTwo("END") & 25000).Select
    'Selection.ClearContents
    'Range((headerFindTwo("Navigation Tools") & 2), hFindTwo("END") & 25000).AutoFilter
    'Range("A2").Select

    'Hides spreadsheet and locks workbook structure for editing

    'Sheets("Entry").Visible = False
    'Sheets("User Entry").Activate
    'Range("N11").Select
    'Workbooks("Diversion Data Set.xlsm").Protect Structure:=True, Password:="capstone125", Windows:=False


    'Saves file and password protectes entry

    'ActiveWorkbook.SaveAs Filename:="H:\SJS Analysis\JeanMarie\Diversion Data Set\Diversion Data Set.xlsm", Password:="capstone121"
    'OverwriteExisting = True

    'Saves file to rest of diversion team with different passwords then close
    'ActiveWorkbook.SaveAs Filename:="H:\SJS Analysis\Faith\Diversion Data Set\Diversion Data Set.xlsm", Password:="capstone121"
    'OverwriteExisting = True




    'ActiveWorkbook.Close


    'MASTER COPIES

    'OREN EXPORT

    Dim newFilePath2 As String
    Dim newFileName2 As String


    newFilePath2 = "H:\SJS Analysis\Oren\Aggregate Data Set\"
    newFileName2 = _
        newFilePath2 _
        & "Aggregate Data Set.xlsm"
    If Len(Dir(newFileName2)) Then Kill newFileName2 'identifies if file name already exists and deletes as SaveCopyAs doesn't allow overwrite

    ThisWorkbook.SaveCopyAs _
            Filename:=newFileName2
    OverwriteExisting = True

    Workbooks.Open Filename:="H:\SJS Analysis\Oren\Aggregate Data Set\Aggregate Data Set.xlsm", Password:="DeathStar_911"
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:="H:\SJS Analysis\Oren\Aggregate Data Set\Aggregate Data Set.xlsm", Password:="DeathStar_911"
    OverwriteExisting = True
    ActiveWorkbook.Close

    'Adam EXPORT

    Dim newFilePath3 As String
    Dim newFileName3 As String


    newFilePath3 = "H:\SJS Analysis\Adam\Aggregate Data Set\"
    newFileName3 = _
        newFilePath3 _
        & "Aggregate Data Set.xlsm"
    If Len(Dir(newFileName3)) Then Kill newFileName3 'identifies if file name already exists and deletes as SaveCopyAs doesn't allow overwrite

    ThisWorkbook.SaveCopyAs _
            Filename:=newFileName3
    OverwriteExisting = True

    Workbooks.Open Filename:="H:\SJS Analysis\Adam\Aggregate Data Set\Aggregate Data Set.xlsm", Password:="DeathStar_911"
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:="H:\SJS Analysis\Adam\Aggregate Data Set\Aggregate Data Set.xlsm", Password:="DeathStar_911"
    OverwriteExisting = True
    ActiveWorkbook.Close



    'Karli EXPORT

    Dim newFilePath4 As String
    Dim newFileName4 As String


    newFilePath4 = "H:\SJS Analysis\Karli\Aggregate Data Set\"
    newFileName4 = _
        newFilePath4 _
        & "Aggregate Data Set.xlsm"
    If Len(Dir(newFileName4)) Then Kill newFileName4 'identifies if file name already exists and deletes as SaveCopyAs doesn't allow overwrite

    ThisWorkbook.SaveCopyAs _
            Filename:=newFileName4
    OverwriteExisting = True

    Workbooks.Open Filename:="H:\SJS Analysis\Karli\Aggregate Data Set\Aggregate Data Set.xlsm", Password:="DeathStar_911"
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:="H:\SJS Analysis\Karli\Aggregate Data Set\Aggregate Data Set.xlsm", Password:="DeathStar_911"
    OverwriteExisting = True
    ActiveWorkbook.Close



    'Mike EXPORT

    Dim newFilePath5 As String
    Dim newFileName5 As String


    newFilePath5 = "H:\SJS Analysis\Mike\Aggregate Data Set\"
    newFileName5 = _
        newFilePath5 _
        & "Aggregate Data Set.xlsm"
    If Len(Dir(newFileName5)) Then Kill newFileName5 'identifies if file name already exists and deletes as SaveCopyAs doesn't allow overwrite

    ThisWorkbook.SaveCopyAs _
            Filename:=newFileName5
    OverwriteExisting = True

    Workbooks.Open Filename:="H:\SJS Analysis\Mike\Aggregate Data Set\Aggregate Data Set.xlsm", Password:="DeathStar_911"
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:="H:\SJS Analysis\Mike\Aggregate Data Set\Aggregate Data Set.xlsm", Password:="DeathStar_911"
    OverwriteExisting = True
    ActiveWorkbook.Close



    'Moshe EXPORT

    Dim newFilePath6 As String
    Dim newFileName6 As String


    newFilePath6 = "H:\SJS Analysis\Moshe\Aggregate Data Set\"
    newFileName6 = _
        newFilePath6 _
        & "Aggregate Data Set.xlsm"
    If Len(Dir(newFileName6)) Then Kill newFileName6 'identifies if file name already exists and deletes as SaveCopyAs doesn't allow overwrite

    ThisWorkbook.SaveCopyAs _
            Filename:=newFileName6
    OverwriteExisting = True

    Workbooks.Open Filename:="H:\SJS Analysis\Moshe\Aggregate Data Set\Aggregate Data Set.xlsm", Password:="DeathStar_911"
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:="H:\SJS Analysis\Moshe\Aggregate Data Set\Aggregate Data Set.xlsm", Password:="DeathStar_911"
    OverwriteExisting = True
    ActiveWorkbook.Close


    'Ebony EXPORT

    Dim newFilePath7 As String
    Dim newFileName7 As String


    newFilePath7 = "H:\SJS Analysis\Ebony\Aggregate Data Set\"
    newFileName7 = _
        newFilePath7 _
        & "Aggregate Data Set.xlsm"
    If Len(Dir(newFileName7)) Then Kill newFileName7 'identifies if file name already exists and deletes as SaveCopyAs doesn't allow overwrite

    ThisWorkbook.SaveCopyAs _
            Filename:=newFileName7
    OverwriteExisting = True

    Workbooks.Open Filename:="H:\SJS Analysis\Ebony\Aggregate Data Set\Aggregate Data Set.xlsm", Password:="DeathStar_911"
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:="H:\SJS Analysis\Ebony\Aggregate Data Set\Aggregate Data Set.xlsm", Password:="DeathStar_911"
    OverwriteExisting = True
    ActiveWorkbook.Close

    'Madeline EXPORT

    Dim newFilePath8 As String
    Dim newFileName8 As String


    newFilePath8 = "H:\SJS Analysis\Madeline\Aggregate Data Set\"
    newFileName8 = _
        newFilePath8 _
        & "Aggregate Data Set.xlsm"
    If Len(Dir(newFileName8)) Then Kill newFileName8 'identifies if file name already exists and deletes as SaveCopyAs doesn't allow overwrite

    ThisWorkbook.SaveCopyAs _
            Filename:=newFileName8
    OverwriteExisting = True

    Workbooks.Open Filename:="H:\SJS Analysis\Madeline\Aggregate Data Set\Aggregate Data Set.xlsm", Password:="DeathStar_911"
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:="H:\SJS Analysis\Madeline\Aggregate Data Set\Aggregate Data Set.xlsm", Password:="DeathStar_911"
    OverwriteExisting = True
    ActiveWorkbook.Close


    'Wes EXPORT

    Dim newFilePath9 As String
    Dim newFileName9 As String


    newFilePath9 = "H:\SJS Analysis\Wes\Aggregate Data Set\"
    newFileName9 = _
        newFilePath9 _
        & "Aggregate Data Set.xlsm"
    If Len(Dir(newFileName9)) Then Kill newFileName9 'identifies if file name already exists and deletes as SaveCopyAs doesn't allow overwrite

    ThisWorkbook.SaveCopyAs _
            Filename:=newFileName9
    OverwriteExisting = True

    Workbooks.Open Filename:="H:\SJS Analysis\Wes\Aggregate Data Set\Aggregate Data Set.xlsm", Password:="DeathStar_911"
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:="H:\SJS Analysis\Wes\Aggregate Data Set\Aggregate Data Set.xlsm", Password:="DeathStar_911"
    OverwriteExisting = True
    ActiveWorkbook.Close

'Aditi EXPORT

    Dim newFilePath10 As String
    Dim newFileName10 As String


    newFilePath10 = "H:\SJS Analysis\Aditi\Aggregate Data Set\"
    newFileName10 = _
        newFilePath10 _
        & "Aggregate Data Set.xlsm"
    If Len(Dir(newFileName10)) Then Kill newFileName10 'identifies if file name already exists and deletes as SaveCopyAs doesn't allow overwrite

    ThisWorkbook.SaveCopyAs _
            Filename:=newFileName10
    OverwriteExisting = True

    Workbooks.Open Filename:="H:\SJS Analysis\Aditi\Aggregate Data Set\Aggregate Data Set.xlsm", Password:="DeathStar_911"
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:="H:\SJS Analysis\Aditi\Aggregate Data Set\Aggregate Data Set.xlsm", Password:="DeathStar_911"
    OverwriteExisting = True
    ActiveWorkbook.Close

'Christian EXPORT

    Dim newFilePath11 As String
    Dim newFileName11 As String


    newFilePath11 = "H:\SJS Analysis\Christian\Aggregate Data Set\"
    newFileName11 = _
        newFilePath11 _
        & "Aggregate Data Set.xlsm"
    If Len(Dir(newFileName11)) Then Kill newFileName11 'identifies if file name already exists and deletes as SaveCopyAs doesn't allow overwrite

    ThisWorkbook.SaveCopyAs _
            Filename:=newFileName11
    OverwriteExisting = True

    Workbooks.Open Filename:="H:\SJS Analysis\Christian\Aggregate Data Set\Aggregate Data Set.xlsm", Password:="DeathStar_911"
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:="H:\SJS Analysis\Christian\Aggregate Data Set\Aggregate Data Set.xlsm", Password:="DeathStar_911"
    OverwriteExisting = True
    ActiveWorkbook.Close


End Sub

