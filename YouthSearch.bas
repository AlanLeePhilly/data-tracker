Attribute VB_Name = "YouthSearch"
Sub ReturnToNavigationYouthSearch()
    Range("K5").Select
End Sub

Sub JumptoCourtroomHistory()
    Range("E81").Select
End Sub

Sub JumptoLegalStatusHistory()
    Range("T81").Select
End Sub
Sub JumptoSupervisionHistory()
    Range("F197").Select
End Sub
Sub JumptoConditionsHistory()
    Range("T197").Select
End Sub
Sub JumptoListingsHisory()
    Range("I483").Select
End Sub
Sub YouthSearchPrint0()

    Call RefreshNamedRanges
    Call Generate_Dictionaries

    Dim PrintSheet As Worksheet
    Dim DataSheet As Worksheet
    Set PrintSheet = Worksheets("Youth Search")
    Set DataSheet = Worksheets("Entry")

    Dim userRow As Long
    userRow = PrintSheet.Range("J5").value

    'identifiers
    PrintSheet.Range("D12").Select
    Selection.ClearContents
    PrintSheet.Range("D14").Select
    Selection.ClearContents
    PrintSheet.Range("I12").Select
    Selection.ClearContents
    PrintSheet.Range("I14").Select
    Selection.ClearContents

    'Demographics
    PrintSheet.Range("D19").Select
    Selection.ClearContents
    PrintSheet.Range("D21").Select
    Selection.ClearContents
    PrintSheet.Range("D23").Select
    Selection.ClearContents
    PrintSheet.Range("D26").Select
    Selection.ClearContents
    PrintSheet.Range("I19").Select
    Selection.ClearContents
    PrintSheet.Range("I21").Select
    Selection.ClearContents
    PrintSheet.Range("I23").Select
    Selection.ClearContents
    PrintSheet.Range("I26").Select
    Selection.ClearContents
    PrintSheet.Range("E28").Select
    Selection.ClearContents
    PrintSheet.Range("E30").Select
    Selection.ClearContents


    'Arrest Info
    PrintSheet.Range("O12").Select
    Selection.ClearContents
    PrintSheet.Range("P14").Select
    Selection.ClearContents
    PrintSheet.Range("P16").Select
    Selection.ClearContents
    PrintSheet.Range("P18").Select
    Selection.ClearContents
    PrintSheet.Range("P20").Select
    Selection.ClearContents
    PrintSheet.Range("N23").Select
    Selection.ClearContents

    'Petition Info
    PrintSheet.Range("T12").Select
    Selection.ClearContents
    PrintSheet.Range("T14").Select
    Selection.ClearContents
    PrintSheet.Range("T16").Select
    Selection.ClearContents
    PrintSheet.Range("T18").Select
    Selection.ClearContents
    PrintSheet.Range("Y12").Select
    Selection.ClearContents
    PrintSheet.Range("Y14").Select
    Selection.ClearContents
    PrintSheet.Range("Y16").Select
    Selection.ClearContents
    PrintSheet.Range("Y18").Select
    Selection.ClearContents
    PrintSheet.Range("Y20").Select
    Selection.ClearContents
    PrintSheet.Range("AC12").Select
    Selection.ClearContents
    PrintSheet.Range("AC14").Select
    Selection.ClearContents
    PrintSheet.Range("AC16").Select
    Selection.ClearContents
    PrintSheet.Range("AC18").Select
    Selection.ClearContents
    PrintSheet.Range("AC20").Select
    Selection.ClearContents
    PrintSheet.Range("T23").Select
    Selection.ClearContents
    PrintSheet.Range("T25").Select
    Selection.ClearContents
    PrintSheet.Range("T27").Select
    Selection.ClearContents
    PrintSheet.Range("T29").Select
    Selection.ClearContents
    PrintSheet.Range("Y23").Select
    Selection.ClearContents
    PrintSheet.Range("Y25").Select
    Selection.ClearContents
    PrintSheet.Range("Y27").Select
    Selection.ClearContents
    PrintSheet.Range("Y29").Select
    Selection.ClearContents
    PrintSheet.Range("Y31").Select
    Selection.ClearContents
    PrintSheet.Range("AC23").Select
    Selection.ClearContents
    PrintSheet.Range("AC25").Select
    Selection.ClearContents
    PrintSheet.Range("AC27").Select
    Selection.ClearContents
    PrintSheet.Range("AC29").Select
    Selection.ClearContents
    PrintSheet.Range("AC31").Select
    Selection.ClearContents
    PrintSheet.Range("T34").Select
    Selection.ClearContents
    PrintSheet.Range("T36").Select
    Selection.ClearContents
    PrintSheet.Range("T38").Select
    Selection.ClearContents
    PrintSheet.Range("T40").Select
    Selection.ClearContents
    PrintSheet.Range("Y34").Select
    Selection.ClearContents
    PrintSheet.Range("Y36").Select
    Selection.ClearContents
    PrintSheet.Range("Y38").Select
    Selection.ClearContents
    PrintSheet.Range("Y40").Select
    Selection.ClearContents
    PrintSheet.Range("Y42").Select
    Selection.ClearContents
    PrintSheet.Range("AC34").Select
    Selection.ClearContents
    PrintSheet.Range("AC36").Select
    Selection.ClearContents
    PrintSheet.Range("AC38").Select
    Selection.ClearContents
    PrintSheet.Range("AC40").Select
    Selection.ClearContents
    PrintSheet.Range("AC42").Select
    Selection.ClearContents
    PrintSheet.Range("T45").Select
    Selection.ClearContents
    PrintSheet.Range("T47").Select
    Selection.ClearContents
    PrintSheet.Range("T49").Select
    Selection.ClearContents
    PrintSheet.Range("T51").Select
    Selection.ClearContents
    PrintSheet.Range("Y45").Select
    Selection.ClearContents
    PrintSheet.Range("Y47").Select
    Selection.ClearContents
    PrintSheet.Range("Y49").Select
    Selection.ClearContents
    PrintSheet.Range("Y51").Select
    Selection.ClearContents
    PrintSheet.Range("Y53").Select
    Selection.ClearContents
    PrintSheet.Range("AC45").Select
    Selection.ClearContents
    PrintSheet.Range("AC47").Select
    Selection.ClearContents
    PrintSheet.Range("AC49").Select
    Selection.ClearContents
    PrintSheet.Range("AC51").Select
    Selection.ClearContents
    PrintSheet.Range("AC53").Select
    Selection.ClearContents


    'Associated Petitions
    PrintSheet.Range("O40").Select
    Selection.ClearContents
    PrintSheet.Range("O42").Select
    Selection.ClearContents
    PrintSheet.Range("O45").Select
    Selection.ClearContents
    PrintSheet.Range("O47").Select
    Selection.ClearContents
    PrintSheet.Range("O49").Select
    Selection.ClearContents
    PrintSheet.Range("O51").Select
    Selection.ClearContents
    PrintSheet.Range("O53").Select
    Selection.ClearContents
    PrintSheet.Range("O55").Select
    Selection.ClearContents

    'Status of Arrest
    PrintSheet.Range("D37").Select
    Selection.ClearContents
    PrintSheet.Range("D39").Select
    Selection.ClearContents
    PrintSheet.Range("D41").Select
    Selection.ClearContents
    PrintSheet.Range("D43").Select
    Selection.ClearContents
    PrintSheet.Range("G37").Select
    Selection.ClearContents
    PrintSheet.Range("G39").Select
    Selection.ClearContents
    PrintSheet.Range("G41").Select
    Selection.ClearContents
    PrintSheet.Range("J37").Select
    Selection.ClearContents
    PrintSheet.Range("J39").Select
    Selection.ClearContents
    PrintSheet.Range("J41").Select
    Selection.ClearContents

    'Active Court Proceedings
    PrintSheet.Range("D48").Select
    Selection.ClearContents
    PrintSheet.Range("D50").Select
    Selection.ClearContents
    PrintSheet.Range("G48").Select
    Selection.ClearContents
    PrintSheet.Range("G50").Select
    Selection.ClearContents
    PrintSheet.Range("J48").Select
    Selection.ClearContents
    PrintSheet.Range("J50").Select
    Selection.ClearContents


    'Active Supervision Programs
    PrintSheet.Range("D55").Select
    Selection.ClearContents
    PrintSheet.Range("D57").Select
    Selection.ClearContents
    PrintSheet.Range("D59").Select
    Selection.ClearContents
    PrintSheet.Range("G55").Select
    Selection.ClearContents
    PrintSheet.Range("G57").Select
    Selection.ClearContents
    PrintSheet.Range("G59").Select
    Selection.ClearContents
    PrintSheet.Range("J55").Select
    Selection.ClearContents
    PrintSheet.Range("J57").Select
    Selection.ClearContents
    PrintSheet.Range("J59").Select
    Selection.ClearContents


    'Active Conditions
    PrintSheet.Range("D64").Select
    Selection.ClearContents
    PrintSheet.Range("D66").Select
    Selection.ClearContents
    PrintSheet.Range("D68").Select
    Selection.ClearContents
    PrintSheet.Range("D70").Select
    Selection.ClearContents
    PrintSheet.Range("D72").Select
    Selection.ClearContents
    PrintSheet.Range("D74").Select
    Selection.ClearContents
    PrintSheet.Range("G64").Select
    Selection.ClearContents
    PrintSheet.Range("G66").Select
    Selection.ClearContents
    PrintSheet.Range("G68").Select
    Selection.ClearContents
    PrintSheet.Range("G70").Select
    Selection.ClearContents
    PrintSheet.Range("G72").Select
    Selection.ClearContents
    PrintSheet.Range("G74").Select
    Selection.ClearContents
    PrintSheet.Range("J64").Select
    Selection.ClearContents
    PrintSheet.Range("J66").Select
    Selection.ClearContents
    PrintSheet.Range("J68").Select
    Selection.ClearContents
    PrintSheet.Range("J70").Select
    Selection.ClearContents
    PrintSheet.Range("J72").Select
    Selection.ClearContents
    PrintSheet.Range("J74").Select
    Selection.ClearContents


    'Courtroom History
    PrintSheet.Range("D84").Select
    Selection.ClearContents
    PrintSheet.Range("G84").Select
    Selection.ClearContents
    PrintSheet.Range("J84").Select
    Selection.ClearContents
    PrintSheet.Range("N84").Select
    Selection.ClearContents
    PrintSheet.Range("D86").Select
    Selection.ClearContents
    PrintSheet.Range("D94").Select
    Selection.ClearContents
    PrintSheet.Range("G94").Select
    Selection.ClearContents
    PrintSheet.Range("J94").Select
    Selection.ClearContents
    PrintSheet.Range("N94").Select
    Selection.ClearContents
    PrintSheet.Range("D96").Select
    Selection.ClearContents
    PrintSheet.Range("D104").Select
    Selection.ClearContents
    PrintSheet.Range("G104").Select
    Selection.ClearContents
    PrintSheet.Range("J104").Select
    Selection.ClearContents
    PrintSheet.Range("N104").Select
    Selection.ClearContents
    PrintSheet.Range("D106").Select
    Selection.ClearContents
    PrintSheet.Range("D114").Select
    Selection.ClearContents
    PrintSheet.Range("G114").Select
    Selection.ClearContents
    PrintSheet.Range("J114").Select
    Selection.ClearContents
    PrintSheet.Range("N114").Select
    Selection.ClearContents
    PrintSheet.Range("D116").Select
    Selection.ClearContents
    PrintSheet.Range("D124").Select
    Selection.ClearContents
    PrintSheet.Range("G124").Select
    Selection.ClearContents
    PrintSheet.Range("J124").Select
    Selection.ClearContents
    PrintSheet.Range("N124").Select
    Selection.ClearContents
    PrintSheet.Range("D126").Select
    Selection.ClearContents
    PrintSheet.Range("D134").Select
    Selection.ClearContents
    PrintSheet.Range("G134").Select
    Selection.ClearContents
    PrintSheet.Range("J134").Select
    Selection.ClearContents
    PrintSheet.Range("N134").Select
    Selection.ClearContents
    PrintSheet.Range("D136").Select
    Selection.ClearContents
    PrintSheet.Range("D144").Select
    Selection.ClearContents
    PrintSheet.Range("G144").Select
    Selection.ClearContents
    PrintSheet.Range("J144").Select
    Selection.ClearContents
    PrintSheet.Range("N144").Select
    Selection.ClearContents
    PrintSheet.Range("D146").Select
    Selection.ClearContents
    PrintSheet.Range("D154").Select
    Selection.ClearContents
    PrintSheet.Range("G154").Select
    Selection.ClearContents
    PrintSheet.Range("J154").Select
    Selection.ClearContents
    PrintSheet.Range("N154").Select
    Selection.ClearContents
    PrintSheet.Range("D156").Select
    Selection.ClearContents
    PrintSheet.Range("D164").Select
    Selection.ClearContents
    PrintSheet.Range("G164").Select
    Selection.ClearContents
    PrintSheet.Range("J164").Select
    Selection.ClearContents
    PrintSheet.Range("N164").Select
    Selection.ClearContents
    PrintSheet.Range("D166").Select
    Selection.ClearContents
    PrintSheet.Range("D174").Select
    Selection.ClearContents
    PrintSheet.Range("G174").Select
    Selection.ClearContents
    PrintSheet.Range("J174").Select
    Selection.ClearContents
    PrintSheet.Range("N174").Select
    Selection.ClearContents
    PrintSheet.Range("D176").Select
    Selection.ClearContents
    PrintSheet.Range("D184").Select
    Selection.ClearContents
    PrintSheet.Range("G184").Select
    Selection.ClearContents
    PrintSheet.Range("J184").Select
    Selection.ClearContents
    PrintSheet.Range("N184").Select
    Selection.ClearContents
    PrintSheet.Range("D186").Select
    Selection.ClearContents


    'Legal Status History

    PrintSheet.Range("S84").Select
    Selection.ClearContents
    PrintSheet.Range("V84").Select
    Selection.ClearContents
    PrintSheet.Range("Z84").Select
    Selection.ClearContents
    PrintSheet.Range("AC84").Select
    Selection.ClearContents
    PrintSheet.Range("V86").Select
    Selection.ClearContents
    PrintSheet.Range("Z86").Select
    Selection.ClearContents
    PrintSheet.Range("V88").Select
    Selection.ClearContents
    PrintSheet.Range("S96").Select
    Selection.ClearContents
    PrintSheet.Range("V96").Select
    Selection.ClearContents
    PrintSheet.Range("Z96").Select
    Selection.ClearContents
    PrintSheet.Range("AC96").Select
    Selection.ClearContents
    PrintSheet.Range("V98").Select
    Selection.ClearContents
    PrintSheet.Range("Z98").Select
    Selection.ClearContents
    PrintSheet.Range("V100").Select
    Selection.ClearContents
    PrintSheet.Range("S108").Select
    Selection.ClearContents
    PrintSheet.Range("V108").Select
    Selection.ClearContents
    PrintSheet.Range("Z108").Select
    Selection.ClearContents
    PrintSheet.Range("AC108").Select
    Selection.ClearContents
    PrintSheet.Range("V110").Select
    Selection.ClearContents
    PrintSheet.Range("Z110").Select
    Selection.ClearContents
    PrintSheet.Range("V112").Select
    Selection.ClearContents
    PrintSheet.Range("S120").Select
    Selection.ClearContents
    PrintSheet.Range("V120").Select
    Selection.ClearContents
    PrintSheet.Range("Z120").Select
    Selection.ClearContents
    PrintSheet.Range("AC120").Select
    Selection.ClearContents
    PrintSheet.Range("V122").Select
    Selection.ClearContents
    PrintSheet.Range("Z122").Select
    Selection.ClearContents
    PrintSheet.Range("V124").Select
    Selection.ClearContents
    PrintSheet.Range("S132").Select
    Selection.ClearContents
    PrintSheet.Range("V132").Select
    Selection.ClearContents
    PrintSheet.Range("Z132").Select
    Selection.ClearContents
    PrintSheet.Range("AC132").Select
    Selection.ClearContents
    PrintSheet.Range("V134").Select
    Selection.ClearContents
    PrintSheet.Range("Z134").Select
    Selection.ClearContents
    PrintSheet.Range("V136").Select
    Selection.ClearContents
    PrintSheet.Range("S144").Select
    Selection.ClearContents
    PrintSheet.Range("V144").Select
    Selection.ClearContents
    PrintSheet.Range("Z144").Select
    Selection.ClearContents
    PrintSheet.Range("AC144").Select
    Selection.ClearContents
    PrintSheet.Range("V146").Select
    Selection.ClearContents
    PrintSheet.Range("Z146").Select
    Selection.ClearContents
    PrintSheet.Range("V148").Select
    Selection.ClearContents
    PrintSheet.Range("S156").Select
    Selection.ClearContents
    PrintSheet.Range("V156").Select
    Selection.ClearContents
    PrintSheet.Range("Z156").Select
    Selection.ClearContents
    PrintSheet.Range("AC156").Select
    Selection.ClearContents
    PrintSheet.Range("V158").Select
    Selection.ClearContents
    PrintSheet.Range("Z158").Select
    Selection.ClearContents
    PrintSheet.Range("V160").Select
    Selection.ClearContents



    'Supervision History
    PrintSheet.Range("D200").Select
    Selection.ClearContents
    PrintSheet.Range("G200").Select
    Selection.ClearContents
    PrintSheet.Range("K200").Select
    Selection.ClearContents
    PrintSheet.Range("N200").Select
    Selection.ClearContents
    PrintSheet.Range("G202").Select
    Selection.ClearContents
    PrintSheet.Range("K202").Select
    Selection.ClearContents
    PrintSheet.Range("G204").Select
    Selection.ClearContents
    PrintSheet.Range("K204").Select
    Selection.ClearContents
    PrintSheet.Range("E206").Select
    Selection.ClearContents
    PrintSheet.Range("D214").Select
    Selection.ClearContents
    PrintSheet.Range("G214").Select
    Selection.ClearContents
    PrintSheet.Range("K214").Select
    Selection.ClearContents
    PrintSheet.Range("N214").Select
    Selection.ClearContents
    PrintSheet.Range("G216").Select
    Selection.ClearContents
    PrintSheet.Range("K216").Select
    Selection.ClearContents
    PrintSheet.Range("G218").Select
    Selection.ClearContents
    PrintSheet.Range("K218").Select
    Selection.ClearContents
    PrintSheet.Range("E220").Select
    Selection.ClearContents
    PrintSheet.Range("D228").Select
    Selection.ClearContents
    PrintSheet.Range("G228").Select
    Selection.ClearContents
    PrintSheet.Range("K228").Select
    Selection.ClearContents
    PrintSheet.Range("N228").Select
    Selection.ClearContents
    PrintSheet.Range("G230").Select
    Selection.ClearContents
    PrintSheet.Range("K230").Select
    Selection.ClearContents
    PrintSheet.Range("G232").Select
    Selection.ClearContents
    PrintSheet.Range("K232").Select
    Selection.ClearContents
    PrintSheet.Range("E234").Select
    Selection.ClearContents
    PrintSheet.Range("D242").Select
    Selection.ClearContents
    PrintSheet.Range("G242").Select
    Selection.ClearContents
    PrintSheet.Range("K242").Select
    Selection.ClearContents
    PrintSheet.Range("N242").Select
    Selection.ClearContents
    PrintSheet.Range("G244").Select
    Selection.ClearContents
    PrintSheet.Range("K244").Select
    Selection.ClearContents
    PrintSheet.Range("G246").Select
    Selection.ClearContents
    PrintSheet.Range("K246").Select
    Selection.ClearContents
    PrintSheet.Range("E248").Select
    Selection.ClearContents
    PrintSheet.Range("D256").Select
    Selection.ClearContents
    PrintSheet.Range("G256").Select
    Selection.ClearContents
    PrintSheet.Range("K256").Select
    Selection.ClearContents
    PrintSheet.Range("N256").Select
    Selection.ClearContents
    PrintSheet.Range("G258").Select
    Selection.ClearContents
    PrintSheet.Range("K258").Select
    Selection.ClearContents
    PrintSheet.Range("G260").Select
    Selection.ClearContents
    PrintSheet.Range("K260").Select
    Selection.ClearContents
    PrintSheet.Range("E262").Select
    Selection.ClearContents
    PrintSheet.Range("D270").Select
    Selection.ClearContents
    PrintSheet.Range("G270").Select
    Selection.ClearContents
    PrintSheet.Range("K270").Select
    Selection.ClearContents
    PrintSheet.Range("N270").Select
    Selection.ClearContents
    PrintSheet.Range("G272").Select
    Selection.ClearContents
    PrintSheet.Range("K272").Select
    Selection.ClearContents
    PrintSheet.Range("G274").Select
    Selection.ClearContents
    PrintSheet.Range("K274").Select
    Selection.ClearContents
    PrintSheet.Range("E276").Select
    Selection.ClearContents
    PrintSheet.Range("D284").Select
    Selection.ClearContents
    PrintSheet.Range("G284").Select
    Selection.ClearContents
    PrintSheet.Range("K284").Select
    Selection.ClearContents
    PrintSheet.Range("N284").Select
    Selection.ClearContents
    PrintSheet.Range("G286").Select
    Selection.ClearContents
    PrintSheet.Range("K286").Select
    Selection.ClearContents
    PrintSheet.Range("G288").Select
    Selection.ClearContents
    PrintSheet.Range("K288").Select
    Selection.ClearContents
    PrintSheet.Range("E290").Select
    Selection.ClearContents
    PrintSheet.Range("D298").Select
    Selection.ClearContents
    PrintSheet.Range("G298").Select
    Selection.ClearContents
    PrintSheet.Range("K298").Select
    Selection.ClearContents
    PrintSheet.Range("N298").Select
    Selection.ClearContents
    PrintSheet.Range("G300").Select
    Selection.ClearContents
    PrintSheet.Range("K300").Select
    Selection.ClearContents
    PrintSheet.Range("G302").Select
    Selection.ClearContents
    PrintSheet.Range("K302").Select
    Selection.ClearContents
    PrintSheet.Range("E304").Select
    Selection.ClearContents
    PrintSheet.Range("D312").Select
    Selection.ClearContents
    PrintSheet.Range("G312").Select
    Selection.ClearContents
    PrintSheet.Range("K312").Select
    Selection.ClearContents
    PrintSheet.Range("N312").Select
    Selection.ClearContents
    PrintSheet.Range("G314").Select
    Selection.ClearContents
    PrintSheet.Range("K314").Select
    Selection.ClearContents
    PrintSheet.Range("G316").Select
    Selection.ClearContents
    PrintSheet.Range("K316").Select
    Selection.ClearContents
    PrintSheet.Range("E318").Select
    Selection.ClearContents
    PrintSheet.Range("D326").Select
    Selection.ClearContents
    PrintSheet.Range("G326").Select
    Selection.ClearContents
    PrintSheet.Range("K326").Select
    Selection.ClearContents
    PrintSheet.Range("N326").Select
    Selection.ClearContents
    PrintSheet.Range("G328").Select
    Selection.ClearContents
    PrintSheet.Range("K328").Select
    Selection.ClearContents
    PrintSheet.Range("G330").Select
    Selection.ClearContents
    PrintSheet.Range("K330").Select
    Selection.ClearContents
    PrintSheet.Range("E332").Select
    Selection.ClearContents
    PrintSheet.Range("D340").Select
    Selection.ClearContents
    PrintSheet.Range("G340").Select
    Selection.ClearContents
    PrintSheet.Range("K340").Select
    Selection.ClearContents
    PrintSheet.Range("N340").Select
    Selection.ClearContents
    PrintSheet.Range("G342").Select
    Selection.ClearContents
    PrintSheet.Range("K342").Select
    Selection.ClearContents
    PrintSheet.Range("G344").Select
    Selection.ClearContents
    PrintSheet.Range("K344").Select
    Selection.ClearContents
    PrintSheet.Range("E346").Select
    Selection.ClearContents
    PrintSheet.Range("D354").Select
    Selection.ClearContents
    PrintSheet.Range("G354").Select
    Selection.ClearContents
    PrintSheet.Range("K354").Select
    Selection.ClearContents
    PrintSheet.Range("N354").Select
    Selection.ClearContents
    PrintSheet.Range("G356").Select
    Selection.ClearContents
    PrintSheet.Range("K356").Select
    Selection.ClearContents
    PrintSheet.Range("G358").Select
    Selection.ClearContents
    PrintSheet.Range("K358").Select
    Selection.ClearContents
    PrintSheet.Range("E360").Select
    Selection.ClearContents
    PrintSheet.Range("D368").Select
    Selection.ClearContents
    PrintSheet.Range("G368").Select
    Selection.ClearContents
    PrintSheet.Range("K368").Select
    Selection.ClearContents
    PrintSheet.Range("N368").Select
    Selection.ClearContents
    PrintSheet.Range("G370").Select
    Selection.ClearContents
    PrintSheet.Range("K370").Select
    Selection.ClearContents
    PrintSheet.Range("G372").Select
    Selection.ClearContents
    PrintSheet.Range("K372").Select
    Selection.ClearContents
    PrintSheet.Range("E374").Select
    Selection.ClearContents
    PrintSheet.Range("D382").Select
    Selection.ClearContents
    PrintSheet.Range("G382").Select
    Selection.ClearContents
    PrintSheet.Range("K382").Select
    Selection.ClearContents
    PrintSheet.Range("N382").Select
    Selection.ClearContents
    PrintSheet.Range("G384").Select
    Selection.ClearContents
    PrintSheet.Range("K384").Select
    Selection.ClearContents
    PrintSheet.Range("G386").Select
    Selection.ClearContents
    PrintSheet.Range("K386").Select
    Selection.ClearContents
    PrintSheet.Range("E388").Select
    Selection.ClearContents
    PrintSheet.Range("D396").Select
    Selection.ClearContents
    PrintSheet.Range("G396").Select
    Selection.ClearContents
    PrintSheet.Range("K396").Select
    Selection.ClearContents
    PrintSheet.Range("N396").Select
    Selection.ClearContents
    PrintSheet.Range("G398").Select
    Selection.ClearContents
    PrintSheet.Range("K398").Select
    Selection.ClearContents
    PrintSheet.Range("G400").Select
    Selection.ClearContents
    PrintSheet.Range("K400").Select
    Selection.ClearContents
    PrintSheet.Range("E402").Select
    Selection.ClearContents
    PrintSheet.Range("D410").Select
    Selection.ClearContents
    PrintSheet.Range("G410").Select
    Selection.ClearContents
    PrintSheet.Range("K410").Select
    Selection.ClearContents
    PrintSheet.Range("N410").Select
    Selection.ClearContents
    PrintSheet.Range("G412").Select
    Selection.ClearContents
    PrintSheet.Range("K412").Select
    Selection.ClearContents
    PrintSheet.Range("G414").Select
    Selection.ClearContents
    PrintSheet.Range("K414").Select
    Selection.ClearContents
    PrintSheet.Range("E416").Select
    Selection.ClearContents
    PrintSheet.Range("D424").Select
    Selection.ClearContents
    PrintSheet.Range("G424").Select
    Selection.ClearContents
    PrintSheet.Range("K424").Select
    Selection.ClearContents
    PrintSheet.Range("N424").Select
    Selection.ClearContents
    PrintSheet.Range("G426").Select
    Selection.ClearContents
    PrintSheet.Range("K426").Select
    Selection.ClearContents
    PrintSheet.Range("G428").Select
    Selection.ClearContents
    PrintSheet.Range("K428").Select
    Selection.ClearContents
    PrintSheet.Range("E430").Select
    Selection.ClearContents
    PrintSheet.Range("D438").Select
    Selection.ClearContents
    PrintSheet.Range("G438").Select
    Selection.ClearContents
    PrintSheet.Range("K438").Select
    Selection.ClearContents
    PrintSheet.Range("N438").Select
    Selection.ClearContents
    PrintSheet.Range("G440").Select
    Selection.ClearContents
    PrintSheet.Range("K440").Select
    Selection.ClearContents
    PrintSheet.Range("G442").Select
    Selection.ClearContents
    PrintSheet.Range("K442").Select
    Selection.ClearContents
    PrintSheet.Range("E444").Select
    Selection.ClearContents
    PrintSheet.Range("D452").Select
    Selection.ClearContents
    PrintSheet.Range("G452").Select
    Selection.ClearContents
    PrintSheet.Range("K452").Select
    Selection.ClearContents
    PrintSheet.Range("N452").Select
    Selection.ClearContents
    PrintSheet.Range("G454").Select
    Selection.ClearContents
    PrintSheet.Range("K454").Select
    Selection.ClearContents
    PrintSheet.Range("G456").Select
    Selection.ClearContents
    PrintSheet.Range("K456").Select
    Selection.ClearContents
    PrintSheet.Range("E458").Select
    Selection.ClearContents
    PrintSheet.Range("D466").Select
    Selection.ClearContents
    PrintSheet.Range("G466").Select
    Selection.ClearContents
    PrintSheet.Range("K466").Select
    Selection.ClearContents
    PrintSheet.Range("N466").Select
    Selection.ClearContents
    PrintSheet.Range("G468").Select
    Selection.ClearContents
    PrintSheet.Range("K468").Select
    Selection.ClearContents
    PrintSheet.Range("G470").Select
    Selection.ClearContents
    PrintSheet.Range("K470").Select
    Selection.ClearContents
    PrintSheet.Range("E472").Select
    Selection.ClearContents


    'Condition History
    PrintSheet.Range("S200").Select
    Selection.ClearContents
    PrintSheet.Range("W200").Select
    Selection.ClearContents
    PrintSheet.Range("AA200").Select
    Selection.ClearContents
    PrintSheet.Range("N200").Select
    Selection.ClearContents
    PrintSheet.Range("W202").Select
    Selection.ClearContents
    PrintSheet.Range("AA202").Select
    Selection.ClearContents
    PrintSheet.Range("W204").Select
    Selection.ClearContents
    PrintSheet.Range("AA204").Select
    Selection.ClearContents
    PrintSheet.Range("U206").Select
    Selection.ClearContents
    PrintSheet.Range("S214").Select
    Selection.ClearContents
    PrintSheet.Range("W214").Select
    Selection.ClearContents
    PrintSheet.Range("AA214").Select
    Selection.ClearContents
    PrintSheet.Range("N214").Select
    Selection.ClearContents
    PrintSheet.Range("W216").Select
    Selection.ClearContents
    PrintSheet.Range("AA216").Select
    Selection.ClearContents
    PrintSheet.Range("W218").Select
    Selection.ClearContents
    PrintSheet.Range("AA218").Select
    Selection.ClearContents
    PrintSheet.Range("U220").Select
    Selection.ClearContents
    PrintSheet.Range("S228").Select
    Selection.ClearContents
    PrintSheet.Range("W228").Select
    Selection.ClearContents
    PrintSheet.Range("AA228").Select
    Selection.ClearContents
    PrintSheet.Range("N228").Select
    Selection.ClearContents
    PrintSheet.Range("W230").Select
    Selection.ClearContents
    PrintSheet.Range("AA230").Select
    Selection.ClearContents
    PrintSheet.Range("W232").Select
    Selection.ClearContents
    PrintSheet.Range("AA232").Select
    Selection.ClearContents
    PrintSheet.Range("U234").Select
    Selection.ClearContents
    PrintSheet.Range("S242").Select
    Selection.ClearContents
    PrintSheet.Range("W242").Select
    Selection.ClearContents
    PrintSheet.Range("AA242").Select
    Selection.ClearContents
    PrintSheet.Range("N242").Select
    Selection.ClearContents
    PrintSheet.Range("W244").Select
    Selection.ClearContents
    PrintSheet.Range("AA244").Select
    Selection.ClearContents
    PrintSheet.Range("W246").Select
    Selection.ClearContents
    PrintSheet.Range("AA246").Select
    Selection.ClearContents
    PrintSheet.Range("U248").Select
    Selection.ClearContents
    PrintSheet.Range("S256").Select
    Selection.ClearContents
    PrintSheet.Range("W256").Select
    Selection.ClearContents
    PrintSheet.Range("AA256").Select
    Selection.ClearContents
    PrintSheet.Range("N256").Select
    Selection.ClearContents
    PrintSheet.Range("W258").Select
    Selection.ClearContents
    PrintSheet.Range("AA258").Select
    Selection.ClearContents
    PrintSheet.Range("W260").Select
    Selection.ClearContents
    PrintSheet.Range("AA260").Select
    Selection.ClearContents
    PrintSheet.Range("U262").Select
    Selection.ClearContents
    PrintSheet.Range("S270").Select
    Selection.ClearContents
    PrintSheet.Range("W270").Select
    Selection.ClearContents
    PrintSheet.Range("AA270").Select
    Selection.ClearContents
    PrintSheet.Range("N270").Select
    Selection.ClearContents
    PrintSheet.Range("W272").Select
    Selection.ClearContents
    PrintSheet.Range("AA272").Select
    Selection.ClearContents
    PrintSheet.Range("W274").Select
    Selection.ClearContents
    PrintSheet.Range("AA274").Select
    Selection.ClearContents
    PrintSheet.Range("U276").Select
    Selection.ClearContents
    PrintSheet.Range("S284").Select
    Selection.ClearContents
    PrintSheet.Range("W284").Select
    Selection.ClearContents
    PrintSheet.Range("AA284").Select
    Selection.ClearContents
    PrintSheet.Range("N284").Select
    Selection.ClearContents
    PrintSheet.Range("W286").Select
    Selection.ClearContents
    PrintSheet.Range("AA286").Select
    Selection.ClearContents
    PrintSheet.Range("W288").Select
    Selection.ClearContents
    PrintSheet.Range("AA288").Select
    Selection.ClearContents
    PrintSheet.Range("U290").Select
    Selection.ClearContents
    PrintSheet.Range("S298").Select
    Selection.ClearContents
    PrintSheet.Range("W298").Select
    Selection.ClearContents
    PrintSheet.Range("AA298").Select
    Selection.ClearContents
    PrintSheet.Range("N298").Select
    Selection.ClearContents
    PrintSheet.Range("W300").Select
    Selection.ClearContents
    PrintSheet.Range("AA300").Select
    Selection.ClearContents
    PrintSheet.Range("W302").Select
    Selection.ClearContents
    PrintSheet.Range("AA302").Select
    Selection.ClearContents
    PrintSheet.Range("U304").Select
    Selection.ClearContents
    PrintSheet.Range("S312").Select
    Selection.ClearContents
    PrintSheet.Range("W312").Select
    Selection.ClearContents
    PrintSheet.Range("AA312").Select
    Selection.ClearContents
    PrintSheet.Range("N312").Select
    Selection.ClearContents
    PrintSheet.Range("W314").Select
    Selection.ClearContents
    PrintSheet.Range("AA314").Select
    Selection.ClearContents
    PrintSheet.Range("W316").Select
    Selection.ClearContents
    PrintSheet.Range("AA316").Select
    Selection.ClearContents
    PrintSheet.Range("U318").Select
    Selection.ClearContents
    PrintSheet.Range("S326").Select
    Selection.ClearContents
    PrintSheet.Range("W326").Select
    Selection.ClearContents
    PrintSheet.Range("AA326").Select
    Selection.ClearContents
    PrintSheet.Range("N326").Select
    Selection.ClearContents
    PrintSheet.Range("W328").Select
    Selection.ClearContents
    PrintSheet.Range("AA328").Select
    Selection.ClearContents
    PrintSheet.Range("W330").Select
    Selection.ClearContents
    PrintSheet.Range("AA330").Select
    Selection.ClearContents
    PrintSheet.Range("U332").Select
    Selection.ClearContents
    PrintSheet.Range("S340").Select
    Selection.ClearContents
    PrintSheet.Range("W340").Select
    Selection.ClearContents
    PrintSheet.Range("AA340").Select
    Selection.ClearContents
    PrintSheet.Range("N340").Select
    Selection.ClearContents
    PrintSheet.Range("W342").Select
    Selection.ClearContents
    PrintSheet.Range("AA342").Select
    Selection.ClearContents
    PrintSheet.Range("W344").Select
    Selection.ClearContents
    PrintSheet.Range("AA344").Select
    Selection.ClearContents
    PrintSheet.Range("U346").Select
    Selection.ClearContents
    PrintSheet.Range("S354").Select
    Selection.ClearContents
    PrintSheet.Range("W354").Select
    Selection.ClearContents
    PrintSheet.Range("AA354").Select
    Selection.ClearContents
    PrintSheet.Range("N354").Select
    Selection.ClearContents
    PrintSheet.Range("W356").Select
    Selection.ClearContents
    PrintSheet.Range("AA356").Select
    Selection.ClearContents
    PrintSheet.Range("W358").Select
    Selection.ClearContents
    PrintSheet.Range("AA358").Select
    Selection.ClearContents
    PrintSheet.Range("U360").Select
    Selection.ClearContents
    PrintSheet.Range("S368").Select
    Selection.ClearContents
    PrintSheet.Range("W368").Select
    Selection.ClearContents
    PrintSheet.Range("AA368").Select
    Selection.ClearContents
    PrintSheet.Range("N368").Select
    Selection.ClearContents
    PrintSheet.Range("W370").Select
    Selection.ClearContents
    PrintSheet.Range("AA370").Select
    Selection.ClearContents
    PrintSheet.Range("W372").Select
    Selection.ClearContents
    PrintSheet.Range("AA372").Select
    Selection.ClearContents
    PrintSheet.Range("U374").Select
    Selection.ClearContents
    PrintSheet.Range("S382").Select
    Selection.ClearContents
    PrintSheet.Range("W382").Select
    Selection.ClearContents
    PrintSheet.Range("AA382").Select
    Selection.ClearContents
    PrintSheet.Range("N382").Select
    Selection.ClearContents
    PrintSheet.Range("W384").Select
    Selection.ClearContents
    PrintSheet.Range("AA384").Select
    Selection.ClearContents
    PrintSheet.Range("W386").Select
    Selection.ClearContents
    PrintSheet.Range("AA386").Select
    Selection.ClearContents
    PrintSheet.Range("U388").Select
    Selection.ClearContents
    PrintSheet.Range("S396").Select
    Selection.ClearContents
    PrintSheet.Range("W396").Select
    Selection.ClearContents
    PrintSheet.Range("AA396").Select
    Selection.ClearContents
    PrintSheet.Range("N396").Select
    Selection.ClearContents
    PrintSheet.Range("W398").Select
    Selection.ClearContents
    PrintSheet.Range("AA398").Select
    Selection.ClearContents
    PrintSheet.Range("W400").Select
    Selection.ClearContents
    PrintSheet.Range("AA400").Select
    Selection.ClearContents
    PrintSheet.Range("U402").Select
    Selection.ClearContents
    PrintSheet.Range("S410").Select
    Selection.ClearContents
    PrintSheet.Range("W410").Select
    Selection.ClearContents
    PrintSheet.Range("AA410").Select
    Selection.ClearContents
    PrintSheet.Range("N410").Select
    Selection.ClearContents
    PrintSheet.Range("W412").Select
    Selection.ClearContents
    PrintSheet.Range("AA412").Select
    Selection.ClearContents
    PrintSheet.Range("W414").Select
    Selection.ClearContents
    PrintSheet.Range("AA414").Select
    Selection.ClearContents
    PrintSheet.Range("U416").Select
    Selection.ClearContents
    PrintSheet.Range("S424").Select
    Selection.ClearContents
    PrintSheet.Range("W424").Select
    Selection.ClearContents
    PrintSheet.Range("AA424").Select
    Selection.ClearContents
    PrintSheet.Range("N424").Select
    Selection.ClearContents
    PrintSheet.Range("W426").Select
    Selection.ClearContents
    PrintSheet.Range("AA426").Select
    Selection.ClearContents
    PrintSheet.Range("W428").Select
    Selection.ClearContents
    PrintSheet.Range("AA428").Select
    Selection.ClearContents
    PrintSheet.Range("U430").Select
    Selection.ClearContents
    PrintSheet.Range("S438").Select
    Selection.ClearContents
    PrintSheet.Range("W438").Select
    Selection.ClearContents
    PrintSheet.Range("AA438").Select
    Selection.ClearContents
    PrintSheet.Range("N438").Select
    Selection.ClearContents
    PrintSheet.Range("W440").Select
    Selection.ClearContents
    PrintSheet.Range("AA440").Select
    Selection.ClearContents
    PrintSheet.Range("W442").Select
    Selection.ClearContents
    PrintSheet.Range("AA442").Select
    Selection.ClearContents
    PrintSheet.Range("U444").Select
    Selection.ClearContents
    PrintSheet.Range("S452").Select
    Selection.ClearContents
    PrintSheet.Range("W452").Select
    Selection.ClearContents
    PrintSheet.Range("AA452").Select
    Selection.ClearContents
    PrintSheet.Range("N452").Select
    Selection.ClearContents
    PrintSheet.Range("W454").Select
    Selection.ClearContents
    PrintSheet.Range("AA454").Select
    Selection.ClearContents
    PrintSheet.Range("W456").Select
    Selection.ClearContents
    PrintSheet.Range("AA456").Select
    Selection.ClearContents
    PrintSheet.Range("U458").Select
    Selection.ClearContents
    PrintSheet.Range("S466").Select
    Selection.ClearContents
    PrintSheet.Range("W466").Select
    Selection.ClearContents
    PrintSheet.Range("AA466").Select
    Selection.ClearContents
    PrintSheet.Range("N466").Select
    Selection.ClearContents
    PrintSheet.Range("W468").Select
    Selection.ClearContents
    PrintSheet.Range("AA468").Select
    Selection.ClearContents
    PrintSheet.Range("W470").Select
    Selection.ClearContents
    PrintSheet.Range("AA470").Select
    Selection.ClearContents
    PrintSheet.Range("U472").Select
    Selection.ClearContents



    'Court Listings History
    PrintSheet.Range("K487").Select
    Selection.ClearContents
    PrintSheet.Range("K489").Select
    Selection.ClearContents
    PrintSheet.Range("K491").Select
    Selection.ClearContents
    PrintSheet.Range("P487").Select
    Selection.ClearContents
    PrintSheet.Range("K495").Select
    Selection.ClearContents
    PrintSheet.Range("K497").Select
    Selection.ClearContents
    PrintSheet.Range("K499").Select
    Selection.ClearContents
    PrintSheet.Range("P495").Select
    Selection.ClearContents
    PrintSheet.Range("K503").Select
    Selection.ClearContents
    PrintSheet.Range("K505").Select
    Selection.ClearContents
    PrintSheet.Range("K507").Select
    Selection.ClearContents
    PrintSheet.Range("P503").Select
    Selection.ClearContents
    PrintSheet.Range("K511").Select
    Selection.ClearContents
    PrintSheet.Range("K513").Select
    Selection.ClearContents
    PrintSheet.Range("K515").Select
    Selection.ClearContents
    PrintSheet.Range("P511").Select
    Selection.ClearContents
    PrintSheet.Range("K519").Select
    Selection.ClearContents
    PrintSheet.Range("K521").Select
    Selection.ClearContents
    PrintSheet.Range("K523").Select
    Selection.ClearContents
    PrintSheet.Range("P519").Select
    Selection.ClearContents
    PrintSheet.Range("K527").Select
    Selection.ClearContents
    PrintSheet.Range("K529").Select
    Selection.ClearContents
    PrintSheet.Range("K531").Select
    Selection.ClearContents
    PrintSheet.Range("P527").Select
    Selection.ClearContents
    PrintSheet.Range("K535").Select
    Selection.ClearContents
    PrintSheet.Range("K537").Select
    Selection.ClearContents
    PrintSheet.Range("K539").Select
    Selection.ClearContents
    PrintSheet.Range("P535").Select
    Selection.ClearContents
    PrintSheet.Range("K543").Select
    Selection.ClearContents
    PrintSheet.Range("K545").Select
    Selection.ClearContents
    PrintSheet.Range("K547").Select
    Selection.ClearContents
    PrintSheet.Range("P543").Select
    Selection.ClearContents
    PrintSheet.Range("K551").Select
    Selection.ClearContents
    PrintSheet.Range("K553").Select
    Selection.ClearContents
    PrintSheet.Range("K555").Select
    Selection.ClearContents
    PrintSheet.Range("P551").Select
    Selection.ClearContents
    PrintSheet.Range("K559").Select
    Selection.ClearContents
    PrintSheet.Range("K561").Select
    Selection.ClearContents
    PrintSheet.Range("K563").Select
    Selection.ClearContents
    PrintSheet.Range("P559").Select
    Selection.ClearContents
    PrintSheet.Range("K567").Select
    Selection.ClearContents
    PrintSheet.Range("K569").Select
    Selection.ClearContents
    PrintSheet.Range("K571").Select
    Selection.ClearContents
    PrintSheet.Range("P567").Select
    Selection.ClearContents
    PrintSheet.Range("K575").Select
    Selection.ClearContents
    PrintSheet.Range("K577").Select
    Selection.ClearContents
    PrintSheet.Range("K579").Select
    Selection.ClearContents
    PrintSheet.Range("P575").Select
    Selection.ClearContents
    PrintSheet.Range("K583").Select
    Selection.ClearContents
    PrintSheet.Range("K585").Select
    Selection.ClearContents
    PrintSheet.Range("K587").Select
    Selection.ClearContents
    PrintSheet.Range("P583").Select
    Selection.ClearContents
    PrintSheet.Range("K591").Select
    Selection.ClearContents
    PrintSheet.Range("K593").Select
    Selection.ClearContents
    PrintSheet.Range("K595").Select
    Selection.ClearContents
    PrintSheet.Range("P591").Select
    Selection.ClearContents
    PrintSheet.Range("K599").Select
    Selection.ClearContents
    PrintSheet.Range("K601").Select
    Selection.ClearContents
    PrintSheet.Range("K603").Select
    Selection.ClearContents
    PrintSheet.Range("P599").Select
    Selection.ClearContents
    PrintSheet.Range("K607").Select
    Selection.ClearContents
    PrintSheet.Range("K609").Select
    Selection.ClearContents
    PrintSheet.Range("K611").Select
    Selection.ClearContents
    PrintSheet.Range("P607").Select
    Selection.ClearContents
    PrintSheet.Range("K615").Select
    Selection.ClearContents
    PrintSheet.Range("K617").Select
    Selection.ClearContents
    PrintSheet.Range("K619").Select
    Selection.ClearContents
    PrintSheet.Range("P615").Select
    Selection.ClearContents
    PrintSheet.Range("K623").Select
    Selection.ClearContents
    PrintSheet.Range("K625").Select
    Selection.ClearContents
    PrintSheet.Range("K627").Select
    Selection.ClearContents
    PrintSheet.Range("P623").Select
    Selection.ClearContents
    PrintSheet.Range("K631").Select
    Selection.ClearContents
    PrintSheet.Range("K633").Select
    Selection.ClearContents
    PrintSheet.Range("K635").Select
    Selection.ClearContents
    PrintSheet.Range("P631").Select
    Selection.ClearContents
    PrintSheet.Range("K639").Select
    Selection.ClearContents
    PrintSheet.Range("K641").Select
    Selection.ClearContents
    PrintSheet.Range("K643").Select
    Selection.ClearContents
    PrintSheet.Range("P639").Select
    Selection.ClearContents
    PrintSheet.Range("K647").Select
    Selection.ClearContents
    PrintSheet.Range("K649").Select
    Selection.ClearContents
    PrintSheet.Range("K651").Select
    Selection.ClearContents
    PrintSheet.Range("P647").Select
    Selection.ClearContents
    PrintSheet.Range("K655").Select
    Selection.ClearContents
    PrintSheet.Range("K657").Select
    Selection.ClearContents
    PrintSheet.Range("K659").Select
    Selection.ClearContents
    PrintSheet.Range("P655").Select
    Selection.ClearContents
    PrintSheet.Range("K663").Select
    Selection.ClearContents
    PrintSheet.Range("K665").Select
    Selection.ClearContents
    PrintSheet.Range("K667").Select
    Selection.ClearContents
    PrintSheet.Range("P663").Select
    Selection.ClearContents
    PrintSheet.Range("K671").Select
    Selection.ClearContents
    PrintSheet.Range("K673").Select
    Selection.ClearContents
    PrintSheet.Range("K675").Select
    Selection.ClearContents
    PrintSheet.Range("P671").Select
    Selection.ClearContents
    PrintSheet.Range("K679").Select
    Selection.ClearContents
    PrintSheet.Range("K681").Select
    Selection.ClearContents
    PrintSheet.Range("K683").Select
    Selection.ClearContents
    PrintSheet.Range("P679").Select
    Selection.ClearContents
    PrintSheet.Range("K687").Select
    Selection.ClearContents
    PrintSheet.Range("K689").Select
    Selection.ClearContents
    PrintSheet.Range("K691").Select
    Selection.ClearContents
    PrintSheet.Range("P687").Select
    Selection.ClearContents
    PrintSheet.Range("K695").Select
    Selection.ClearContents
    PrintSheet.Range("K697").Select
    Selection.ClearContents
    PrintSheet.Range("K699").Select
    Selection.ClearContents
    PrintSheet.Range("P695").Select
    Selection.ClearContents
    PrintSheet.Range("K703").Select
    Selection.ClearContents
    PrintSheet.Range("K705").Select
    Selection.ClearContents
    PrintSheet.Range("K707").Select
    Selection.ClearContents
    PrintSheet.Range("P703").Select
    Selection.ClearContents
    PrintSheet.Range("K711").Select
    Selection.ClearContents
    PrintSheet.Range("K713").Select
    Selection.ClearContents
    PrintSheet.Range("K715").Select
    Selection.ClearContents
    PrintSheet.Range("P711").Select
    Selection.ClearContents
    PrintSheet.Range("K719").Select
    Selection.ClearContents
    PrintSheet.Range("K721").Select
    Selection.ClearContents
    PrintSheet.Range("K723").Select
    Selection.ClearContents
    PrintSheet.Range("P719").Select
    Selection.ClearContents
    PrintSheet.Range("K727").Select
    Selection.ClearContents
    PrintSheet.Range("K729").Select
    Selection.ClearContents
    PrintSheet.Range("K731").Select
    Selection.ClearContents
    PrintSheet.Range("P727").Select
    Selection.ClearContents
    PrintSheet.Range("K735").Select
    Selection.ClearContents
    PrintSheet.Range("K737").Select
    Selection.ClearContents
    PrintSheet.Range("K739").Select
    Selection.ClearContents
    PrintSheet.Range("P735").Select
    Selection.ClearContents
    PrintSheet.Range("K743").Select
    Selection.ClearContents
    PrintSheet.Range("K745").Select
    Selection.ClearContents
    PrintSheet.Range("K747").Select
    Selection.ClearContents
    PrintSheet.Range("P743").Select
    Selection.ClearContents
    PrintSheet.Range("K751").Select
    Selection.ClearContents
    PrintSheet.Range("K753").Select
    Selection.ClearContents
    PrintSheet.Range("K755").Select
    Selection.ClearContents
    PrintSheet.Range("P751").Select
    Selection.ClearContents
    PrintSheet.Range("K759").Select
    Selection.ClearContents
    PrintSheet.Range("K761").Select
    Selection.ClearContents
    PrintSheet.Range("K763").Select
    Selection.ClearContents
    PrintSheet.Range("P759").Select
    Selection.ClearContents
    PrintSheet.Range("K767").Select
    Selection.ClearContents
    PrintSheet.Range("K769").Select
    Selection.ClearContents
    PrintSheet.Range("K771").Select
    Selection.ClearContents
    PrintSheet.Range("P767").Select
    Selection.ClearContents
    PrintSheet.Range("K775").Select
    Selection.ClearContents
    PrintSheet.Range("K777").Select
    Selection.ClearContents
    PrintSheet.Range("K779").Select
    Selection.ClearContents
    PrintSheet.Range("P775").Select
    Selection.ClearContents
    PrintSheet.Range("K783").Select
    Selection.ClearContents
    PrintSheet.Range("K785").Select
    Selection.ClearContents
    PrintSheet.Range("K787").Select
    Selection.ClearContents
    PrintSheet.Range("P783").Select
    Selection.ClearContents
    PrintSheet.Range("K791").Select
    Selection.ClearContents
    PrintSheet.Range("K793").Select
    Selection.ClearContents
    PrintSheet.Range("K795").Select
    Selection.ClearContents
    PrintSheet.Range("P791").Select
    Selection.ClearContents
    PrintSheet.Range("K799").Select
    Selection.ClearContents
    PrintSheet.Range("K801").Select
    Selection.ClearContents
    PrintSheet.Range("K803").Select
    Selection.ClearContents
    PrintSheet.Range("P799").Select
    Selection.ClearContents




    PrintSheet.Range("K5").Select

    Call YouthSearchPrint



End Sub

Sub YouthSearchPrint()
    Call RefreshNamedRanges
    Call Generate_Dictionaries

    Dim PrintSheet As Worksheet
    Dim DataSheet As Worksheet
    Set PrintSheet = Worksheets("Youth Search")
    Set DataSheet = Worksheets("Entry")

    Dim userRow As Long
    userRow = PrintSheet.Range("J5").value

    'PRINT IDENTIFIERS BOX

    'example stringing together two different cell values to print
    PrintSheet.Range("D12").value = DataSheet.Range(hFind("First Name") & userRow).value & " " & DataSheet.Range(hFind("Last Name") & userRow).value

    'basic examples
    PrintSheet.Range("D14").value = DataSheet.Range(hFind("DOB") & userRow).value
    PrintSheet.Range("I12").value = DataSheet.Range(hFind("PID #") & userRow).value
    PrintSheet.Range("I14").value = DataSheet.Range(hFind("SID #") & userRow).value

    'examples with Lookup dictionary
    PrintSheet.Range("D19").value = Lookup("Sex_Num")(DataSheet.Range(hFind("Sex") & userRow).value)
    PrintSheet.Range("D21").value = Lookup("Race_Num")(DataSheet.Range(hFind("Race") & userRow).value)

    'basic example
    PrintSheet.Range("D23").value = DataSheet.Range(hFind("Address") & userRow).value
    PrintSheet.Range("D26").value = DataSheet.Range(hFind("Zipcode", "DEMOGRAPHICS") & userRow).value

    'example using a custom function
    PrintSheet.Range("I19").value = ageAtTime(VBA.Format(Now(), "mm/dd/yyyy"), userRow)

    PrintSheet.Range("I21").value = DataSheet.Range(hFind("Age @ Intake") & userRow).value
    PrintSheet.Range("I23").value = DataSheet.Range(hFind("Guardian First") & userRow).value & " " & DataSheet.Range(hFind("Guardian Last") & userRow).value

    PrintSheet.Range("E28").value = DataSheet.Range(hFind("School") & userRow).value
    PrintSheet.Range("E30").value = DataSheet.Range(hFind("Grade") & userRow).value
    PrintSheet.Range("I26").value = DataSheet.Range(hFind("Phone #") & userRow).value


    'PRINT ARREST & PETITION INFO BOX
    'Arrest info
    PrintSheet.Range("O12").value = DataSheet.Range(hFind("DC #") & userRow).value
    PrintSheet.Range("P14").value = DataSheet.Range(hFind("Incident District") & userRow).value
    PrintSheet.Range("P16").value = DataSheet.Range(hFind("Arrest Date") & userRow).value
    PrintSheet.Range("P18").value = Lookup("Generic_YNOU_Num")(DataSheet.Range(hFind("Active in System at Time of Arrest?", "Petition") & userRow).value)
    PrintSheet.Range("P20").value = DataSheet.Range(hFind("# of Prior Arrests") & userRow).value

    'Notes from Intake
    PrintSheet.Range("N23").value = DataSheet.Range(hFind("General Notes from Intake") & userRow).value

    'Petition #1
    PrintSheet.Range("T12").value = DataSheet.Range(hFind("Petition #1") & userRow).value
    PrintSheet.Range("Y12").value = DataSheet.Range(hFind("Lead Charge Name", "Petition #1") & userRow).value
    PrintSheet.Range("AC12").value = Lookup("Charge_Grade_Specific_Num")(DataSheet.Range(hFind("Charge Grade (specific) #1", "Petition #1") & userRow).value)
    PrintSheet.Range("Y14").value = DataSheet.Range(hFind("Charge Name #2", "Petition #1") & userRow).value
    PrintSheet.Range("AC14").value = Lookup("Charge_Grade_Specific_Num")(DataSheet.Range(hFind("Charge Grade (specific) #2", "Petition #1") & userRow).value)
    PrintSheet.Range("Y16").value = DataSheet.Range(hFind("Charge Name #3", "Petition #1") & userRow).value
    PrintSheet.Range("AC16").value = Lookup("Charge_Grade_Specific_Num")(DataSheet.Range(hFind("Charge Grade (specific) #3", "Petition #1") & userRow).value)
    PrintSheet.Range("Y18").value = DataSheet.Range(hFind("Charge Name #4", "Petition #1") & userRow).value
    PrintSheet.Range("AC18").value = Lookup("Charge_Grade_Specific_Num")(DataSheet.Range(hFind("Charge Grade (specific) #4", "Petition #1") & userRow).value)
    PrintSheet.Range("Y20").value = DataSheet.Range(hFind("Charge Name #5", "Petition #1") & userRow).value
    PrintSheet.Range("AC20").value = Lookup("Charge_Grade_Specific_Num")(DataSheet.Range(hFind("Charge Grade (specific) #5", "Petition #1") & userRow).value)

    PrintSheet.Range("T14").value = DataSheet.Range(hFind("Date Filed", "Petition #1") & userRow).value


    'Petition #2
    PrintSheet.Range("T23").value = DataSheet.Range(hFind("Petition #2") & userRow).value
    PrintSheet.Range("Y23").value = DataSheet.Range(hFind("Lead Charge Name", "Petition #2") & userRow).value
    PrintSheet.Range("AC23").value = Lookup("Charge_Grade_Specific_Num")(DataSheet.Range(hFind("Charge Grade (specific) #1", "Petition #2") & userRow).value)
    PrintSheet.Range("Y25").value = DataSheet.Range(hFind("Charge Name #2", "Petition #2") & userRow).value
    PrintSheet.Range("AC25").value = Lookup("Charge_Grade_Specific_Num")(DataSheet.Range(hFind("Charge Grade (specific) #2", "Petition #2") & userRow).value)
    PrintSheet.Range("Y27").value = DataSheet.Range(hFind("Charge Name #3", "Petition #2") & userRow).value
    PrintSheet.Range("AC27").value = Lookup("Charge_Grade_Specific_Num")(DataSheet.Range(hFind("Charge Grade (specific) #3", "Petition #2") & userRow).value)
    PrintSheet.Range("Y29").value = DataSheet.Range(hFind("Charge Name #4", "Petition #2") & userRow).value
    PrintSheet.Range("AC29").value = Lookup("Charge_Grade_Specific_Num")(DataSheet.Range(hFind("Charge Grade (specific) #4", "Petition #2") & userRow).value)
    PrintSheet.Range("Y31").value = DataSheet.Range(hFind("Charge Name #5", "Petition #2") & userRow).value
    PrintSheet.Range("AC31").value = Lookup("Charge_Grade_Specific_Num")(DataSheet.Range(hFind("Charge Grade (specific) #5", "Petition #2") & userRow).value)

    PrintSheet.Range("T25").value = DataSheet.Range(hFind("Date Filed", "Petition #2") & userRow).value


    'Petition #3
    PrintSheet.Range("T34").value = DataSheet.Range(hFind("Petition #3") & userRow).value
    PrintSheet.Range("Y34").value = DataSheet.Range(hFind("Lead Charge Name", "Petition #3") & userRow).value
    PrintSheet.Range("AC34").value = Lookup("Charge_Grade_Specific_Num")(DataSheet.Range(hFind("Charge Grade (specific) #1", "Petition #3") & userRow).value)
    PrintSheet.Range("Y36").value = DataSheet.Range(hFind("Charge Name #2", "Petition #3") & userRow).value
    PrintSheet.Range("AC36").value = Lookup("Charge_Grade_Specific_Num")(DataSheet.Range(hFind("Charge Grade (specific) #2", "Petition #3") & userRow).value)
    PrintSheet.Range("Y38").value = DataSheet.Range(hFind("Charge Name #3", "Petition #3") & userRow).value
    PrintSheet.Range("AC38").value = Lookup("Charge_Grade_Specific_Num")(DataSheet.Range(hFind("Charge Grade (specific) #3", "Petition #3") & userRow).value)
    PrintSheet.Range("Y40").value = DataSheet.Range(hFind("Charge Name #4", "Petition #3") & userRow).value
    PrintSheet.Range("AC40").value = Lookup("Charge_Grade_Specific_Num")(DataSheet.Range(hFind("Charge Grade (specific) #4", "Petition #3") & userRow).value)
    PrintSheet.Range("Y42").value = DataSheet.Range(hFind("Charge Name #5", "Petition #3") & userRow).value
    PrintSheet.Range("AC42").value = Lookup("Charge_Grade_Specific_Num")(DataSheet.Range(hFind("Charge Grade (specific) #5", "Petition #3") & userRow).value)

    PrintSheet.Range("T36").value = DataSheet.Range(hFind("Date Filed", "Petition #2") & userRow).value


    'Petition #4
    PrintSheet.Range("T45").value = DataSheet.Range(hFind("Petition #4") & userRow).value
    PrintSheet.Range("Y45").value = DataSheet.Range(hFind("Lead Charge Name", "Petition #4") & userRow).value
    PrintSheet.Range("AC45").value = Lookup("Charge_Grade_Specific_Num")(DataSheet.Range(hFind("Charge Grade (specific) #1", "Petition #4") & userRow).value)
    PrintSheet.Range("Y47").value = DataSheet.Range(hFind("Charge Name #2", "Petition #4") & userRow).value
    PrintSheet.Range("AC47").value = Lookup("Charge_Grade_Specific_Num")(DataSheet.Range(hFind("Charge Grade (specific) #2", "Petition #4") & userRow).value)
    PrintSheet.Range("Y49").value = DataSheet.Range(hFind("Charge Name #3", "Petition #4") & userRow).value
    PrintSheet.Range("AC49").value = Lookup("Charge_Grade_Specific_Num")(DataSheet.Range(hFind("Charge Grade (specific) #3", "Petition #4") & userRow).value)
    PrintSheet.Range("Y51").value = DataSheet.Range(hFind("Charge Name #4", "Petition #4") & userRow).value
    PrintSheet.Range("AC51").value = Lookup("Charge_Grade_Specific_Num")(DataSheet.Range(hFind("Charge Grade (specific) #4", "Petition #4") & userRow).value)
    PrintSheet.Range("Y53").value = DataSheet.Range(hFind("Charge Name #5", "Petition #4") & userRow).value
    PrintSheet.Range("AC53").value = Lookup("Charge_Grade_Specific_Num")(DataSheet.Range(hFind("Charge Grade (specific) #5", "Petition #4") & userRow).value)

    PrintSheet.Range("T47").value = DataSheet.Range(hFind("Date Filed", "Petition #2") & userRow).value


    'STATUS OF ARREST INCIDENT
    Dim activeStatus As String
    activeStatus = Lookup("Active_Num")(DataSheet.Range(hFind("Active or Discharged (in courtroom)?") & userRow).value)
    PrintSheet.Range("D37").value = activeStatus
    PrintSheet.Range("G37").value = Lookup("Generic_YNOU_Num")(DataSheet.Range(hFind("Did Youth Enter an Admission?", "COURT PROCEEDINGS", "AGGREGATES") & userRow).value)
    PrintSheet.Range("J37").value = Lookup("Generic_YNOU_Num")(DataSheet.Range(hFind("Adjudicated Delinquent?", "COURT PROCEEDINGS", "AGGREGATES") & userRow).value)
    PrintSheet.Range("G39").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Admissions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("J39").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Adjudicating Courtroom", "Adjudications", "AGGREGATES") & userRow).value)


    'SHONA & ADAM'S LOS CALCULATIONS

    If StrComp(activeStatus, "Active") = 0 Then
        'LoS for arrest and petition
        Dim losArrest As Integer
        Dim losPetition As Integer

        Dim ArrestDate As String
        ArrestDate = DataSheet.Range(headerFind("Arrest Date (current petition)") & userRow).value
        losArrest = DateDiff("d", ArrestDate, VBA.Format(Now(), "mm/dd/yyyy"))
        PrintSheet.Range("D39").value = losArrest & " days"

        Dim petitionDate As String
        petitionDate = DataSheet.Range(hFind("Date Filed", "Petition") & userRow).value
        losPetition = DateDiff("d", petitionDate, VBA.Format(Now(), "mm/dd/yyyy"))
        PrintSheet.Range("D41").value = losPetition & " days"


        'ACTIVE COURT PROCEEDINGS

        'Print listing type from lookup values

        PrintSheet.Range("D50").value = Lookup("Listing_Type_Num")(DataSheet.Range(hFind("Listing Type", "DEMOGRAPHICS") & userRow).value)

        Dim Courtroom As String
        Dim losCourtroom As Integer
        Dim courtroomOptions(1 To 6) As String
        courtroomOptions(1) = "4G"
        courtroomOptions(2) = "4E"
        courtroomOptions(3) = "6F"
        courtroomOptions(4) = "6H"
        courtroomOptions(5) = "3E"
        courtroomOptions(6) = "ADULT"
        Courtroom = findFirstValue(DataSheet, userRow, "4G", courtroomOptions, "Start Date", "End Date")
        PrintSheet.Range("G48") = Courtroom

        'if courtroom exists
        If Not StrComp(Courtroom, "") = 0 Then
            losCourtroom = DateDiff("d", DataSheet.Range(hFind("Start Date", Courtroom, "4G") & userRow).value, _
                VBA.Format(Now(), "mm/dd/yyyy"))
            PrintSheet.Range("J48") = losCourtroom & " days"

        Else


            Dim SpecialtyCourtroom As String
            Dim losSpecialtyCourtroom As Integer
            Dim SpecialtyCourtroomOptions(1 To 3) As String
            SpecialtyCourtroomOptions(1) = "Crossover"
            SpecialtyCourtroomOptions(2) = "WRAP"
            SpecialtyCourtroomOptions(3) = "JTC"
            SpecialtyCourtroom = findFirstValue(DataSheet, userRow, "Crossover", SpecialtyCourtroomOptions, "Referral Date", "Date of Overall Discharge")
            PrintSheet.Range("G48") = SpecialtyCourtroom


            'if courtroom exists
            If Not StrComp(SpecialtyCourtroom, "") = 0 Then
                losSpecialtyCourtroom = DateDiff("d", DataSheet.Range(hFind("Referral Date", SpecialtyCourtroom, "Crossover") & userRow).value, _
                VBA.Format(Now(), "mm/dd/yyyy"))
                PrintSheet.Range("J48") = losSpecialtyCourtroom & " days"


            End If

        End If



        Dim legalStatus As String
        Dim losLegalStatus As Integer
        Dim lostLegalStatus As Integer
        Dim legalStatusOptions(1 To 5) As String
        legalStatusOptions(1) = "Pretrial"
        legalStatusOptions(2) = "Consent Decree"
        legalStatusOptions(3) = "Interim Probation"
        legalStatusOptions(4) = "Probation"
        legalStatusOptions(5) = "Aftercare Probation"
        legalStatus = findFirstValue(DataSheet, userRow, "Aggregates", legalStatusOptions, "Start Date", "End Date")
        PrintSheet.Range("G50") = legalStatus

        'if legal status exists
        If Not StrComp(legalStatus, "") = 0 Then
            losLegalStatus = DateDiff("d", DataSheet.Range(hFind("Start Date", legalStatus, "Aggregates") & userRow).value, _
            VBA.Format(Now(), "mm/dd/yyyy"))
            PrintSheet.Range("J50") = losLegalStatus & " days"

        End If



        Dim SpecialtyLegalStatus As String
        Dim losSpecialtyLegalStatus As Integer
        Dim lostSpecialtyLegalStatus As Integer
        Dim SpecialtyLegalStatusOptions(1 To 3) As String
        SpecialtyLegalStatusOptions(1) = "Crossover"
        SpecialtyLegalStatusOptions(2) = "WRAP"
        SpecialtyLegalStatusOptions(3) = "JTC"

        'Find if youth has "Accepted Date" to courtroom and then return that he/she is on that legal status
        SpecialtyLegalStatus = findFirstValue(DataSheet, userRow, "Crossover", SpecialtyLegalStatusOptions, "Accepted Date", "Date of Overall Discharge")
        PrintSheet.Range("G50") = SpecialtyLegalStatus

        'If youth as been accepted to that courtroom, date-diff acceptance date to show LOS of specialty court legal status
        If Not StrComp(SpecialtyLegalStatus, "") = 0 Then
            losSpecialtyLegalStatus = DateDiff("d", DataSheet.Range(hFind("Accepted Date", SpecialtyLegalStatus) & userRow).value, _
            VBA.Format(Now(), "mm/dd/yyyy"))
            PrintSheet.Range("J50") = losSpecialtyLegalStatus & " days"

        Else

            'If youth is in specialty courtroom and accepted date is still "0" (ie, has not yet been accepted),
            'print legal status as "Pretrial" - LOS Pretrial will get handled by standard LOS sweep)

            PrintSheet.Range("G50") = "Pretrial"


        End If




        'Supervision Programs
        '        Dim supervisionProgramColumns() As String
        '        supervisionProgramColumns = findAllValues(DataSheet, userRow, "Aggregates", "Supervision Ordered", "Start Date", "End Date")
        '        Dim supervisionArrLength As Integer
        '
        '        If (Not supervisionProgramColumns) = -1 Then
        '            supervisionArrLength = 0
        '        Else
        '            supervisionArrLength = UBound(supervisionProgramColumns) - LBound(supervisionProgramColumns) + 1
        '        End If
        '
        '        Dim supervisionI As Integer
        '        Dim supervisionStart As String
        '
        '        For supervisionI = 1 To supervisionArrLength
        '            PrintSheet.Range("D" & 53 + 2 * supervisionI) = Lookup("Supervision_Program_Num")(DataSheet.Range(hFind(supervisionProgramColumns(supervisionI - 1), "Aggregates") & userRow).value)
        '            supervisionStart = DataSheet.Range(hFind("Start Date", supervisionProgramColumns(supervisionI - 1), "Aggregates") & userRow).value
        '            PrintSheet.Range("J" & 53 + 2 * supervisionI) = DateDiff("d", supervisionStart, VBA.Format(Now(), "mm/dd/yyyy")) & " days"
        '            If supervisionI = 3 Then Exit For
        '
        '        Next supervisionI
        '
        '
        '        'Supervision Providers
        '        'Provider
        '        Dim cbsupervisionProviderColumnsI() As String
        '        cbsupervisionProviderColumnsI = findAllValues2(DataSheet, userRow, "Aggregates", "Supervision Ordered", "Community-Based Agency", "Start Date", "End Date")
        '        Dim cbsupervisionProviderArrLengthI As Integer
        '
        '
        '        If (Not cbsupervisionProviderColumnsI) = -1 Then
        '            cbsupervisionProviderArrLengthI = 0
        '        Else
        '            cbsupervisionProviderArrLengthI = UBound(cbsupervisionProviderColumnsI) - LBound(cbsupervisionProviderColumnsI) + 1
        '
        '        End If
        '
        '
        '        Dim cbsupervisionProviderI As Integer
        '        Dim cbsupervisionProviderStart As String
        '
        '        Dim i As Integer
        '        Dim j As Integer
        '            i = (DataSheet.Range(hFind("Community-Based Agency", "Supervision Ordered #1", "Aggregates", "Start Date", "End Date") & userRow).value)
        '            j = (DataSheet.Range(hFind("Residential Agency", "Supervision Ordered #1", "Aggregates", "Start Date", "End Date") & userRow).value)
        '
        '
        '       For cbsupervisionProviderI = 1 To cbsupervisionProviderArrLengthI
        '        'If i > 0 Then
        '            PrintSheet.Range("G55") = Lookup("Community_Based_Supervision_Provider_Num")(DataSheet.Range(hFind(cbsupervisionProviderColumnsI(cbsupervisionProviderI - 1), "Aggregates") & userRow).value)
        '        'End If
        '
        '        'If j > 0 Then
        '            'PrintSheet.Range("G55") = Lookup("Residential_Supervision_Provider_Num")(DataSheet.Range(hFind(cbsupervisionProviderColumnsI(cbsupervisionProviderI - 1), "Aggregates") & userRow).value)
        '        'End If
        '         'End If
        '        If cbsupervisionProviderI = 3 Then Exit For
        '
        '
        '        Next cbsupervisionProviderI
        '

        Dim i As Integer
        Dim bucketHead As String
        Dim printRow As Long
        Dim programType As String
        Dim providerName As String
        Dim thing1 As String, thing2 As String

        printRow = 55

        For i = 1 To 30
            bucketHead = hFind("Supervision Ordered #" & i, "AGGREGATES")

            thing1 = DataSheet.Range(headerFind("Start Date", bucketHead) & userRow).value
            thing2 = DataSheet.Range(headerFind("End Date", bucketHead) & userRow).value

            If isNotEmptyOrZero(DataSheet.Range(headerFind("Start Date", bucketHead) & userRow)) _
              And isEmptyOrZero(DataSheet.Range(headerFind("End Date", bucketHead) & userRow)) Then

                programType = Lookup("Supervision_Program_Num")(DataSheet.Range(bucketHead & userRow).value)

                If isResidential(programType) Then
                    providerName = Lookup("Residential_Supervision_Provider_Num")(DataSheet.Range(headerFind("Residential Agency", bucketHead) & userRow).value)
                Else
                    providerName = Lookup("Community_Based_Supervision_Provider_Num")(DataSheet.Range(headerFind("Community-Based Agency", bucketHead) & userRow).value)
                End If

                If printRow > 59 Then
                    MsgBox "The following Active Supervision will not be printed due to space constraints: " _
                        & vbNewLine & "Program: " & programType _
                        & vbNewLine & "Provider: " & providerName _
                        & vbNewLine & "Start Date: " & DataSheet.Range(headerFind("Start Date", bucketHead) & userRow).value
                End If

                PrintSheet.Range("D" & printRow) = programType
                PrintSheet.Range("G" & printRow) = providerName
                PrintSheet.Range("J" & printRow) = DateDiff("d", DataSheet.Range(headerFind("Start Date", bucketHead) & userRow).value, Date) & " days"

                printRow = printRow + 2
            End If

        Next i


        'El


        'residential
        'Dim rsupervisionProviderColumnsI() As String
        'rsupervisionProviderColumnsI = findAllValues(DataSheet, userRow, "Aggregates", "Supervision Ordered", "Start Date", "End Date")
        'Dim rsupervisionProviderArrLengthI As Integer

        'If (Not rsupervisionProviderColumnsI) = -1 Then
        'rsupervisionProviderArrLengthI = 0
        'Else
        'rsupervisionProviderArrLengthI = UBound(rsupervisionProviderColumnsI) - LBound(rsupervisionProviderColumnsI) + 1
        'End If

        'Dim rsupervisionProviderI As Integer
        'Dim rsupervisionProviderStart As String

        'For rsupervisionProviderI = 1 To rsupervisionProviderArrLengthI
        'PrintSheet.Range("G" & 53 + 2 * rsupervisionProviderI) = Lookup("Residential_Supervision_Provider_Num")(DataSheet.Range(hFind(rsupervisionProviderColumnsI(rsupervisionProviderI - 1), "Aggregates") & userRow).value)
        'PrintSheet.Range("G" & 53 + 2 * supervisionProviderI) = Lookup("Residential_Supervision_Provider_Num")(DataSheet.Range(hFind(supervisionProviderColumns(supervisionProviderI - 1), "Aggregates") & userRow).value)
        'If rsupervisionProviderI = 3 Then Exit For

        'Next cbsupervisionProviderI






        'Conditions
        Dim conditionsColumns() As String
        conditionsColumns = findAllValues(DataSheet, userRow, "Aggregates", "Condition Ordered", "Start Date", "End Date")
        Dim conditionArrLength As Integer

        If (Not conditionsColumns) = -1 Then
            conditionArrLength = 0
        Else
            conditionArrLength = UBound(conditionsColumns) - LBound(conditionsColumns) + 1
        End If
        Dim conditionI As Integer
        Dim conditionStart As String

        For conditionI = 1 To conditionArrLength
            PrintSheet.Range("D" & 62 + 2 * conditionI) = Lookup("Condition_Num")(DataSheet.Range(hFind(conditionsColumns(conditionI - 1), "Aggregates") & userRow).value)
            conditionStart = DataSheet.Range(hFind("Start Date", conditionsColumns(conditionI - 1), "Aggregates") & userRow).value
            PrintSheet.Range("J" & 62 + 2 * conditionI) = DateDiff("d", conditionStart, VBA.Format(Now(), "mm/dd/yyyy")) & " days"
            If conditionI = 6 Then Exit For
        Next conditionI
    End If

    PrintSheet.Range("D43").value = DataSheet.Range(hFind("Total LOS From Arrest", "Petition Outcomes") & userRow).value

    'COURT PROCEEDINGS
    PrintSheet.Range("D48").value = DataSheet.Range(hFind("Next Court Date") & userRow).value
    'Replaced by courtroom & legal status calculation above
    'PrintSheet.Range("G48").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Active Courtroom") & userRow).value)
    'PrintSheet.Range("G50").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status") & userRow).value)


    'SUPERVISION PROGRAMS
    'PrintSheet.Range("D55").value = Lookup("Supervision_Program_Num")(DataSheet.Range(hFind("Active Supervision") & userRow).value)
    'PrintSheet.Range("G55").value = Lookup("Residential_Supervision_Provider_Num")(DataSheet.Range(hFind("Active Supervision Provider") & userRow).value)

    'COURTROOM HISTORY
    'Detention
    PrintSheet.Range("D84").value = Lookup("Generic_YNOU_Num")(DataSheet.Range(hFind("Did Youth Have Initial Detention Hearing?", "DETENTION") & userRow).value)
    PrintSheet.Range("G84").value = DataSheet.Range(hFind("Date of Initial Detention Hearing") & userRow).value
    PrintSheet.Range("J84").value = DataSheet.Range(hFind("Date of Release", "DETENTION") & userRow).value
    PrintSheet.Range("N84").value = DataSheet.Range(hFind("LOS in Detention", "DETENTION") & userRow).value
    PrintSheet.Range("D86").value = DataSheet.Range(hFind("Notes on Detention", "DETENTION") & userRow).value

    '4G
    PrintSheet.Range("D94").value = Lookup("Generic_YNOU_Num")(DataSheet.Range(hFind("Was Youth in 4G?") & userRow).value)
    PrintSheet.Range("G94").value = DataSheet.Range(hFind("Start Date", "4G") & userRow).value
    PrintSheet.Range("J94").value = DataSheet.Range(hFind("End Date", "4G") & userRow).value
    PrintSheet.Range("N94").value = DataSheet.Range(hFind("LOS", "4G") & userRow).value
    PrintSheet.Range("D96").value = DataSheet.Range(hFind("Notes on 4G", "4G") & userRow).value

    '3E
    PrintSheet.Range("D104").value = Lookup("Generic_YNOU_Num")(DataSheet.Range(hFind("Was Youth in 3E?") & userRow).value)
    PrintSheet.Range("G104").value = DataSheet.Range(hFind("Start Date", "3E") & userRow).value
    PrintSheet.Range("J104").value = DataSheet.Range(hFind("End Date", "3E") & userRow).value
    PrintSheet.Range("N104").value = DataSheet.Range(hFind("LOS", "3E") & userRow).value
    PrintSheet.Range("D106").value = DataSheet.Range(hFind("Notes on 3E", "3E") & userRow).value

    'JTC
    PrintSheet.Range("D114").value = Lookup("Generic_YNOU_Num")(DataSheet.Range(hFind("Was Youth in JTC?") & userRow).value)
    PrintSheet.Range("G114").value = DataSheet.Range(hFind("Referral Date", "JTC") & userRow).value
    PrintSheet.Range("J114").value = DataSheet.Range(hFind("Date of Overall Discharge", "JTC") & userRow).value
    PrintSheet.Range("N114").value = DataSheet.Range(hFind("Total LOS in JTC", "JTC") & userRow).value
    PrintSheet.Range("D116").value = DataSheet.Range(hFind("Notes on Outcome", "JTC") & userRow).value

    'CROSSOVER
    PrintSheet.Range("D124").value = Lookup("Generic_YNOU_Num")(DataSheet.Range(hFind("Was Youth on Crossover Status?") & userRow).value)
    PrintSheet.Range("G124").value = DataSheet.Range(hFind("Referral Date", "Crossover") & userRow).value
    PrintSheet.Range("J124").value = DataSheet.Range(hFind("End Date", "Crossover") & userRow).value
    PrintSheet.Range("N124").value = DataSheet.Range(hFind("LOS", "Crossover") & userRow).value
    PrintSheet.Range("D126").value = DataSheet.Range(hFind("Notes on Outcome", "Crossover") & userRow).value

    'WRAP
    PrintSheet.Range("D134").value = Lookup("Generic_YNOU_Num")(DataSheet.Range(hFind("Was Youth on WRAP Status?") & userRow).value)
    PrintSheet.Range("G134").value = DataSheet.Range(hFind("Referral Date", "WRAP") & userRow).value
    PrintSheet.Range("J134").value = DataSheet.Range(hFind("End Date", "WRAP") & userRow).value
    PrintSheet.Range("N134").value = DataSheet.Range(hFind("LOS", "WRAP") & userRow).value
    PrintSheet.Range("D136").value = DataSheet.Range(hFind("Notes on Outcome", "WRAP") & userRow).value

    '4E
    PrintSheet.Range("D144").value = Lookup("Generic_YNOU_Num")(DataSheet.Range(hFind("Was Youth in 4E?") & userRow).value)
    PrintSheet.Range("G144").value = DataSheet.Range(hFind("Start Date", "4E") & userRow).value
    PrintSheet.Range("J144").value = DataSheet.Range(hFind("End Date", "4E") & userRow).value
    PrintSheet.Range("N144").value = DataSheet.Range(hFind("LOS", "4E") & userRow).value
    PrintSheet.Range("D146").value = DataSheet.Range(hFind("Notes on 4E", "4E") & userRow).value

    '5F


    '6F
    PrintSheet.Range("D164").value = Lookup("Generic_YNOU_Num")(DataSheet.Range(hFind("Was Youth in 6F?") & userRow).value)
    PrintSheet.Range("G164").value = DataSheet.Range(hFind("Start Date", "6F") & userRow).value
    PrintSheet.Range("J164").value = DataSheet.Range(hFind("End Date", "6F") & userRow).value
    PrintSheet.Range("N164").value = DataSheet.Range(hFind("LOS", "6F") & userRow).value
    PrintSheet.Range("D166").value = DataSheet.Range(hFind("Notes on 6F", "6F") & userRow).value

    '6H
    PrintSheet.Range("D174").value = Lookup("Generic_YNOU_Num")(DataSheet.Range(hFind("Was Youth in 6H?") & userRow).value)
    PrintSheet.Range("G174").value = DataSheet.Range(hFind("Start Date", "6H") & userRow).value
    PrintSheet.Range("J174").value = DataSheet.Range(hFind("End Date", "6H") & userRow).value
    PrintSheet.Range("N174").value = DataSheet.Range(hFind("LOS", "6H") & userRow).value
    PrintSheet.Range("D176").value = DataSheet.Range(hFind("Notes on 6H", "6H") & userRow).value

    'Adult
    PrintSheet.Range("D184").value = Lookup("Generic_YNOU_Num")(DataSheet.Range(hFind("Was Youth in Adult?") & userRow).value)
    PrintSheet.Range("G184").value = DataSheet.Range(hFind("Start Date", "Adult") & userRow).value
    PrintSheet.Range("J184").value = DataSheet.Range(hFind("End Date", "Adult") & userRow).value
    PrintSheet.Range("N184").value = DataSheet.Range(hFind("LOS", "Adult") & userRow).value
    PrintSheet.Range("D186").value = DataSheet.Range(hFind("Notes on Adult", "Adult") & userRow).value



    'LEGAL STATUS HISTORY
    'Diversion
    PrintSheet.Range("S84").value = Lookup("Generic_YNOU_Num")(DataSheet.Range(hFind("Referred to Diversion?") & userRow).value)
    PrintSheet.Range("V84").value = DataSheet.Range(hFind("Referral Date", "DIVERSION") & userRow).value
    PrintSheet.Range("Z84").value = DataSheet.Range(hFind("Discharge Date", "DIVERSION") & userRow).value
    PrintSheet.Range("AC84").value = DataSheet.Range(hFind("LOS Diversion") & userRow).value
    PrintSheet.Range("V88").value = DataSheet.Range(hFind("Diversion Notes") & userRow).value

    'Pretrial
    PrintSheet.Range("S96").value = Lookup("Generic_YNOU_Num")(DataSheet.Range(hFind("Was Youth on Pretrial?", "AGGREGATES") & userRow).value)
    PrintSheet.Range("V96").value = DataSheet.Range(hFind("Start Date", "Pretrial", "AGGREGATES") & userRow).value
    PrintSheet.Range("Z96").value = DataSheet.Range(hFind("End Date", "Pretrial", "AGGREGATES") & userRow).value
    PrintSheet.Range("AC96").value = DataSheet.Range(hFind("LOS", "Pretrial", "AGGREGATES") & userRow).value
    PrintSheet.Range("V98").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Origin", "Pretrial", "AGGREGATES") & userRow).value)
    PrintSheet.Range("Z98").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Discharging Courtroom", "Pretrial", "AGGREGATES") & userRow).value)
    PrintSheet.Range("V100").value = DataSheet.Range(hFind("Notes on Pretrial", "AGGREGATES") & userRow).value

    'Consent Decree
    PrintSheet.Range("S108").value = Lookup("Generic_YNOU_Num")(DataSheet.Range(hFind("Was Youth on Consent Decree?", "AGGREGATES") & userRow).value)
    PrintSheet.Range("V108").value = DataSheet.Range(hFind("Start Date", "Consent Decree", "AGGREGATES") & userRow).value
    PrintSheet.Range("Z108").value = DataSheet.Range(hFind("End Date", "Consent Decree", "AGGREGATES") & userRow).value
    PrintSheet.Range("AC108").value = DataSheet.Range(hFind("LOS", "Consent Decree", "AGGREGATES") & userRow).value
    PrintSheet.Range("V110").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Origin", "Consent Decree", "AGGREGATES") & userRow).value)
    PrintSheet.Range("Z110").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Discharging Courtroom", "Consent Decree", "AGGREGATES") & userRow).value)
    PrintSheet.Range("V112").value = DataSheet.Range(hFind("Notes on Consent Decree", "AGGREGATES") & userRow).value

    'Interim/Deferred
    PrintSheet.Range("S120").value = Lookup("Generic_YNOU_Num")(DataSheet.Range(hFind("Was Youth on Interim Probation?", "AGGREGATES") & userRow).value)
    PrintSheet.Range("V120").value = DataSheet.Range(hFind("Start Date", "Interim Probation", "AGGREGATES") & userRow).value
    PrintSheet.Range("Z120").value = DataSheet.Range(hFind("End Date", "Interim Probation", "AGGREGATES") & userRow).value
    PrintSheet.Range("AC120").value = DataSheet.Range(hFind("LOS", "Interim Probation", "AGGREGATES") & userRow).value
    PrintSheet.Range("V122").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Origin", "Interim Probation", "AGGREGATES") & userRow).value)
    PrintSheet.Range("Z122").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Discharging Courtroom", "Interim Probation", "AGGREGATES") & userRow).value)
    PrintSheet.Range("V124").value = DataSheet.Range(hFind("Notes on Interim Probation", "AGGREGATES") & userRow).value


    'Probation
    PrintSheet.Range("S132").value = Lookup("Generic_YNOU_Num")(DataSheet.Range(hFind("Was Youth on Probation?", "AGGREGATES") & userRow).value)
    PrintSheet.Range("V132").value = DataSheet.Range(hFind("Start Date", "Probation", "AGGREGATES") & userRow).value
    PrintSheet.Range("Z132").value = DataSheet.Range(hFind("End Date", "Probation", "AGGREGATES") & userRow).value
    PrintSheet.Range("AC132").value = DataSheet.Range(hFind("LOS", "Probation", "AGGREGATES") & userRow).value
    PrintSheet.Range("V134").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Origin", "Probation", "AGGREGATES") & userRow).value)
    PrintSheet.Range("Z134").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Discharging Courtroom", "Probation", "AGGREGATES") & userRow).value)
    PrintSheet.Range("V136").value = DataSheet.Range(hFind("Notes on Probation", "AGGREGATES") & userRow).value

    'Aftercare Probation
    PrintSheet.Range("S144").value = Lookup("Generic_YNOU_Num")(DataSheet.Range(hFind("Was Youth on Aftercare Probation?", "AGGREGATES") & userRow).value)
    PrintSheet.Range("V144").value = DataSheet.Range(hFind("Start Date", "Aftercare Probation", "AGGREGATES") & userRow).value
    PrintSheet.Range("Z144").value = DataSheet.Range(hFind("End Date", "Aftercare Probation", "AGGREGATES") & userRow).value
    PrintSheet.Range("AC144").value = DataSheet.Range(hFind("LOS", "Aftercare Probation", "AGGREGATES") & userRow).value
    PrintSheet.Range("V146").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Origin", "Aftercare Probation", "AGGREGATES") & userRow).value)
    PrintSheet.Range("Z146").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Discharging Courtroom", "Aftercare Probation", "AGGREGATES") & userRow).value)
    PrintSheet.Range("V148").value = DataSheet.Range(hFind("Notes on Aftercare Probation", "AGGREGATES") & userRow).value

    'Adult
    PrintSheet.Range("S144").value = Lookup("Generic_YNOU_Num")(DataSheet.Range(hFind("Was Youth on Aftercare Probation?", "AGGREGATES") & userRow).value)
    PrintSheet.Range("V144").value = DataSheet.Range(hFind("Start Date", "Aftercare Probation", "AGGREGATES") & userRow).value
    PrintSheet.Range("Z144").value = DataSheet.Range(hFind("End Date", "Aftercare Probation", "AGGREGATES") & userRow).value
    PrintSheet.Range("AC144").value = DataSheet.Range(hFind("LOS", "Aftercare Probation", "AGGREGATES") & userRow).value
    PrintSheet.Range("V146").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Origin", "Aftercare Probation", "AGGREGATES") & userRow).value)
    PrintSheet.Range("Z146").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Discharging Courtroom", "Aftercare Probation", "AGGREGATES") & userRow).value)
    PrintSheet.Range("V148").value = DataSheet.Range(hFind("Notes on Aftercare Probation", "AGGREGATES") & userRow).value


    Call YouthSearchPrint2

End Sub


Sub YouthSearchPrint2()

    Call RefreshNamedRanges
    Call Generate_Dictionaries

    Dim PrintSheet As Worksheet
    Dim DataSheet As Worksheet
    Set PrintSheet = Worksheets("Youth Search")
    Set DataSheet = Worksheets("Entry")

    Dim userRow As Long
    userRow = PrintSheet.Range("J5").value


    'SUPERVISION HISTORY
    '#1
    PrintSheet.Range("D200").value = Lookup("Supervision_Program_Num")(DataSheet.Range(hFind("Supervision Ordered #1", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("G200").value = DataSheet.Range(hFind("Start Date", "Supervision Ordered #1", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G202").value = DataSheet.Range(hFind("End Date", "Supervision Ordered #1", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G204").value = DataSheet.Range(hFind("LOS", "Supervision Ordered #1", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("E206").value = DataSheet.Range(hFind("Supervision Description", "Supervision Ordered #1", "Supervision Programs", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Supervision Ordered #1", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("K200").value = Lookup("Community_Based_Supervision_Provider_Num")(DataSheet.Range(hFind("Community-Based Agency", "Supervision Ordered #1", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("N200").value = Lookup("Residential_Supervision_Provider_Num")(DataSheet.Range(hFind("Residential Agency", "Supervision Ordered #1", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K202").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Supervision Ordered #1", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K204").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Supervision Ordered #1", "Supervision Programs", "AGGREGATES") & userRow).value)

    '#2
    PrintSheet.Range("D214").value = Lookup("Supervision_Program_Num")(DataSheet.Range(hFind("Supervision Ordered #2", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("G214").value = DataSheet.Range(hFind("Start Date", "Supervision Ordered #2", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G216").value = DataSheet.Range(hFind("End Date", "Supervision Ordered #2", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G218").value = DataSheet.Range(hFind("LOS", "Supervision Ordered #2", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("E220").value = DataSheet.Range(hFind("Supervision Description", "Supervision Ordered #2", "Supervision Programs", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Supervision Ordered #2", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("K214").value = Lookup("Community_Based_Supervision_Provider_Num")(DataSheet.Range(hFind("Community-Based Agency", "Supervision Ordered #2", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("N214").value = Lookup("Residential_Supervision_Provider_Num")(DataSheet.Range(hFind("Residential Agency", "Supervision Ordered #2", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K216").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Supervision Ordered #2", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K218").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Supervision Ordered #2", "Supervision Programs", "AGGREGATES") & userRow).value)

    '#3
    PrintSheet.Range("D228").value = Lookup("Supervision_Program_Num")(DataSheet.Range(hFind("Supervision Ordered #3", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("G228").value = DataSheet.Range(hFind("Start Date", "Supervision Ordered #3", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G230").value = DataSheet.Range(hFind("End Date", "Supervision Ordered #3", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G232").value = DataSheet.Range(hFind("LOS", "Supervision Ordered #3", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("E234").value = DataSheet.Range(hFind("Supervision Description", "Supervision Ordered #3", "Supervision Programs", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Supervision Ordered #3", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("K228").value = Lookup("Community_Based_Supervision_Provider_Num")(DataSheet.Range(hFind("Community-Based Agency", "Supervision Ordered #3", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("N228").value = Lookup("Residential_Supervision_Provider_Num")(DataSheet.Range(hFind("Residential Agency", "Supervision Ordered #3", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K230").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Supervision Ordered #3", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K232").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Supervision Ordered #3", "Supervision Programs", "AGGREGATES") & userRow).value)

    '#4
    PrintSheet.Range("D242").value = Lookup("Supervision_Program_Num")(DataSheet.Range(hFind("Supervision Ordered #4", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("G242").value = DataSheet.Range(hFind("Start Date", "Supervision Ordered #4", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G244").value = DataSheet.Range(hFind("End Date", "Supervision Ordered #4", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G246").value = DataSheet.Range(hFind("LOS", "Supervision Ordered #4", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("E248").value = DataSheet.Range(hFind("Supervision Description", "Supervision Ordered #4", "Supervision Programs", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Supervision Ordered #4", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("K242").value = Lookup("Community_Based_Supervision_Provider_Num")(DataSheet.Range(hFind("Community-Based Agency", "Supervision Ordered #4", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("N242").value = Lookup("Residential_Supervision_Provider_Num")(DataSheet.Range(hFind("Residential Agency", "Supervision Ordered #4", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K244").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Supervision Ordered #4", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K246").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Supervision Ordered #4", "Supervision Programs", "AGGREGATES") & userRow).value)

    '#5
    PrintSheet.Range("D256").value = Lookup("Supervision_Program_Num")(DataSheet.Range(hFind("Supervision Ordered #5", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("G256").value = DataSheet.Range(hFind("Start Date", "Supervision Ordered #5", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G258").value = DataSheet.Range(hFind("End Date", "Supervision Ordered #5", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G260").value = DataSheet.Range(hFind("LOS", "Supervision Ordered #5", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("E262").value = DataSheet.Range(hFind("Supervision Description", "Supervision Ordered #5", "Supervision Programs", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Supervision Ordered #5", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("K256").value = Lookup("Community_Based_Supervision_Provider_Num")(DataSheet.Range(hFind("Community-Based Agency", "Supervision Ordered #5", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("N256").value = Lookup("Residential_Supervision_Provider_Num")(DataSheet.Range(hFind("Residential Agency", "Supervision Ordered #5", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K258").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Supervision Ordered #5", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K260").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Supervision Ordered #5", "Supervision Programs", "AGGREGATES") & userRow).value)

    '#6
    PrintSheet.Range("D270").value = Lookup("Supervision_Program_Num")(DataSheet.Range(hFind("Supervision Ordered #6", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("G270").value = DataSheet.Range(hFind("Start Date", "Supervision Ordered #6", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G272").value = DataSheet.Range(hFind("End Date", "Supervision Ordered #6", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G274").value = DataSheet.Range(hFind("LOS", "Supervision Ordered #6", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("E276").value = DataSheet.Range(hFind("Supervision Description", "Supervision Ordered #6", "Supervision Programs", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Supervision Ordered #6", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("K270").value = Lookup("Community_Based_Supervision_Provider_Num")(DataSheet.Range(hFind("Community-Based Agency", "Supervision Ordered #6", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("N270").value = Lookup("Residential_Supervision_Provider_Num")(DataSheet.Range(hFind("Residential Agency", "Supervision Ordered #6", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K272").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Supervision Ordered #6", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K274").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Supervision Ordered #6", "Supervision Programs", "AGGREGATES") & userRow).value)

    '#7
    PrintSheet.Range("D284").value = Lookup("Supervision_Program_Num")(DataSheet.Range(hFind("Supervision Ordered #7", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("G284").value = DataSheet.Range(hFind("Start Date", "Supervision Ordered #7", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G286").value = DataSheet.Range(hFind("End Date", "Supervision Ordered #7", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G288").value = DataSheet.Range(hFind("LOS", "Supervision Ordered #7", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("E290").value = DataSheet.Range(hFind("Supervision Description", "Supervision Ordered #7", "Supervision Programs", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Supervision Ordered #7", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("K284").value = Lookup("Community_Based_Supervision_Provider_Num")(DataSheet.Range(hFind("Community-Based Agency", "Supervision Ordered #7", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("N284").value = Lookup("Residential_Supervision_Provider_Num")(DataSheet.Range(hFind("Residential Agency", "Supervision Ordered #7", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K286").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Supervision Ordered #7", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K288").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Supervision Ordered #7", "Supervision Programs", "AGGREGATES") & userRow).value)

    '#8
    PrintSheet.Range("D298").value = Lookup("Supervision_Program_Num")(DataSheet.Range(hFind("Supervision Ordered #8", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("G298").value = DataSheet.Range(hFind("Start Date", "Supervision Ordered #8", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G300").value = DataSheet.Range(hFind("End Date", "Supervision Ordered #8", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G302").value = DataSheet.Range(hFind("LOS", "Supervision Ordered #8", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("E304").value = DataSheet.Range(hFind("Supervision Description", "Supervision Ordered #8", "Supervision Programs", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Supervision Ordered #8", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("K298").value = Lookup("Community_Based_Supervision_Provider_Num")(DataSheet.Range(hFind("Community-Based Agency", "Supervision Ordered #8", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("N298").value = Lookup("Residential_Supervision_Provider_Num")(DataSheet.Range(hFind("Residential Agency", "Supervision Ordered #8", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K300").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Supervision Ordered #8", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K302").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Supervision Ordered #8", "Supervision Programs", "AGGREGATES") & userRow).value)

    '#9
    PrintSheet.Range("D312").value = Lookup("Supervision_Program_Num")(DataSheet.Range(hFind("Supervision Ordered #9", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("G312").value = DataSheet.Range(hFind("Start Date", "Supervision Ordered #9", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G314").value = DataSheet.Range(hFind("End Date", "Supervision Ordered #9", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G316").value = DataSheet.Range(hFind("LOS", "Supervision Ordered #9", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("E318").value = DataSheet.Range(hFind("Supervision Description", "Supervision Ordered #9", "Supervision Programs", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Supervision Ordered #9", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("K312").value = Lookup("Community_Based_Supervision_Provider_Num")(DataSheet.Range(hFind("Community-Based Agency", "Supervision Ordered #9", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("N312").value = Lookup("Residential_Supervision_Provider_Num")(DataSheet.Range(hFind("Residential Agency", "Supervision Ordered #9", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K314").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Supervision Ordered #9", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K316").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Supervision Ordered #9", "Supervision Programs", "AGGREGATES") & userRow).value)

    '#10
    PrintSheet.Range("D326").value = Lookup("Supervision_Program_Num")(DataSheet.Range(hFind("Supervision Ordered #10", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("G326").value = DataSheet.Range(hFind("Start Date", "Supervision Ordered #10", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G328").value = DataSheet.Range(hFind("End Date", "Supervision Ordered #10", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G330").value = DataSheet.Range(hFind("LOS", "Supervision Ordered #10", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("E332").value = DataSheet.Range(hFind("Supervision Description", "Supervision Ordered #10", "Supervision Programs", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Supervision Ordered #10", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("K326").value = Lookup("Community_Based_Supervision_Provider_Num")(DataSheet.Range(hFind("Community-Based Agency", "Supervision Ordered #10", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("N326").value = Lookup("Residential_Supervision_Provider_Num")(DataSheet.Range(hFind("Residential Agency", "Supervision Ordered #10", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K328").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Supervision Ordered #10", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K330").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Supervision Ordered #10", "Supervision Programs", "AGGREGATES") & userRow).value)

    '#11
    PrintSheet.Range("D340").value = Lookup("Supervision_Program_Num")(DataSheet.Range(hFind("Supervision Ordered #11", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("G340").value = DataSheet.Range(hFind("Start Date", "Supervision Ordered #11", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G342").value = DataSheet.Range(hFind("End Date", "Supervision Ordered #11", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G344").value = DataSheet.Range(hFind("LOS", "Supervision Ordered #11", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("E346").value = DataSheet.Range(hFind("Supervision Description", "Supervision Ordered #11", "Supervision Programs", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Supervision Ordered #11", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("K340").value = Lookup("Community_Based_Supervision_Provider_Num")(DataSheet.Range(hFind("Community-Based Agency", "Supervision Ordered #11", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("N340").value = Lookup("Residential_Supervision_Provider_Num")(DataSheet.Range(hFind("Residential Agency", "Supervision Ordered #11", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K342").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Supervision Ordered #11", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K344").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Supervision Ordered #11", "Supervision Programs", "AGGREGATES") & userRow).value)

    '#12
    PrintSheet.Range("D354").value = Lookup("Supervision_Program_Num")(DataSheet.Range(hFind("Supervision Ordered #12", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("G354").value = DataSheet.Range(hFind("Start Date", "Supervision Ordered #12", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G356").value = DataSheet.Range(hFind("End Date", "Supervision Ordered #12", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G358").value = DataSheet.Range(hFind("LOS", "Supervision Ordered #12", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("E360").value = DataSheet.Range(hFind("Supervision Description", "Supervision Ordered #12", "Supervision Programs", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Supervision Ordered #12", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("K354").value = Lookup("Community_Based_Supervision_Provider_Num")(DataSheet.Range(hFind("Community-Based Agency", "Supervision Ordered #12", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("N354").value = Lookup("Residential_Supervision_Provider_Num")(DataSheet.Range(hFind("Residential Agency", "Supervision Ordered #12", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K356").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Supervision Ordered #12", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K358").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Supervision Ordered #12", "Supervision Programs", "AGGREGATES") & userRow).value)

    '#13
    PrintSheet.Range("D368").value = Lookup("Supervision_Program_Num")(DataSheet.Range(hFind("Supervision Ordered #13", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("G368").value = DataSheet.Range(hFind("Start Date", "Supervision Ordered #13", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G370").value = DataSheet.Range(hFind("End Date", "Supervision Ordered #13", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G372").value = DataSheet.Range(hFind("LOS", "Supervision Ordered #13", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("E374").value = DataSheet.Range(hFind("Supervision Description", "Supervision Ordered #13", "Supervision Programs", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Supervision Ordered #13", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("K368").value = Lookup("Community_Based_Supervision_Provider_Num")(DataSheet.Range(hFind("Community-Based Agency", "Supervision Ordered #13", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("N368").value = Lookup("Residential_Supervision_Provider_Num")(DataSheet.Range(hFind("Residential Agency", "Supervision Ordered #13", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K370").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Supervision Ordered #13", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K372").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Supervision Ordered #13", "Supervision Programs", "AGGREGATES") & userRow).value)

    '#14
    PrintSheet.Range("D382").value = Lookup("Supervision_Program_Num")(DataSheet.Range(hFind("Supervision Ordered #14", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("G382").value = DataSheet.Range(hFind("Start Date", "Supervision Ordered #14", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G384").value = DataSheet.Range(hFind("End Date", "Supervision Ordered #14", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G386").value = DataSheet.Range(hFind("LOS", "Supervision Ordered #14", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("E388").value = DataSheet.Range(hFind("Supervision Description", "Supervision Ordered #14", "Supervision Programs", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Supervision Ordered #14", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("K382").value = Lookup("Community_Based_Supervision_Provider_Num")(DataSheet.Range(hFind("Community-Based Agency", "Supervision Ordered #14", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("N382").value = Lookup("Residential_Supervision_Provider_Num")(DataSheet.Range(hFind("Residential Agency", "Supervision Ordered #14", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K384").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Supervision Ordered #14", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K386").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Supervision Ordered #14", "Supervision Programs", "AGGREGATES") & userRow).value)

    '#15
    PrintSheet.Range("D396").value = Lookup("Supervision_Program_Num")(DataSheet.Range(hFind("Supervision Ordered #15", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("G396").value = DataSheet.Range(hFind("Start Date", "Supervision Ordered #15", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G398").value = DataSheet.Range(hFind("End Date", "Supervision Ordered #15", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G400").value = DataSheet.Range(hFind("LOS", "Supervision Ordered #15", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("E402").value = DataSheet.Range(hFind("Supervision Description", "Supervision Ordered #15", "Supervision Programs", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Supervision Ordered #15", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("K396").value = Lookup("Community_Based_Supervision_Provider_Num")(DataSheet.Range(hFind("Community-Based Agency", "Supervision Ordered #15", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("N396").value = Lookup("Residential_Supervision_Provider_Num")(DataSheet.Range(hFind("Residential Agency", "Supervision Ordered #15", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K398").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Supervision Ordered #15", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K400").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Supervision Ordered #15", "Supervision Programs", "AGGREGATES") & userRow).value)

    '#16
    PrintSheet.Range("D410").value = Lookup("Supervision_Program_Num")(DataSheet.Range(hFind("Supervision Ordered #16", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("G410").value = DataSheet.Range(hFind("Start Date", "Supervision Ordered #16", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G412").value = DataSheet.Range(hFind("End Date", "Supervision Ordered #16", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G414").value = DataSheet.Range(hFind("LOS", "Supervision Ordered #16", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("E416").value = DataSheet.Range(hFind("Supervision Description", "Supervision Ordered #16", "Supervision Programs", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Supervision Ordered #16", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("K410").value = Lookup("Community_Based_Supervision_Provider_Num")(DataSheet.Range(hFind("Community-Based Agency", "Supervision Ordered #16", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("N410").value = Lookup("Residential_Supervision_Provider_Num")(DataSheet.Range(hFind("Residential Agency", "Supervision Ordered #16", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K412").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Supervision Ordered #16", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K414").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Supervision Ordered #16", "Supervision Programs", "AGGREGATES") & userRow).value)

    '#17
    PrintSheet.Range("D424").value = Lookup("Supervision_Program_Num")(DataSheet.Range(hFind("Supervision Ordered #17", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("G424").value = DataSheet.Range(hFind("Start Date", "Supervision Ordered #17", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G426").value = DataSheet.Range(hFind("End Date", "Supervision Ordered #17", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G428").value = DataSheet.Range(hFind("LOS", "Supervision Ordered #17", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("E430").value = DataSheet.Range(hFind("Supervision Description", "Supervision Ordered #17", "Supervision Programs", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Supervision Ordered #17", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("K424").value = Lookup("Community_Based_Supervision_Provider_Num")(DataSheet.Range(hFind("Community-Based Agency", "Supervision Ordered #17", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("N424").value = Lookup("Residential_Supervision_Provider_Num")(DataSheet.Range(hFind("Residential Agency", "Supervision Ordered #17", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K426").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Supervision Ordered #17", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K428").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Supervision Ordered #17", "Supervision Programs", "AGGREGATES") & userRow).value)

    '#18
    PrintSheet.Range("D438").value = Lookup("Supervision_Program_Num")(DataSheet.Range(hFind("Supervision Ordered #18", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("G438").value = DataSheet.Range(hFind("Start Date", "Supervision Ordered #18", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G440").value = DataSheet.Range(hFind("End Date", "Supervision Ordered #18", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G442").value = DataSheet.Range(hFind("LOS", "Supervision Ordered #18", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("E444").value = DataSheet.Range(hFind("Supervision Description", "Supervision Ordered #18", "Supervision Programs", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Supervision Ordered #18", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("K438").value = Lookup("Community_Based_Supervision_Provider_Num")(DataSheet.Range(hFind("Community-Based Agency", "Supervision Ordered #18", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("N438").value = Lookup("Residential_Supervision_Provider_Num")(DataSheet.Range(hFind("Residential Agency", "Supervision Ordered #18", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K440").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Supervision Ordered #18", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K442").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Supervision Ordered #18", "Supervision Programs", "AGGREGATES") & userRow).value)

    '#19
    PrintSheet.Range("D452").value = Lookup("Supervision_Program_Num")(DataSheet.Range(hFind("Supervision Ordered #19", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("G452").value = DataSheet.Range(hFind("Start Date", "Supervision Ordered #19", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G454").value = DataSheet.Range(hFind("End Date", "Supervision Ordered #19", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G456").value = DataSheet.Range(hFind("LOS", "Supervision Ordered #19", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("E458").value = DataSheet.Range(hFind("Supervision Description", "Supervision Ordered #19", "Supervision Programs", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Supervision Ordered #19", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("K452").value = Lookup("Community_Based_Supervision_Provider_Num")(DataSheet.Range(hFind("Community-Based Agency", "Supervision Ordered #19", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("N452").value = Lookup("Residential_Supervision_Provider_Num")(DataSheet.Range(hFind("Residential Agency", "Supervision Ordered #19", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K454").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Supervision Ordered #19", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K456").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Supervision Ordered #19", "Supervision Programs", "AGGREGATES") & userRow).value)

    '#20
    PrintSheet.Range("D466").value = Lookup("Supervision_Program_Num")(DataSheet.Range(hFind("Supervision Ordered #20", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("G466").value = DataSheet.Range(hFind("Start Date", "Supervision Ordered #20", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G468").value = DataSheet.Range(hFind("End Date", "Supervision Ordered #20", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("G470").value = DataSheet.Range(hFind("LOS", "Supervision Ordered #20", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("E472").value = DataSheet.Range(hFind("Supervision Description", "Supervision Ordered #20", "Supervision Programs", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Supervision Ordered #20", "Supervision Programs", "AGGREGATES") & userRow).value
    PrintSheet.Range("K466").value = Lookup("Community_Based_Supervision_Provider_Num")(DataSheet.Range(hFind("Community-Based Agency", "Supervision Ordered #20", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("N466").value = Lookup("Residential_Supervision_Provider_Num")(DataSheet.Range(hFind("Residential Agency", "Supervision Ordered #20", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K468").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Supervision Ordered #20", "Supervision Programs", "AGGREGATES") & userRow).value)
    PrintSheet.Range("K470").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Supervision Ordered #20", "Supervision Programs", "AGGREGATES") & userRow).value)


    Call YouthSearchPrint3

End Sub


Sub YouthSearchPrint3()


    Call RefreshNamedRanges
    Call Generate_Dictionaries

    Dim PrintSheet As Worksheet
    Dim DataSheet As Worksheet
    Set PrintSheet = Worksheets("Youth Search")
    Set DataSheet = Worksheets("Entry")

    Dim userRow As Long
    userRow = PrintSheet.Range("J5").value


    'Conditions HISTORY
    '#1
    PrintSheet.Range("S200").value = Lookup("Condition_Num")(DataSheet.Range(hFind("Condition Ordered #1", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("W200").value = DataSheet.Range(hFind("Start Date", "Condition Ordered #1", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("W202").value = DataSheet.Range(hFind("End Date", "Condition Ordered #1", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("W204").value = DataSheet.Range(hFind("LOS", "Condition Ordered #1", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("U206").value = DataSheet.Range(hFind("Condition Description", "Condition Ordered #1", "Conditions", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Condition Ordered #1", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("AA200").value = Lookup("Condition_Provider_Num")(DataSheet.Range(hFind("Condition Agency", "Condition Ordered #1", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA202").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Condition Ordered #1", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA204").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Condition Ordered #1", "Conditions", "AGGREGATES") & userRow).value)

    '#2
    PrintSheet.Range("S214").value = Lookup("Condition_Num")(DataSheet.Range(hFind("Condition Ordered #2", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("W214").value = DataSheet.Range(hFind("Start Date", "Condition Ordered #2", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("W216").value = DataSheet.Range(hFind("End Date", "Condition Ordered #2", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("W218").value = DataSheet.Range(hFind("LOS", "Condition Ordered #2", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("U220").value = DataSheet.Range(hFind("Condition Description", "Condition Ordered #2", "Conditions", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Condition Ordered #2", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("AA214").value = Lookup("Condition_Provider_Num")(DataSheet.Range(hFind("Condition Agency", "Condition Ordered #2", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA216").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Condition Ordered #2", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA218").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Condition Ordered #2", "Conditions", "AGGREGATES") & userRow).value)

    '#3
    PrintSheet.Range("S228").value = Lookup("Condition_Num")(DataSheet.Range(hFind("Condition Ordered #3", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("W228").value = DataSheet.Range(hFind("Start Date", "Condition Ordered #3", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("W230").value = DataSheet.Range(hFind("End Date", "Condition Ordered #3", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("W232").value = DataSheet.Range(hFind("LOS", "Condition Ordered #3", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("U234").value = DataSheet.Range(hFind("Condition Description", "Condition Ordered #3", "Conditions", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Condition Ordered #3", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("AA228").value = Lookup("Condition_Provider_Num")(DataSheet.Range(hFind("Condition Agency", "Condition Ordered #3", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA230").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Condition Ordered #3", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA232").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Condition Ordered #3", "Conditions", "AGGREGATES") & userRow).value)

    '#4
    PrintSheet.Range("S242").value = Lookup("Condition_Num")(DataSheet.Range(hFind("Condition Ordered #4", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("W242").value = DataSheet.Range(hFind("Start Date", "Condition Ordered #4", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("W244").value = DataSheet.Range(hFind("End Date", "Condition Ordered #4", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("W246").value = DataSheet.Range(hFind("LOS", "Condition Ordered #4", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("U248").value = DataSheet.Range(hFind("Condition Description", "Condition Ordered #4", "Conditions", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Condition Ordered #4", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("AA242").value = Lookup("Condition_Provider_Num")(DataSheet.Range(hFind("Condition Agency", "Condition Ordered #4", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA244").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Condition Ordered #4", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA246").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Condition Ordered #4", "Conditions", "AGGREGATES") & userRow).value)

    '#5
    PrintSheet.Range("S256").value = Lookup("Condition_Num")(DataSheet.Range(hFind("Condition Ordered #5", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("W256").value = DataSheet.Range(hFind("Start Date", "Condition Ordered #5", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("W258").value = DataSheet.Range(hFind("End Date", "Condition Ordered #5", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("W260").value = DataSheet.Range(hFind("LOS", "Condition Ordered #5", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("U262").value = DataSheet.Range(hFind("Condition Description", "Condition Ordered #5", "Conditions", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Condition Ordered #5", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("AA256").value = Lookup("Condition_Provider_Num")(DataSheet.Range(hFind("Condition Agency", "Condition Ordered #5", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA258").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Condition Ordered #5", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA260").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Condition Ordered #5", "Conditions", "AGGREGATES") & userRow).value)

    '#6
    PrintSheet.Range("S270").value = Lookup("Condition_Num")(DataSheet.Range(hFind("Condition Ordered #6", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("W270").value = DataSheet.Range(hFind("Start Date", "Condition Ordered #6", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("W272").value = DataSheet.Range(hFind("End Date", "Condition Ordered #6", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("W274").value = DataSheet.Range(hFind("LOS", "Condition Ordered #6", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("U276").value = DataSheet.Range(hFind("Condition Description", "Condition Ordered #6", "Conditions", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Condition Ordered #6", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("AA270").value = Lookup("Condition_Provider_Num")(DataSheet.Range(hFind("Condition Agency", "Condition Ordered #6", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA272").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Condition Ordered #6", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA274").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Condition Ordered #6", "Conditions", "AGGREGATES") & userRow).value)


    '#7
    PrintSheet.Range("S284").value = Lookup("Condition_Num")(DataSheet.Range(hFind("Condition Ordered #7", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("V284").value = DataSheet.Range(hFind("Start Date", "Condition Ordered #7", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V286").value = DataSheet.Range(hFind("End Date", "Condition Ordered #7", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V288").value = DataSheet.Range(hFind("LOS", "Condition Ordered #7", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("U290").value = DataSheet.Range(hFind("Condition Description", "Condition Ordered #7", "Conditions", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Condition Ordered #7", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("AA284").value = Lookup("Condition_Provider_Num")(DataSheet.Range(hFind("Condition Agency", "Condition Ordered #7", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA286").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Condition Ordered #7", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA288").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Condition Ordered #7", "Conditions", "AGGREGATES") & userRow).value)

    '#8
    PrintSheet.Range("S298").value = Lookup("Condition_Num")(DataSheet.Range(hFind("Condition Ordered #8", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("V298").value = DataSheet.Range(hFind("Start Date", "Condition Ordered #8", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V300").value = DataSheet.Range(hFind("End Date", "Condition Ordered #8", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V302").value = DataSheet.Range(hFind("LOS", "Condition Ordered #8", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("U304").value = DataSheet.Range(hFind("Condition Description", "Condition Ordered #8", "Conditions", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Condition Ordered #8", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("AA298").value = Lookup("Condition_Provider_Num")(DataSheet.Range(hFind("Condition Agency", "Condition Ordered #8", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA300").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Condition Ordered #8", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA302").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Condition Ordered #8", "Conditions", "AGGREGATES") & userRow).value)

    '#9
    PrintSheet.Range("S312").value = Lookup("Condition_Num")(DataSheet.Range(hFind("Condition Ordered #9", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("V312").value = DataSheet.Range(hFind("Start Date", "Condition Ordered #9", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V314").value = DataSheet.Range(hFind("End Date", "Condition Ordered #9", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V316").value = DataSheet.Range(hFind("LOS", "Condition Ordered #9", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("U318").value = DataSheet.Range(hFind("Condition Description", "Condition Ordered #9", "Conditions", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Condition Ordered #9", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("AA312").value = Lookup("Condition_Provider_Num")(DataSheet.Range(hFind("Condition Agency", "Condition Ordered #9", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA314").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Condition Ordered #9", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA316").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Condition Ordered #9", "Conditions", "AGGREGATES") & userRow).value)

    '#10
    PrintSheet.Range("S326").value = Lookup("Condition_Num")(DataSheet.Range(hFind("Condition Ordered #10", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("V326").value = DataSheet.Range(hFind("Start Date", "Condition Ordered #10", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V328").value = DataSheet.Range(hFind("End Date", "Condition Ordered #10", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V330").value = DataSheet.Range(hFind("LOS", "Condition Ordered #10", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("U332").value = DataSheet.Range(hFind("Condition Description", "Condition Ordered #10", "Conditions", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Condition Ordered #10", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("AA326").value = Lookup("Condition_Provider_Num")(DataSheet.Range(hFind("Condition Agency", "Condition Ordered #10", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA328").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Condition Ordered #10", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA330").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Condition Ordered #10", "Conditions", "AGGREGATES") & userRow).value)


    '#11
    PrintSheet.Range("S340").value = Lookup("Condition_Num")(DataSheet.Range(hFind("Condition Ordered #11", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("V340").value = DataSheet.Range(hFind("Start Date", "Condition Ordered #11", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V342").value = DataSheet.Range(hFind("End Date", "Condition Ordered #11", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V344").value = DataSheet.Range(hFind("LOS", "Condition Ordered #11", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("U346").value = DataSheet.Range(hFind("Condition Description", "Condition Ordered #11", "Conditions", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Condition Ordered #11", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("AA340").value = Lookup("Condition_Provider_Num")(DataSheet.Range(hFind("Condition Agency", "Condition Ordered #11", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA342").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Condition Ordered #11", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA344").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Condition Ordered #11", "Conditions", "AGGREGATES") & userRow).value)

    '#12
    PrintSheet.Range("S354").value = Lookup("Condition_Num")(DataSheet.Range(hFind("Condition Ordered #12", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("V354").value = DataSheet.Range(hFind("Start Date", "Condition Ordered #12", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V356").value = DataSheet.Range(hFind("End Date", "Condition Ordered #12", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V358").value = DataSheet.Range(hFind("LOS", "Condition Ordered #12", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("U360").value = DataSheet.Range(hFind("Condition Description", "Condition Ordered #12", "Conditions", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Condition Ordered #12", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("AA354").value = Lookup("Condition_Provider_Num")(DataSheet.Range(hFind("Condition Agency", "Condition Ordered #12", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA356").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Condition Ordered #12", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA358").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Condition Ordered #12", "Conditions", "AGGREGATES") & userRow).value)

    '#13
    PrintSheet.Range("S368").value = Lookup("Condition_Num")(DataSheet.Range(hFind("Condition Ordered #13", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("V368").value = DataSheet.Range(hFind("Start Date", "Condition Ordered #13", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V370").value = DataSheet.Range(hFind("End Date", "Condition Ordered #13", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V372").value = DataSheet.Range(hFind("LOS", "Condition Ordered #13", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("U374").value = DataSheet.Range(hFind("Condition Description", "Condition Ordered #13", "Conditions", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Condition Ordered #13", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("AA368").value = Lookup("Condition_Provider_Num")(DataSheet.Range(hFind("Condition Agency", "Condition Ordered #13", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA370").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Condition Ordered #13", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA372").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Condition Ordered #13", "Conditions", "AGGREGATES") & userRow).value)

    '#14
    PrintSheet.Range("S382").value = Lookup("Condition_Num")(DataSheet.Range(hFind("Condition Ordered #14", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("V382").value = DataSheet.Range(hFind("Start Date", "Condition Ordered #14", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V384").value = DataSheet.Range(hFind("End Date", "Condition Ordered #14", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V386").value = DataSheet.Range(hFind("LOS", "Condition Ordered #14", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("U388").value = DataSheet.Range(hFind("Condition Description", "Condition Ordered #14", "Conditions", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Condition Ordered #14", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("AA382").value = Lookup("Condition_Provider_Num")(DataSheet.Range(hFind("Condition Agency", "Condition Ordered #14", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA384").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Condition Ordered #14", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA386").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Condition Ordered #14", "Conditions", "AGGREGATES") & userRow).value)

    '#15
    PrintSheet.Range("S396").value = Lookup("Condition_Num")(DataSheet.Range(hFind("Condition Ordered #15", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("V396").value = DataSheet.Range(hFind("Start Date", "Condition Ordered #15", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V398").value = DataSheet.Range(hFind("End Date", "Condition Ordered #15", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V400").value = DataSheet.Range(hFind("LOS", "Condition Ordered #15", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("U402").value = DataSheet.Range(hFind("Condition Description", "Condition Ordered #15", "Conditions", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Condition Ordered #15", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("AA396").value = Lookup("Condition_Provider_Num")(DataSheet.Range(hFind("Condition Agency", "Condition Ordered #15", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA398").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Condition Ordered #15", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA400").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Condition Ordered #15", "Conditions", "AGGREGATES") & userRow).value)

    '#16
    PrintSheet.Range("S410").value = Lookup("Condition_Num")(DataSheet.Range(hFind("Condition Ordered #16", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("V410").value = DataSheet.Range(hFind("Start Date", "Condition Ordered #16", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V412").value = DataSheet.Range(hFind("End Date", "Condition Ordered #16", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V414").value = DataSheet.Range(hFind("LOS", "Condition Ordered #16", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("U416").value = DataSheet.Range(hFind("Condition Description", "Condition Ordered #16", "Conditions", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Condition Ordered #16", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("AA410").value = Lookup("Condition_Provider_Num")(DataSheet.Range(hFind("Condition Agency", "Condition Ordered #16", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA412").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Condition Ordered #16", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA414").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Condition Ordered #16", "Conditions", "AGGREGATES") & userRow).value)

    '#17
    PrintSheet.Range("S424").value = Lookup("Condition_Num")(DataSheet.Range(hFind("Condition Ordered #17", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("V424").value = DataSheet.Range(hFind("Start Date", "Condition Ordered #17", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V426").value = DataSheet.Range(hFind("End Date", "Condition Ordered #17", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V428").value = DataSheet.Range(hFind("LOS", "Condition Ordered #17", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("U430").value = DataSheet.Range(hFind("Condition Description", "Condition Ordered #17", "Conditions", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Condition Ordered #17", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("AA424").value = Lookup("Condition_Provider_Num")(DataSheet.Range(hFind("Condition Agency", "Condition Ordered #17", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA426").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Condition Ordered #17", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA428").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Condition Ordered #17", "Conditions", "AGGREGATES") & userRow).value)

    '#18
    PrintSheet.Range("S438").value = Lookup("Condition_Num")(DataSheet.Range(hFind("Condition Ordered #18", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("V438").value = DataSheet.Range(hFind("Start Date", "Condition Ordered #18", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V440").value = DataSheet.Range(hFind("End Date", "Condition Ordered #18", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V442").value = DataSheet.Range(hFind("LOS", "Condition Ordered #18", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("U444").value = DataSheet.Range(hFind("Condition Description", "Condition Ordered #18", "Conditions", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Condition Ordered #18", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("AA438").value = Lookup("Condition_Provider_Num")(DataSheet.Range(hFind("Condition Agency", "Condition Ordered #18", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA440").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Condition Ordered #18", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA442").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Condition Ordered #18", "Conditions", "AGGREGATES") & userRow).value)

    '#19
    PrintSheet.Range("S452").value = Lookup("Condition_Num")(DataSheet.Range(hFind("Condition Ordered #19", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("V452").value = DataSheet.Range(hFind("Start Date", "Condition Ordered #19", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V454").value = DataSheet.Range(hFind("End Date", "Condition Ordered #19", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V456").value = DataSheet.Range(hFind("LOS", "Condition Ordered #19", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("U458").value = DataSheet.Range(hFind("Condition Description", "Condition Ordered #19", "Conditions", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Condition Ordered #19", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("AA452").value = Lookup("Condition_Provider_Num")(DataSheet.Range(hFind("Condition Agency", "Condition Ordered #19", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA454").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Condition Ordered #19", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA456").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Condition Ordered #19", "Conditions", "AGGREGATES") & userRow).value)

    '#20
    PrintSheet.Range("S466").value = Lookup("Condition_Num")(DataSheet.Range(hFind("Condition Ordered #20", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("V466").value = DataSheet.Range(hFind("Start Date", "Condition Ordered #20", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V468").value = DataSheet.Range(hFind("End Date", "Condition Ordered #20", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("V470").value = DataSheet.Range(hFind("LOS", "Condition Ordered #20", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("U472").value = DataSheet.Range(hFind("Condition Description", "Condition Ordered #20", "Conditions", "AGGREGATES") & userRow).value & "; DISCHARGE - " & DataSheet.Range(hFind("Discharge Description", "Condition Ordered #20", "Conditions", "AGGREGATES") & userRow).value
    PrintSheet.Range("K466").value = Lookup("Condition_Provider_Num")(DataSheet.Range(hFind("Condition Agency", "Condition Ordered #20", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA468").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom of Order", "Condition Ordered #20", "Conditions", "AGGREGATES") & userRow).value)
    PrintSheet.Range("AA470").value = Lookup("Legal_Status_Num")(DataSheet.Range(hFind("Legal Status of Order", "Condition Ordered #20", "Conditions", "AGGREGATES") & userRow).value)

    Call YouthSearchPrint4

End Sub


Sub YouthSearchPrint4()


    Call RefreshNamedRanges
    Call Generate_Dictionaries

    Dim PrintSheet As Worksheet
    Dim DataSheet As Worksheet
    Set PrintSheet = Worksheets("Youth Search")
    Set DataSheet = Worksheets("Entry")

    Dim userRow As Long
    userRow = PrintSheet.Range("J5").value



    'COURT LISTINGS HISTORY
    '#1
    PrintSheet.Range("K487").value = DataSheet.Range(hFind("Court Date #1", "LISTINGS") & userRow).value
    PrintSheet.Range("K489").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #1", "LISTINGS") & userRow).value)
    PrintSheet.Range("P487").value = DataSheet.Range(hFind("Notes", "Court Date #1", "LISTINGS") & userRow).value

    '#2
    PrintSheet.Range("K495").value = DataSheet.Range(hFind("Court Date #2", "LISTINGS") & userRow).value
    PrintSheet.Range("K497").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #2", "LISTINGS") & userRow).value)
    PrintSheet.Range("P495").value = DataSheet.Range(hFind("Notes", "Court Date #2", "LISTINGS") & userRow).value

    '#3
    PrintSheet.Range("K503").value = DataSheet.Range(hFind("Court Date #3", "LISTINGS") & userRow).value
    PrintSheet.Range("K505").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #3", "LISTINGS") & userRow).value)
    PrintSheet.Range("P503").value = DataSheet.Range(hFind("Notes", "Court Date #3", "LISTINGS") & userRow).value

    '#4
    PrintSheet.Range("K511").value = DataSheet.Range(hFind("Court Date #4", "LISTINGS") & userRow).value
    PrintSheet.Range("K513").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #4", "LISTINGS") & userRow).value)
    PrintSheet.Range("P511").value = DataSheet.Range(hFind("Notes", "Court Date #4", "LISTINGS") & userRow).value

    '#5
    PrintSheet.Range("K519").value = DataSheet.Range(hFind("Court Date #5", "LISTINGS") & userRow).value
    PrintSheet.Range("K521").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #5", "LISTINGS") & userRow).value)
    PrintSheet.Range("P519").value = DataSheet.Range(hFind("Notes", "Court Date #5", "LISTINGS") & userRow).value

    '#6
    PrintSheet.Range("K527").value = DataSheet.Range(hFind("Court Date #6", "LISTINGS") & userRow).value
    PrintSheet.Range("K529").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #6", "LISTINGS") & userRow).value)
    PrintSheet.Range("P527").value = DataSheet.Range(hFind("Notes", "Court Date #6", "LISTINGS") & userRow).value

    '#7
    PrintSheet.Range("K535").value = DataSheet.Range(hFind("Court Date #7", "LISTINGS") & userRow).value
    PrintSheet.Range("K537").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #7", "LISTINGS") & userRow).value)
    PrintSheet.Range("P535").value = DataSheet.Range(hFind("Notes", "Court Date #7", "LISTINGS") & userRow).value

    '#8
    PrintSheet.Range("K543").value = DataSheet.Range(hFind("Court Date #8", "LISTINGS") & userRow).value
    PrintSheet.Range("K545").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #8", "LISTINGS") & userRow).value)
    PrintSheet.Range("P543").value = DataSheet.Range(hFind("Notes", "Court Date #8", "LISTINGS") & userRow).value

    '#9
    PrintSheet.Range("K551").value = DataSheet.Range(hFind("Court Date #9", "LISTINGS") & userRow).value
    PrintSheet.Range("K553").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #9", "LISTINGS") & userRow).value)
    PrintSheet.Range("P551").value = DataSheet.Range(hFind("Notes", "Court Date #9", "LISTINGS") & userRow).value

    '#10
    PrintSheet.Range("K559").value = DataSheet.Range(hFind("Court Date #10", "LISTINGS") & userRow).value
    PrintSheet.Range("K561").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #10", "LISTINGS") & userRow).value)
    PrintSheet.Range("P559").value = DataSheet.Range(hFind("Notes", "Court Date #10", "LISTINGS") & userRow).value

    '#11
    PrintSheet.Range("K567").value = DataSheet.Range(hFind("Court Date #11", "LISTINGS") & userRow).value
    PrintSheet.Range("K569").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #11", "LISTINGS") & userRow).value)
    PrintSheet.Range("P567").value = DataSheet.Range(hFind("Notes", "Court Date #11", "LISTINGS") & userRow).value

    '#12
    PrintSheet.Range("K575").value = DataSheet.Range(hFind("Court Date #12", "LISTINGS") & userRow).value
    PrintSheet.Range("K577").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #12", "LISTINGS") & userRow).value)
    PrintSheet.Range("P575").value = DataSheet.Range(hFind("Notes", "Court Date #12", "LISTINGS") & userRow).value

    '#13
    PrintSheet.Range("K583").value = DataSheet.Range(hFind("Court Date #13", "LISTINGS") & userRow).value
    PrintSheet.Range("K585").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #13", "LISTINGS") & userRow).value)
    PrintSheet.Range("P583").value = DataSheet.Range(hFind("Notes", "Court Date #13", "LISTINGS") & userRow).value

    '#14
    PrintSheet.Range("K591").value = DataSheet.Range(hFind("Court Date #14", "LISTINGS") & userRow).value
    PrintSheet.Range("K593").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #14", "LISTINGS") & userRow).value)
    PrintSheet.Range("P591").value = DataSheet.Range(hFind("Notes", "Court Date #14", "LISTINGS") & userRow).value

    '#15
    PrintSheet.Range("K599").value = DataSheet.Range(hFind("Court Date #15", "LISTINGS") & userRow).value
    PrintSheet.Range("K601").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #15", "LISTINGS") & userRow).value)
    PrintSheet.Range("P599").value = DataSheet.Range(hFind("Notes", "Court Date #15", "LISTINGS") & userRow).value

    '#16
    PrintSheet.Range("K607").value = DataSheet.Range(hFind("Court Date #16", "LISTINGS") & userRow).value
    PrintSheet.Range("K609").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #16", "LISTINGS") & userRow).value)
    PrintSheet.Range("P607").value = DataSheet.Range(hFind("Notes", "Court Date #16", "LISTINGS") & userRow).value

    '#17
    PrintSheet.Range("K615").value = DataSheet.Range(hFind("Court Date #17", "LISTINGS") & userRow).value
    PrintSheet.Range("K617").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #17", "LISTINGS") & userRow).value)
    PrintSheet.Range("P615").value = DataSheet.Range(hFind("Notes", "Court Date #17", "LISTINGS") & userRow).value

    '#18
    PrintSheet.Range("K623").value = DataSheet.Range(hFind("Court Date #18", "LISTINGS") & userRow).value
    PrintSheet.Range("K625").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #18", "LISTINGS") & userRow).value)
    PrintSheet.Range("P623").value = DataSheet.Range(hFind("Notes", "Court Date #18", "LISTINGS") & userRow).value

    '#19
    PrintSheet.Range("K631").value = DataSheet.Range(hFind("Court Date #19", "LISTINGS") & userRow).value
    PrintSheet.Range("K633").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #19", "LISTINGS") & userRow).value)
    PrintSheet.Range("P631").value = DataSheet.Range(hFind("Notes", "Court Date #19", "LISTINGS") & userRow).value

    '#20
    PrintSheet.Range("K639").value = DataSheet.Range(hFind("Court Date #20", "LISTINGS") & userRow).value
    PrintSheet.Range("K641").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #20", "LISTINGS") & userRow).value)
    PrintSheet.Range("P639").value = DataSheet.Range(hFind("Notes", "Court Date #20", "LISTINGS") & userRow).value

    '#21
    PrintSheet.Range("K647").value = DataSheet.Range(hFind("Court Date #21", "LISTINGS") & userRow).value
    PrintSheet.Range("K649").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #21", "LISTINGS") & userRow).value)
    PrintSheet.Range("P647").value = DataSheet.Range(hFind("Notes", "Court Date #21", "LISTINGS") & userRow).value

    '#22
    PrintSheet.Range("K655").value = DataSheet.Range(hFind("Court Date #22", "LISTINGS") & userRow).value
    PrintSheet.Range("K657").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #22", "LISTINGS") & userRow).value)
    PrintSheet.Range("P655").value = DataSheet.Range(hFind("Notes", "Court Date #22", "LISTINGS") & userRow).value

    '#23
    PrintSheet.Range("K663").value = DataSheet.Range(hFind("Court Date #23", "LISTINGS") & userRow).value
    PrintSheet.Range("K665").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #23", "LISTINGS") & userRow).value)
    PrintSheet.Range("P663").value = DataSheet.Range(hFind("Notes", "Court Date #23", "LISTINGS") & userRow).value

    '#24
    PrintSheet.Range("K671").value = DataSheet.Range(hFind("Court Date #24", "LISTINGS") & userRow).value
    PrintSheet.Range("K673").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #24", "LISTINGS") & userRow).value)
    PrintSheet.Range("P671").value = DataSheet.Range(hFind("Notes", "Court Date #24", "LISTINGS") & userRow).value

    '#25
    PrintSheet.Range("K679").value = DataSheet.Range(hFind("Court Date #25", "LISTINGS") & userRow).value
    PrintSheet.Range("K681").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #25", "LISTINGS") & userRow).value)
    PrintSheet.Range("P679").value = DataSheet.Range(hFind("Notes", "Court Date #25", "LISTINGS") & userRow).value

    '#26
    PrintSheet.Range("K687").value = DataSheet.Range(hFind("Court Date #26", "LISTINGS") & userRow).value
    PrintSheet.Range("K689").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #26", "LISTINGS") & userRow).value)
    PrintSheet.Range("P687").value = DataSheet.Range(hFind("Notes", "Court Date #26", "LISTINGS") & userRow).value

    '#27
    PrintSheet.Range("K695").value = DataSheet.Range(hFind("Court Date #27", "LISTINGS") & userRow).value
    PrintSheet.Range("K697").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #27", "LISTINGS") & userRow).value)
    PrintSheet.Range("P695").value = DataSheet.Range(hFind("Notes", "Court Date #27", "LISTINGS") & userRow).value

    '#28
    PrintSheet.Range("K703").value = DataSheet.Range(hFind("Court Date #28", "LISTINGS") & userRow).value
    PrintSheet.Range("K705").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #28", "LISTINGS") & userRow).value)
    PrintSheet.Range("P703").value = DataSheet.Range(hFind("Notes", "Court Date #28", "LISTINGS") & userRow).value

    '#29
    PrintSheet.Range("K711").value = DataSheet.Range(hFind("Court Date #29", "LISTINGS") & userRow).value
    PrintSheet.Range("K713").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #29", "LISTINGS") & userRow).value)
    PrintSheet.Range("P711").value = DataSheet.Range(hFind("Notes", "Court Date #29", "LISTINGS") & userRow).value

    '#30
    PrintSheet.Range("K719").value = DataSheet.Range(hFind("Court Date #30", "LISTINGS") & userRow).value
    PrintSheet.Range("K721").value = Lookup("Courtroom_Num")(DataSheet.Range(hFind("Courtroom", "Court Date #30", "LISTINGS") & userRow).value)
    PrintSheet.Range("P719").value = DataSheet.Range(hFind("Notes", "Court Date #30", "LISTINGS") & userRow).value




End Sub
