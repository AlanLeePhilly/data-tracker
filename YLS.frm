VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} YLS 
   Caption         =   "UserForm1"
   ClientHeight    =   10470
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7965
   OleObjectBlob   =   "YLS.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "YLS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DateBox_Enter()
    DateBox.value = CalendarForm.GetDate(RangeOfYears:=5)
End Sub


Sub DateBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Set ctl = Me.DateBox

    Call DateValidation(ctl, Cancel)
End Sub

Private Sub PJJSC_Submit_Click()
    
    If DateBox.value = "" Then
        MsgBox "YLS Date required"
        Exit Sub
    End If
    
    'ADD MORE VALIDATIONS HERE IF REQUIRED
    
    
    
    Call formSubmitStart(updateRow)
    
    Dim YLSHead As String
    Dim bucketHead As String
    Dim count As Integer
    Dim NUM_YLS_BUCKETS As Integer
    
    NUM_YLS_BUCKETS = 1 'update as necessary
    
    YLSHead = headerFind("YLS")
    
    'I think this field should be removed and YLS.value should be "Yes" (in quotes)
    'Right now a user is defining that first column every time, but usually that column is flagged yes automatically
    Range(headerFind("Did Youth Have YLS?", YLSHead) & updateRow) = Lookup("Generic_YNOU_Name")(YLS.value)
    
    'find header for first empty bucket
    For count = 1 To NUM_YLS_BUCKETS
        If isEmptyOrZero(Range(hFind("Assessment Date", "YLS #" & count, "YLS") & updateRow)) Then
            bucketHead = hFind("YLS #" & count, "YLS")
            Exit For
        Else
            If count = NUM_YLS_BUCKETS Then
                MsgBox "Selected user has already filled all " & count & " available YLS buckets"
                Exit Sub
            End If
        End If
    Next count
    
    
    'Example Lookup-based entries
    Range(headerFind("First YLS?", bucketHead) & updateRow) = Lookup("Generic_YNOU_Name")(FirstYLS.value)
    Range(headerFind("YLS Given Post Arrest?", bucketHead) & updateRow) = Lookup("Generic_YNOU_Name")(PostArrest.value)
    
    'Example direct fields
    Range(headerFind("Assessment Date", bucketHead) & updateRow) = DateBox.value
    Range(headerFind("Score", bucketHead) & updateRow) = YLSScore.value
    
    
    Range(headerFind("Strength", bucketHead) & updateRow) = Lookup("Generic_YNOU_Name")(YLSStrength.value)
    Range(headerFind("Risk Status", bucketHead) & updateRow) = Lookup("YLS_Level_Name")(YLSRisk.value)
    Range(headerFind("Percent", bucketHead) & updateRow) = YLSPercent.value
    Range(headerFind("Interviewer", bucketHead) & updateRow) = YLSInterviewer.value
    
    'Field based subsections
    Range(headerFind("Prior and Current Offense Score", bucketHead) & updateRow) = Field1Score.value
    Range(headerFind("Prior and Current Offense Risk", bucketHead) & updateRow) = Lookup("YLS_Level_Name")(Field1Risk.value)
    Range(headerFind("Prior and Current Offense Strength", bucketHead) & updateRow) = Lookup("Generic_YNOU_Name")(Field1Strength.value)
    
    Range(headerFind("Family Circumstances Score", bucketHead) & updateRow) = Field2Score.value
    Range(headerFind("Family Circumstances Risk", bucketHead) & updateRow) = Lookup("YLS_Level_Name")(Field2Risk.value)
    Range(headerFind("Family Circumstances Strength", bucketHead) & updateRow) = Lookup("Generic_YNOU_Name")(Field2Strength.value)
    
    Range(headerFind("Education/Employment Score", bucketHead) & updateRow) = Field3Score.value
    Range(headerFind("Education/Employment Risk", bucketHead) & updateRow) = Lookup("YLS_Level_Name")(Field3Risk.value)
    Range(headerFind("Education/Employment Strength", bucketHead) & updateRow) = Lookup("Generic_YNOU_Name")(Field3Strength.value)
    
    Range(headerFind("Peer Relations Score", bucketHead) & updateRow) = Field4Score.value
    Range(headerFind("Peer Relations Risk", bucketHead) & updateRow) = Lookup("YLS_Level_Name")(Field4Risk.value)
    Range(headerFind("Peer Relations Strength", bucketHead) & updateRow) = Lookup("Generic_YNOU_Name")(Field4Strength.value)
    
    Range(headerFind("Substance Abuse Score", bucketHead) & updateRow) = Field5Score.value
    Range(headerFind("Substance Abuse Risk", bucketHead) & updateRow) = Lookup("YLS_Level_Name")(Field5Risk.value)
    Range(headerFind("Substance Abuse Strength", bucketHead) & updateRow) = Lookup("Generic_YNOU_Name")(Field5Strength.value)
    
    Range(headerFind("Leisure/Recreation Score", bucketHead) & updateRow) = Field6Score.value
    Range(headerFind("Leisure/Recreation Risk", bucketHead) & updateRow) = Lookup("YLS_Level_Name")(Field6Risk.value)
    Range(headerFind("Leisure/Recreation Strength", bucketHead) & updateRow) = Lookup("Generic_YNOU_Name")(Field6Strength.value)
    
    Range(headerFind("Personality/Behavior Score", bucketHead) & updateRow) = Field7Score.value
    Range(headerFind("Personality/Behavior Risk", bucketHead) & updateRow) = Lookup("YLS_Level_Name")(Field7Risk.value)
    Range(headerFind("Personality/Behavior Strength", bucketHead) & updateRow) = Lookup("Generic_YNOU_Name")(Field7Strength.value)
    
    Range(headerFind("Attitudes/Orientation Score", bucketHead) & updateRow) = Field8Score.value
    Range(headerFind("Attitudes/Orientation Risk", bucketHead) & updateRow) = Lookup("YLS_Level_Name")(Field8Risk.value)
    Range(headerFind("Attitudes/Orientation Strength", bucketHead) & updateRow) = Lookup("Generic_YNOU_Name")(Field8Strength.value)
    
    
    
    
    'Also if something doesn't work, don't forget to check the names of the form controls!
    'Good luck!
    
    
    Call formSubmitEnd
    Unload Me
End Sub

