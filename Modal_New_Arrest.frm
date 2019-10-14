VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Modal_New_Arrest 
   Caption         =   "UserForm1"
   ClientHeight    =   11610
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17325
   OleObjectBlob   =   "Modal_New_Arrest.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Modal_New_Arrest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AddPetition_Click()
    If PetitionBox.ListCount < 5 Then
        Load Modal_NewClient_Add_Petition
        Modal_NewClient_Add_Petition.headline.Caption = "Re-Arrest"
        Modal_NewClient_Add_Petition.Show
    Else
        MsgBox "Maximum of five petitions for new client"
    End If
End Sub

Private Sub ArrestDate_Enter() 'when a user "Enters" (clicks) the text box...
    'this boxe's value is defined as a result of the picker being called and completed
    ArrestDate.value = CalendarForm.GetDate(RangeOfYears:=5) '(set range from today)
End Sub

Private Sub ArrestDate_Exit(ByVal Cancel As MSForms.ReturnBoolean) 'when a user "Exits" (clicks outside of) the text box...
    Set ctl = Me.ArrestDate 'set the text box to a variable (the whole text box, not its value)
    Call DateValidation(ctl, Cancel) 'send the text box to a custom date validation function
End Sub

Private Sub IncidentDate_Enter()
    IncidentDate.value = CalendarForm.GetDate(RangeOfYears:=5)
End Sub
Private Sub IncidentDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Set ctl = Me.IncidentDate
    Call DateValidation(ctl, Cancel)
End Sub

Private Sub DeletePetition_Click()
    Dim petitionNum As String
    Dim i As Integer
    Dim listIndex As Integer

    If PetitionBox.listIndex = -1 Then
        Exit Sub
    End If

    petitionNum = PetitionBox.List(PetitionBox.listIndex, 1)

    MsgBox "Removing Petition #" & petitionNum
    listIndex = ChargeBox.ListCount - 1
    For i = listIndex To 0 Step -1
        If ChargeBox.List(i, 0) = petitionNum Then
            ChargeBox.RemoveItem (i)
        End If
    Next i
    PetitionBox.RemoveItem (PetitionBox.listIndex)
End Sub


Private Sub Cancel_Click()
    Call Clear_Click
    Modal_New_Arrest.Hide
End Sub

Private Sub Clear_Click()
    Dim ctl As Control ' Removed MSForms.

    For Each ctl In Me.Controls
        Select Case TypeName(ctl)
            Case "TextBox"
                ctl.value = ""
            Case "CheckBox", "ToggleButton" ' Removed OptionButton
                ctl.value = False
            Case "OptionGroup" ' Add OptionGroup
                ctl = Null
            Case "OptionButton" ' Add OptionButton
                ' Do not reset an optionbutton if it is part of an OptionGroup
                If TypeName(ctl.Parent) <> "OptionGroup" Then ctl.value = False
            Case "ComboBox", "ListBox"
                ctl.listIndex = -1
        End Select
    Next ctl
End Sub

Private Sub SameDate_Click()
    ArrestDate.value = IncidentDate.value
End Sub
Private Sub SameTime_Click()
    TimeOfArrest_H.value = TimeOfIncident_H.value
    TimeOfArrest_M.value = TimeOfIncident_M.value
    TimeOfArrest_P.value = TimeOfIncident_P.value
End Sub
Private Sub SameDistrict_Click()
    ArrestingDistrict.value = IncidentDistrict.value
End Sub

Private Sub Submit_Click()
    Dim restorer As Variant

    Call Generate_Dictionaries
    'define variable Long(a big integer) named updateRow
    Dim updateRow As Long

    'activate the spreadsheet as default selector
    Worksheets("Entry").Activate

    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With

    updateRow = Active_Row.Caption
    restorer = Range("C" & updateRow & ":" & hFind("END") & updateRow).value

    If isNotEmptyOrZero(Range(hFind("Arrest Date #5", "REARRESTS", "AGGREGATES") & updateRow)) Then
        MsgBox "It looks like there are already 5 re-arrests on this client. Supported maximum reached"
        Exit Sub
    End If

    On Error GoTo err
    Dim i As Integer
    Dim j As Integer
    Dim sectionHead As String
    Dim bucketHead As String
    Dim altHead As String

    sectionHead = hFind("REARRESTS", "AGGREGATES")
    Call flagYes(Range(headerFind("Was Youth Rearrested?", sectionHead) & updateRow))

    For i = 1 To 5
        If isEmptyOrZero(Range(headerFind("Arrest Date #" & i, sectionHead) & updateRow)) Then
            bucketHead = headerFind("Arrest Date #" & i, sectionHead)
            i = 5
        End If
    Next i

    Range(bucketHead & updateRow).value _
            = ArrestDate.value

    Range(headerFind("Active Courtroom", bucketHead) & updateRow).value _
            = Range(headerFind("Active Courtroom") & updateRow).value
    Range(headerFind("Active Legal Status", bucketHead) & updateRow).value _
            = Range(headerFind("Legal Status") & updateRow).value

    j = 1
    For i = 1 To 30
        altHead = hFind("Supervision Ordered #" & i, "AGGREGATES")
        If isNotEmptyOrZero(Range(headerFind("Start Date", altHead) & updateRow)) _
        And isEmptyOrZero(Range(headerFind("End Date", altHead) & updateRow)) _
        And j < 3 Then

            Range(headerFind("Active Supervision #" & j, bucketHead) & updateRow).value _
                    = Range(altHead & updateRow).value
            Range(headerFind("Active Community-Based Agency #" & j, bucketHead) & updateRow).value _
                    = Range(headerFind("Community-Based Agency", altHead) & updateRow)
            Range(headerFind("Active Residential Agency #" & j, bucketHead) & updateRow).value _
                    = Range(headerFind("Residential Agency", altHead) & updateRow)
            j = j + 1
        End If
    Next i


    Range(headerFind("Day of Arrest", bucketHead) & updateRow).value _
            = Weekday(ArrestDate.value, vbMonday) * 2 - 1
    Range(headerFind("Time of Arrest", bucketHead) & updateRow).value _
            = TimeOfArrest_H.value & ":" & TimeOfArrest_M.value & " " & TimeOfArrest_P.value
    Range(headerFind("Time Category of Arrest", bucketHead) & updateRow).value _
            = calcTimeGroup(TimeOfArrest_H.value, TimeOfArrest_P.value)
    Range(headerFind("Arresting District", bucketHead) & updateRow).value _
            = ArrestingDistrict.value
    Range(headerFind("Time of Referral to DA") & updateRow).value _
            = TimeReferredToDA_H.value & ":" & TimeReferredToDA_M.value & " " & TimeReferredToDA_P.value

    Range(headerFind("DC #", bucketHead) & updateRow).value = DCNum.value
    Range(headerFind("PID #", bucketHead) & updateRow).value = PIDNum.value
    Range(headerFind("DC-PID #", bucketHead) & updateRow).value = DCNum.value & "-" & PIDNum.value
    Range(headerFind("SID #", bucketHead) & updateRow).value = SIDNum.value

    Range(headerFind("Officer #1", bucketHead) & updateRow).value = Officer1.value
    Range(headerFind("Officer #2", bucketHead) & updateRow).value = Officer2.value
    Range(headerFind("Officer #3", bucketHead) & updateRow).value = Officer3.value
    Range(headerFind("Officer #4", bucketHead) & updateRow).value = Officer4.value
    Range(headerFind("Officer #5", bucketHead) & updateRow).value = Officer5.value

    Range(headerFind("Victim First Name", bucketHead) & updateRow) _
            = VictimFirstName.value
    Range(headerFind("Victim Last Name", bucketHead) & updateRow) _
            = VictimLastName.value

    Range(headerFind("Incident Date", bucketHead) & updateRow).value _
            = IncidentDate.value
    Range(headerFind("Day of Incident", bucketHead) & updateRow).value _
            = Weekday(IncidentDate.value, vbMonday) * 2 - 1
    Range(headerFind("Time of Incident", bucketHead) & updateRow).value _
            = TimeOfIncident_H.value & ":" & TimeOfIncident_M.value & " " & TimeOfIncident_P.value
    Range(headerFind("Time Category of Incident", bucketHead) & updateRow).value _
            = calcTimeGroup(TimeOfIncident_H.value, TimeOfIncident_P.value)
    Range(headerFind("Incident District", bucketHead) & updateRow).value _
            = IncidentDistrict.value
    Range(headerFind("Incident Address", bucketHead) & updateRow).value _
            = IncidentAddress.value
    Range(headerFind("Incident Zipcode", bucketHead) & updateRow).value _
            = IncidentZipcode.value
    Range(headerFind("LOS Until Rearrest", bucketHead) & updateRow).value _
                = calcLOS(Range(hFind("Arrest Date", "PETITION") & updateRow).value, ArrestDate.value)



    Dim Num As Long

    For Num = 1 To PetitionBox.ListCount
        tempHead = headerFind("Petition #" & Num, bucketHead)

        Range(headerFind("Petition Filed?", tempHead) & updateRow).value _
                = Lookup("Generic_YNOU_Name")("Yes")
        Range(headerFind("Was Petition Transferred from Other County?", tempHead) & updateRow).value _
                = Lookup("Generic_YNOU_Name")(PetitionBox.List(Num - 1, 6))
        Range(tempHead & updateRow).value _
                = PetitionBox.List(Num - 1, 1)
        Range(headerFind("Date Filed", tempHead) & updateRow).value _
                = PetitionBox.List(Num - 1, 0)
        Range(headerFind("Lead Charge Code", tempHead) & updateRow).value _
                = PetitionBox.List(Num - 1, 4)
        Range(headerFind("Lead Charge Name", tempHead) & updateRow).value _
                = PetitionBox.List(Num - 1, 5)
        Range(headerFind("Charge Category #1", tempHead) & updateRow).value _
                = Lookup("Charge_Name")(PetitionBox.List(Num - 1, 3))
        Range(headerFind("Charge Grade (specific) #1", tempHead) & updateRow).value _
                = Lookup("Charge_Grade_Specific_Name")(PetitionBox.List(Num - 1, 2))
        Range(headerFind("Charge Grade (broad) #1", tempHead) & updateRow).value _
                = calcChargeBroad(PetitionBox.List(Num - 1, 2))

        j = 2
        For i = 0 To ChargeBox.ListCount - 1
            If ChargeBox.ListCount > 0 Then
                If ChargeBox.List(i, 0) = PetitionBox.List(Num - 1, 1) Then
                    Range(headerFind("Charge Code #" & j, tempHead) & updateRow).value _
                            = ChargeBox.List(i, 3)
                    Range(headerFind("Charge Name #" & j, tempHead) & updateRow).value _
                            = ChargeBox.List(i, 4)
                    Range(headerFind("Charge Category #" & j, tempHead) & updateRow).value _
                            = Lookup("Charge_Name")(ChargeBox.List(i, 2))
                    Range(headerFind("Charge Grade (specific) #" & j, tempHead) & updateRow).value _
                            = Lookup("Charge_Grade_Specific_Name")(ChargeBox.List(i, 1))
                    Range(headerFind("Charge Grade (broad) #" & j, tempHead) & updateRow).value _
                            = calcChargeBroad(ChargeBox.List(i, 1))
                    j = j + 1
                End If
            End If
        Next i
    Next Num



    'ZERO FILL

    'Dim counter As Long


    'For counter = (alphaToNum(headerFind("PETITION")) + 1) _
    '    To (alphaToNum(headerFind("DRAI")) - 1)
    '
    '    If Cells(updateRow, counter).value = "" Then
    '        Cells(updateRow, counter).value = 0
    '    End If
    'Next counter

    'For counter = (alphaToNum(headerFind("DEMOGRAPHICS")) + 1) _
    '    To (alphaToNum(headerFind("PETITION")) - 1)
    '
    '    If Cells(updateRow, counter).value = "" Then
    '        Cells(updateRow, counter).value = 0
    '    End If
    'Next counter
    'Call aggFlag(updateRow)
done:
    'Call SaveAs_Countdown
    Call Save_Countdown
    Call UnloadAll

    Worksheets("User Entry").Activate
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
    End With
    Exit Sub
err:

    Range("C" & updateRow & ":" & headerFind("END") & updateRow).value = restorer

    Stop 'press F8 twice to see the error point
    Resume
    MsgBox "Something went wrong. Database has been restored to state prior to submission. " _
      & vbNewLine & vbNewLine & "Message: " & vbNewLine & err.Description _
      & vbNewLine & vbNewLine & "Source: " & vbNewLine & err.Source

    Call UnloadAll
End Sub





Private Sub TimeReferredToDA_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    TimeReferredToDA.value = VBA.format(TimeReferredToDA.value, "hh:mm AM/PM")
End Sub




Private Sub TestFill1_Click()
    IncidentDate.value = "02/01/2019"
    TimeOfIncident_H.value = "04"
    TimeOfIncident_M.value = "00"
    TimeOfIncident_P.value = "PM"
    IncidentDistrict.value = "22"
    IncidentAddress.value = "1501 Market St."
    IncidentZipcode.value = "19102"

    ArrestDate.value = "02/01/2019"
    TimeOfArrest_H.value = "02"
    TimeOfArrest_M.value = "30"
    TimeOfArrest_P.value = "PM"
    TimeReferredToDA_H.value = "10"
    TimeReferredToDA_M.value = "00"
    TimeReferredToDA_P.value = "AM"
    ArrestingDistrict.value = "22"

    ActiveAtArrest.value = "No"
    NumOfPriorArrests.value = 0
    DCNum.value = "12345"
    PIDNum.value = "5467"
    SIDNum.value = "87980"

    Officer1.value = "AO1Test"
    Officer2.value = "AO2Test"
    Officer3.value = "AO3Test"
    Officer4.value = "AO4Test"
    Officer4.value = "AO5Test"

    VictimFirstName = "VicFirstTest"
    VictimLastName = "VicLastTest"

    'Petitions

    With PetitionBox
        .ColumnCount = 7
        .ColumnWidths = "50;50;30;50;65;50;0"
        .AddItem "07/01/2017"
        .List(0, 1) = "13579"
        .List(0, 2) = "F"
        .List(0, 3) = "Assaults"
        .List(0, 4) = "18 - 2702"
        .List(0, 5) = "AGGRAVATED ASSAULT"
        .List(0, 6) = "No"
    End With

End Sub
