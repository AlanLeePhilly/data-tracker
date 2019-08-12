VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Adult_Reslate_Juvenile_Petition 
   Caption         =   "UserForm1"
   ClientHeight    =   7575
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18990
   OleObjectBlob   =   "Adult_Reslate_Juvenile_Petition.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Adult_Reslate_Juvenile_Petition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public NCF_userRow As Long
Public NCF_rearrestNum As Long

Private Sub ConfOutcome_Change()
    Select Case ConfOutcome.value
        Case "Release for Diversion"
            DiversionProgram.value = "Yes"
        Case Else
            DiversionProgram.value = "No"
    End Select
End Sub

Private Sub DRAI_Score_Change()
    If IsNumeric(DRAI_Score.value) Then
        Select Case DRAI_Score.value
            Case Is < 10
                DRAI_Rec.value = "Release"
            Case Is < 15
                DRAI_Rec.value = "Release w/ Supervision"
            Case Is >= 15
                DRAI_Rec.value = "Hold"
            Case Else
                DRAI_Rec.value = "Unknown"
        End Select
    End If
End Sub
Private Sub DRAI_Action_Change()
    Select Case DRAI_Action.value
        Case "Follow - Hold", "Override - Hold"
            NextHearingLocation = "PJJSC"
            DetentionFacility.Enabled = True
            DetentionFacilityLabel.Enabled = True

            Supv1.Enabled = False
            Supv1Pro.Enabled = False
            Supv1Re1.Enabled = False
            Supv1Re2.Enabled = False
            Supv1Re3.Enabled = False

            Supv2.Enabled = False
            Supv2Pro.Enabled = False
            Supv2Re1.Enabled = False
            Supv2Re2.Enabled = False
            Supv2Re3.Enabled = False

            Cond1.Enabled = False
            Cond1Pro.Enabled = False
            Cond2.Enabled = False
            Cond2Pro.Enabled = False
            Cond3.Enabled = False
            Cond3Pro.Enabled = False

        Case Else
            Supv1.Enabled = True
            Supv1Pro.Enabled = True
            Supv1Re1.Enabled = True
            Supv1Re2.Enabled = True
            Supv1Re3.Enabled = True

            Supv2.Enabled = True
            Supv2Pro.Enabled = True
            Supv2Re1.Enabled = True
            Supv2Re2.Enabled = True
            Supv2Re3.Enabled = True

            Cond1.Enabled = True
            Cond1Pro.Enabled = True
            Cond2.Enabled = True
            Cond2Pro.Enabled = True
            Cond3.Enabled = True
            Cond3Pro.Enabled = True

    End Select
End Sub



Private Sub InConfDate_Enter()
    InConfDate.value = CalendarForm.GetDate(RangeOfYears:=5)
End Sub
Private Sub InConfDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Set ctl = Me.InConfDate
    Call DateValidation(ctl, Cancel)
End Sub
Private Sub CallInDate_Enter()
    CallInDate.value = CalendarForm.GetDate(RangeOfYears:=5)
End Sub
Private Sub CallInDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Set ctl = Me.CallInDate
    Call DateValidation(ctl, Cancel)
End Sub

Private Sub NextHearingLocation_Change()
    If NextHearingLocation.value = "Intake Conf." Then
        MsgBox "Not a valid value for this prompt"
        NextHearingLocation.value = "N/A"
        Exit Sub
    End If
End Sub


Private Sub UserForm_Initialize()
    Me.ScrollTop = 0
    DetentionFacilityLabel.Enabled = False
    DetentionFacility.Enabled = False
End Sub

Private Sub AddPetition_Click()
    If PetitionBox.ListCount < 5 Then
        Load Modal_NewClient_Add_Petition
        Modal_NewClient_Add_Petition.headline.Caption = "Reslate"
        Modal_NewClient_Add_Petition.Show
    Else
        MsgBox "Maximum of five petitions for reslate"
    End If
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


Private Sub DiversionProgramReferralDate_Enter()
    DiversionProgramReferralDate.value = CalendarForm.GetDate(RangeOfYears:=5)
End Sub
Private Sub DiversionProgramReferralDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Set ctl = Me.DiversionProgramReferralDate
    Call DateValidation(ctl, Cancel)
End Sub

Private Sub Cancel_Click()
    Call Clear_Click
    Adult_Reslate_Juvenile_Petition.Hide
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
    Call UserForm_Initialize

End Sub

Private Sub DiversionProgram_Change()
    Select Case DiversionProgram.value
        Case "No"

            DiversionProgramReferralDateLabel.Enabled = False
            DiversionProgramReferralDate.Enabled = False
            DiversionProgramReferralDate.value = ""

            ReferralSource.Enabled = False
            ReferralSource.value = "N/A"
            ReferralSourceLabel.Enabled = False

            NameOfProgramLabel.Enabled = False
            NameOfProgram.Enabled = False
            NameOfProgram.value = "N/A"

            YAPDistrictLabel.Enabled = False
            YAPDistrict.Enabled = False
            YAPDistrict.value = ""

            NoDiversionReason1.Enabled = True
            NoDiversionReason2.Enabled = True
            NoDiversionReason3.Enabled = True
        Case Else

            DiversionProgramReferralDateLabel.Enabled = True
            DiversionProgramReferralDate.Enabled = True
            DiversionProgramReferralDate.value = InitialHearingDate.value

            ReferralSource.Enabled = True
            ReferralSourceLabel.Enabled = True

            NameOfProgramLabel.Enabled = True
            NameOfProgram.Enabled = True

            YAPDistrictLabel.Enabled = True
            YAPDistrict.Enabled = True

            NoDiversionReason1.Enabled = False
            NoDiversionReason2.Enabled = False
            NoDiversionReason3.Enabled = False
            NoDiversionReason1.value = "N/A"
            NoDiversionReason2.value = "N/A"
            NoDiversionReason3.value = "N/A"
    End Select
End Sub

Private Sub NameOfProgram_Change()
    If NameOfProgram = "YAP" Then
        YAPDistrictLabel.Enabled = True
        YAPDistrict.Enabled = True
    Else
        YAPDistrictLabel.Enabled = False
        YAPDistrict.Enabled = False
        YAPDistrict.value = ""
    End If
End Sub

Private Sub Submit_Click()

    '''''''''''''
    'Validations'
    '''''''''''''

    If InConfDate.value = "" And InConfRecord.value = "Yes" Then
        MsgBox "Intake Date Required if record available"
        Exit Sub
    End If

    If ConfOutcome.value = "N/A" And InConfRecord.value = "Yes" Then
        MsgBox "Conference Outcome Required if record available"
        Exit Sub
    End If

    If CallInDate.value = "" And CallInRecord.value = "Yes" Then
        MsgBox "Call-in Date required if record available"
        Exit Sub
    End If

    If PetitionBox.ListCount = 0 Then
        MsgBox "Petition required"
        Exit Sub
    End If

    If GunCase.value = "" Then
        MsgBox "'Gun Case?' required"
        Exit Sub
    End If

    If GunInvolved.value = "" Then
        MsgBox "'Gun Involved?' required"
        Exit Sub
    End If

    If DRAI_Action.value = "Follow - Hold" Or DRAI_Action.value = "Override - Hold" Then
        If DetentionFacility.value = "N/A" Then
            MsgBox "Detention facility required for call-in hold"
            Exit Sub
        End If
    End If

    If DiversionProgram.value = "No" And NoDiversionReason1.value = "N/A" Then
        MsgBox "Reason Not Diverted Required"
        Exit Sub
    End If
    
    If NextHearingLocation = "N/A" Or NextHearingLocation = "Adult" Then
        MsgBox "New Hearing Location Required"
        Exit Sub
    End If


    Adult_Reslate_Juvenile_Petition.Hide
    ClientUpdateForm.Adult_Reslate_Update.BackColor = selectedColor
    ClientUpdateForm.Adult_Reslate_Remain.BackColor = unselectedColor
    Exit Sub
err:

    Range("C" & emptyRow & ":" & headerFind("END") & emptyRow).value = restorer

    Stop 'press F8 twice to see the error point
    Resume
    MsgBox "Something went wrong. Database has been restored to state prior to submission. " _
      & vbNewLine & vbNewLine & "Message: " & vbNewLine & err.Description _
      & vbNewLine & vbNewLine & "Source: " & vbNewLine & err.Source

    Call UnloadAll
End Sub

Private Sub TestFillPetition_Click()
    With PetitionBox
        .ColumnCount = 7
        .ColumnWidths = "50;50;30;50;65;50;0"
        .AddItem "01/01/2019"
        .List(0, 1) = "13579"
        .List(0, 2) = "F"
        .List(0, 3) = "Assaults"
        .List(0, 4) = "18 - 2702"
        .List(0, 5) = "AGGRAVATED ASSAULT"
        .List(0, 6) = "No"
    End With

    CallInDate.value = "02/01/2019"
    Was_DRAI_Administered.value = "Yes"
    DRAI_Score.value = "4"
    DRAI_Rec.value = "Release"
    DRAI_Action.value = "Follow - Release"

    InConfDate.value = "02/01/2019"
    ConfOutcome.value = "Release for Court"

    NoDiversionReason1 = "Charge Ineligible"

    NextHearingLocation.value = "3E"

    InitialHearingDate = "2/12/2019"


End Sub

Private Sub TestFillDiversion_Click()
    'Petitions

    With PetitionBox
        .ColumnCount = 7
        .ColumnWidths = "50;50;30;50;65;50;0"
        .AddItem "08/01/2018"
        .List(0, 1) = "13579"
        .List(0, 2) = "F"
        .List(0, 3) = "Assaults"
        .List(0, 4) = "18 - 2702"
        .List(0, 5) = "AGGRAVATED ASSAULT"
        .List(0, 6) = "No"
    End With

    CallInDate.value = "02/01/2019"
    Was_DRAI_Administered.value = "Yes"
    DRAI_Score.value = "4"
    DRAI_Rec.value = "Release"
    DRAI_Action.value = "Follow - Release"

    InConfDate.value = "02/01/2019"
    ConfOutcome.value = "Release for Diversion"

    DiversionProgram.value = "Yes"
    DiversionProgramReferralDate.value = "2/1/19"
    ReferralSource.value = "Pre-Petition DA"
    NameOfProgram.value = "YAP"
    YAPDistrict.value = 2
    
    InitialHearingDate.value = "02/01/2019"
    NextHearingLocation.value = "Diversion"
    ListingType.value = "Diversion"

End Sub

Private Sub TestFillIntake_Click()
    With PetitionBox
        .ColumnCount = 7
        .ColumnWidths = "50;50;30;50;65;50;0"
        .AddItem "09/08/2018"
        .List(0, 1) = "13579"
        .List(0, 2) = "F"
        .List(0, 3) = "Assaults"
        .List(0, 4) = "18 - 2702"
        .List(0, 5) = "AGGRAVATED ASSAULT"
        .List(0, 6) = "No"
    End With

    InConfRecord.value = "Yes"
    InConfType.value = "DA"
    Was_DRAI_Administered = "Yes"
    DRAI_Score.value = 25
    DRAI_Rec.value = "Hold"
    DRAI_Action = "Follow - Hold"
    OverrideHoldRe1.value = "B/W"
    OverrideHoldRe2.value = "Drug Screens"
    OverrideHoldRe3.value = "N/A"
    ConfOutcome.value = "Release for Court"

    InitialHearingDate.value = "02/01/2019"
    NextHearingLocation.value = "3E"

    DiversionProgram.value = "No"

End Sub




