Attribute VB_Name = "JTC_Buttons"
'''''''''''''''''''
'JTC_PHASE_UPDATES'
'''''''''''''''''''
Sub JTC_Accept_Click()
    Modal_JTC_Accept.Show
End Sub
Sub JTC_Reject_Click()
    Modal_JTC_Reject.Show
End Sub
Sub JTC_Phase_Stepup_Click()
    Modal_JTC_Stepup.Show
End Sub
Sub JTC_Phase_Pushback_Click()
    Modal_JTC_Phase_Pushback.Show
End Sub
Sub JTC_Discharge_Click()
    Modal_JTC_Discharge.Show
End Sub
Sub JTC_Expungement_Click()
    Modal_JTC_Expungement.Show
End Sub

Sub JTC_Phase_Remain_Click()
    'take displayed phase and stepup
    JTC_Return_Phase.Caption = JTC_Fetch_Phase.Caption
    JTC_Return_Stepup_Date.Caption = JTC_Fetch_Stepup_Date.Caption

    'hide pushback display
    JTC_Pushback_Reason_Label.Visible = False
    JTC_Pushback_Reason1.Caption = ""
    JTC_Pushback_Reason2.Caption = ""
    JTC_Pushback_Reason3.Caption = ""

    'color buttons
    JTC_Phase_Remain.BackColor = &H8000000A
    JTC_Phase_Stepup.BackColor = &H8000000F
    JTC_Phase_Pushback.BackColor = &H8000000F
    JTC_Discharge.BackColor = &H8000000F
End Sub

'''''''''''''''''''''''
'JTC_TREATMENT_UPDATES'
'''''''''''''''''''''''

Sub JTC_Treatment_Stepdown_Click()
    Modal_JTC_Stepdown.Show
End Sub

Sub JTC_Treatment_Provider_Update_Click()
    Modal_JTC_Provider.Show
End Sub

Sub JTC_Treatment_Provider_Remain_Click()
    'push provider name
    JTC_Return_Treatment_Provider.Caption = JTC_Fetch_Treatment_Provider.Caption

    'hide stepdown display
    JTC_Return_Stepdown_Date.Caption = ""
    ClientUpdateForm.JTC_Stepdown_Label.Visible = False

    'color buttons
    JTC_Treatment_Provider_Remain.BackColor = &H8000000A
    JTC_Treatment_Stepdown.BackColor = &H8000000F
    JTC_Treatment_Provider_Update.BackColor = &H8000000F
    JTC_Treatment_Discharge.BackColor = &H8000000F
End Sub

Sub JTC_Treatment_Discharge_Click()
    JTC_Treatment_Provider_Remain.BackColor = &H8000000F
    JTC_Treatment_Stepdown.BackColor = &H8000000F
    JTC_Treatment_Provider_Update.BackColor = &H8000000F
    JTC_Treatment_Discharge.BackColor = &H8000000A
End Sub

Sub JTC_Service_Add_Click()
    Modal_JTC_Add_Service.Show
End Sub

Sub JTC_Service_Discharge_Click()
    Modal_JTC_Drop_Service.Show
End Sub

Sub JTC_Service_Remain_Click()
    If JTC_Service_Remain.BackColor = &H8000000F Then
        JTC_Service_Remain.BackColor = &H8000000A
    Else
        JTC_Service_Remain.BackColor = &H8000000F
    End If
End Sub
