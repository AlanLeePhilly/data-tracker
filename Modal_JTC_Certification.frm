VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Modal_JTC_Certification 
   Caption         =   "JTC Certification"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5910
   OleObjectBlob   =   "Modal_JTC_Certification.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Modal_JTC_Certification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub UserForm_Initialize()
    Prior_Status.Caption = ClientUpdateForm.JTC_Fetch_Certification
    If ClientUpdateForm.JTC_Fetch_Certification = "Filed" Then
        MultiPage1.value = 1
    Else
        MultiPage1.value = 0
    End If
End Sub
                        
                        
                        '''''''''''''
                        'VALIDATIONS'
                        '''''''''''''
Private Sub Motion_Date_Enter()
    Motion_Date.value = CalendarForm.GetDate(RangeOfYears:=5)
End Sub
Private Sub Motion_Date_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Set ctl = Modal_JTC_Certification.Motion_Date

    Call DateValidation(ctl, Cancel)
End Sub

                        ''''''''''''''''''
                        '''''BUTTONS''''''
                        ''''''''''''''''''
    
Private Sub InsertDoH_Click()
    Motion_Date = ClientUpdateForm.DateOfHearing
End Sub

Private Sub Cancel_Click()
    Unload Modal_JTC_Certification
End Sub
                         
                        '''''''''''''''''''''''
                        '''''SUBMIT LOGIC''''''
                        '''''''''''''''''''''''

Private Sub Continue_Click()
    'VALIDATIONS
    If Prior_Status.Caption = "Filed" Then
        If Motion_Result.value = "N/A" Then
            MsgBox "Result of Motion Required"
            Exit Sub
        End If
    Else
        If Not HasContent(Motion_Date) Then
            MsgBox "Date of Motion Required"
            Exit Sub
        End If
        If Not HasContent(Was_Motion_Filed) Then
            MsgBox "'Was Motion Filed?' Required"
            Exit Sub
        End If
    End If
    
    
    
    ClientUpdateForm.JTC_Certification_Update.BackColor = selectedColor
    ClientUpdateForm.JTC_Certification_Remain.BackColor = unselectedColor
    
    If Prior_Status.Caption = "Filed" Then
        ClientUpdateForm.JTC_Return_Certification.Caption = Motion_Result
    Else
        If Was_Motion_Filed = "Yes" Then
            ClientUpdateForm.JTC_Return_Certification.Caption = "Filed"
        Else
            ClientUpdateForm.JTC_Return_Certification.Caption = "None"
        End If
    End If
    
    Modal_JTC_Certification.Hide
End Sub



