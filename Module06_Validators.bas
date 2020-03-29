Attribute VB_Name = "Module06_Validators"
'generic sub for date validation, which takes in a textbox as a an argument
'usually called upon when user exits focus from a textbox which requires date formatting
Public Sub DateValidation(ByRef DateBox As Control, ByRef Cancel As MSForms.ReturnBoolean)
    Dim splitDate() As String: splitDate = Split(DateBox.Text, "/")

    'if the textbox value is blank (Trim function removes leading spaces if present,
    If Trim(DateBox.value) = "" Then
        'set textbox color to regular color and exit sub
        DateBox.BackColor = &H80000005
        Exit Sub

        'else if it's not a date
    Else
        If IsDate(DateBox.Text) Then
            If Len(splitDate(0)) < 3 Then
                If Len(splitDate(1)) < 3 Then
                    If Len(splitDate(2)) > 1 And Len(splitDate(2)) < 5 Then
                        'set textbox color to regular color and exit sub
                        DateBox.BackColor = &H80000005
                        Exit Sub
                    End If
                End If
            End If
        End If
    End If

    'set textbox color to red, throw error pop-up
    DateBox.BackColor = &HFF& 'change the color of the textbox to red
    MsgBox "Illegal date value" + vbNewLine _
        + "Use format mm/dd/yyyy"
    ' setting Cancel to True means the user cannot cannot complete whatever action triggered the validation
    ' in most cases for this form, it will mean that the user can't
    Cancel = True
End Sub

Public Function HasContent(text_box As Object) As Boolean
    HasContent = (Len(Trim(text_box.value)) > 0)
End Function


Sub Clear_Form(myForm As UserForm)

    'for each control (generic name for any field in form
    For Each ctl In myForm.Controls

        'determine the type of control it is and reset value accordingly
        Select Case TypeName(ctl)
            Case "TextBox"
                ctl.value = ""
            Case "CheckBox", "ToggleButton"
                ctl.value = False
            Case "OptionGroup"
                ctl = Null
            Case "OptionButton"
                ' Do not reset an optionbutton if it is part of an OptionGroup
                If TypeName(ctl.Parent) <> "OptionGroup" Then ctl.value = False
            Case "ComboBox"
                ctl.listIndex = -1
            Case "ListBox"
                ctl.Clear
        End Select
    Next ctl
End Sub
