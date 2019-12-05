VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Log_Payment 
   Caption         =   "Filed Payment & Hours"
   ClientHeight    =   6405
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   12540
   OleObjectBlob   =   "Log_Payment.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Log_Payment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Filing_Type_Change()
    Select Case Filing_Type.value
        Case "Restitution"
            Amount_Label.Caption = "Amount:  $"
        Case "Court Costs"
            Amount_Label.Caption = "Amount:  $"
        Case "Community Service"
            Amount_Label.Caption = "Amount:   "
    
    End Select
End Sub

Private Sub InsertDoH_Click()
    DateOf.value = ClientUpdateForm.DateOfHearing.value
End Sub


Private Sub Submit_Click()
    If Filing.value = False And Payment.value = False Then
        MsgBox "Must select Filing or Payment"
        Exit Sub
    End If
    
    If Amount.value = "" Then
        MsgBox "Must provide amount."
        MsgBox updateRow
        Exit Sub
    End If
    
    Select Case True
        Case Filing_Type.value = "Restitution" And Filing.value = True
            Call updateRestitution( _
                Courtroom:=ClientUpdateForm.Courtroom.value, _
                DA:=ClientUpdateForm.DA.value, _
                userRow:=updateRow, _
                DateOf:=DateOf.value, _
                amountFiled:=Amount.value)
        Case Filing_Type.value = "Restitution" And Payment.value = True
            Call updateRestitution( _
                Courtroom:=ClientUpdateForm.Courtroom.value, _
                DA:=ClientUpdateForm.DA.value, _
                userRow:=updateRow, _
                DateOf:=DateOf.value, _
                amountPaid:=Amount.value)
        Case Filing_Type.value = "Court Cost" And Filing.value = True
            Call updateCourtCost( _
                Courtroom:=ClientUpdateForm.Courtroom.value, _
                DA:=ClientUpdateForm.DA.value, _
                userRow:=updateRow, _
                DateOf:=DateOf.value, _
                amountFiled:=Amount.value)
        Case Filing_Type.value = "Court Cost" And Payment.value = True
            Call updateCourtCost( _
                Courtroom:=ClientUpdateForm.Courtroom.value, _
                DA:=ClientUpdateForm.DA.value, _
                userRow:=updateRow, _
                DateOf:=DateOf.value, _
                amountPaid:=Amount.value)
    End Select
    
    Unload Me
    
End Sub


Private Sub UserForm_Initialize()
    Call fetchFiledRecord(updateRow)
    With ClientUpdateForm.SearchResultsBox
        Name_Display.Caption = .List(.listIndex, 1) + " " + .List(.listIndex, 2)
    
    End With
    
End Sub
