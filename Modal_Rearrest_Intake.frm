VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Modal_Rearrest_Intake 
   Caption         =   "UserForm1"
   ClientHeight    =   3660
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   5484
   OleObjectBlob   =   "Modal_Rearrest_Intake.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Modal_Rearrest_Intake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Continue_Click()
    If IsNull(Rearrests.value) Then
        MsgBox "Must select an arrest"
        Exit Sub
    Else
        Call RearrestIntake(updateRow, Rearrests.value)
    End If
End Sub


Private Sub UserForm_Initialize()
    With Modal_Rearrest_Intake.Rearrests
        Dim i As Integer
        Dim bucketHead
        For i = 1 To 5
            If isNotEmptyOrZero(Range(hFind("Arrest Date #" & i, "REARRESTS", "AGGREGATES") & updateRow)) Then
                bucketHead = hFind("Arrest Date #" & i, "REARRESTS", "AGGREGATES")

                .ColumnCount = 5
                .ColumnWidths = "0;75;150;20;20;"
                .AddItem i
                .List(.ListCount - 1, 0) = i
                .List(.ListCount - 1, 1) = Range(bucketHead & updateRow).value
                .List(.ListCount - 1, 2) = Range(headerFind("Lead Charge Name", bucketHead) & updateRow).value
            End If
        Next i
    End With
End Sub
