VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Concurrency 
   Caption         =   "Concurrency"
   ClientHeight    =   5100
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   7632
   OleObjectBlob   =   "Concurrency.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Concurrency"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub NewUpdate_Click()
    Dim userRow As Long

    If Not RowBox.value = "" Then
        userRow = RowBox.value

        With ClientUpdateForm.SearchResultsBox
            .AddItem userRow
            .List(.ListCount - 1, 1) = Range(headerFind("First Name") & userRow)
            .List(.ListCount - 1, 2) = Range(headerFind("Last Name") & userRow)
            .List(.ListCount - 1, 3) = Range(headerFind("DOB") & userRow)
            .List(.ListCount - 1, 4) = Range(headerFind("Arrest Date") & userRow)
            .List(.ListCount - 1, 5) = Range(headerFind("Petition #1") & userRow)
            .List(.ListCount - 1, 6) = Lookup("Courtroom_Num")(Range(headerFind("Active Courtroom") & userRow).value)
            .List(.ListCount - 1, 7) = Lookup("Legal_Status_Num")(Range(headerFind("Legal Status") & userRow).value)
            .List(.ListCount - 1, 8) = Lookup("Supervision_Program_Num")(Range(headerFind("Active Supervision") & userRow).value)
            .listIndex = 0
        End With

        Call ClientUpdateForm.SearchResultsBox_Click
        ClientUpdateForm.DateOfHearing.value = RowBox.List(0, 5)
        ClientUpdateForm.Show
    End If
End Sub

Private Sub Quit_Click()
    Unload Me
End Sub
