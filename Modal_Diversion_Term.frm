VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Modal_Diversion_Term 
   Caption         =   "Term Edit"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7020
   OleObjectBlob   =   "Modal_Diversion_Term.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Modal_Diversion_Term"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cancel_Click()
    Clear_Form (Modal_Diversion_Term)
    Modal_Diversion_Term.Hide
End Sub

Private Sub Continue_Click()
    Dim i As Integer

    For i = 0 To DiversionUpdateForm.ReturnTerms.ListCount - 1
        If DiversionUpdateForm.ReturnTerms.List(i, 0) = EditTerms.List(EditTerms.listIndex, 0) Then
            DiversionUpdateForm.ReturnTerms.List(i, 1) = NewTerm
            DiversionUpdateForm.ReturnTerms.List(i, 2) = NewTermProvider
        End If
    Next i

    Modal_Diversion_Term.Hide
End Sub

Private Sub EditTerms_Click()
    If EditTerms.listIndex >= 0 Then
        Dim i As Integer

        For i = 0 To EditTerms.ListCount - 1
            NewTerm = EditTerms.List(EditTerms.listIndex, 1)
            NewTermProvider = EditTerms.List(EditTerms.listIndex, 2) '
        Next i
    End If
End Sub

