VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Modal_Client_Edit 
   Caption         =   "Edit Record"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5415
   OleObjectBlob   =   "Modal_Client_Edit.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Modal_Client_Edit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cancel_Click()
    Unload Me
End Sub

Private Sub Submit_Click()
    With ClientEdit.UpdateBox
        .ColumnCount = 4
        .ColumnWidths = "0;20;130;130;"
        .AddItem count
        .List(ClientEdit.UpdateBox.ListCount - 1, 1) _
                    = ClientEdit.LookupBox.List(ClientEdit.LookupBox.listIndex, 1)
        .List(ClientEdit.UpdateBox.ListCount - 1, 2) _
                    = ClientEdit.LookupBox.List(ClientEdit.LookupBox.listIndex, 2)

        If Modal_Client_Edit.New_Value_Text.Visible = True Then
            .List(ClientEdit.UpdateBox.ListCount - 1, 3) _
                        = New_Value_Text
        Else
            .List(ClientEdit.UpdateBox.ListCount - 1, 3) _
                        = New_Value_Box
        End If
    End With

    Unload Me
End Sub

