Attribute VB_Name = "R_Convert"
Sub DeleteControls()
'
' deletecontrols Macro
'
    Worksheets("Entry").Activate
    ActiveSheet.Shapes.Range(Array("Button 3071")).Select
    ActiveSheet.Shapes.SelectAll
    Selection.Delete
    

End Sub
