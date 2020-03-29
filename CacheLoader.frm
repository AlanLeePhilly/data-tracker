VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CacheLoader 
   Caption         =   "CacheLoader"
   ClientHeight    =   4860
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   8580.001
   OleObjectBlob   =   "CacheLoader.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CacheLoader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub LoadButton_Click()
    
    Call loadFromCache(ListBox1.value)
    
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    Dim Cache As Worksheet
    Dim i As Integer
    Dim newIndex As Integer
    
    Set Cache = Worksheets("Cache")
    
    i = 2
    
    With ListBox1
        .Clear
        .ColumnCount = 4
        .ColumnWidths = "0;30;90;75"
        
        
        While isNotEmptyOrZero(Cache.Range("A" & i))
            .AddItem i
                newIndex = .ListCount - 1
                .List(newIndex, 1) = Cache.Range("A" & i).value
                .List(newIndex, 2) = Cache.Range("B" & i).value
                .List(newIndex, 3) = Cache.Range("C" & i).value
        
            i = i + 1
        Wend
    End With
End Sub



