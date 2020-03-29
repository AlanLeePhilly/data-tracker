Attribute VB_Name = "Module10_Cache"
Sub launchCacheLoader()
    CacheLoader.Show
End Sub

Sub cacheRow(dataRow As Long)
    Dim Cache As Worksheet
    Dim Data As Worksheet
    Dim userData As Range
    
    If dataRow < 3 Then
        MsgBox "Debug: CacheRow attempting to cache row " & dataRow
        Exit Sub
    End If
    
    Set Cache = Worksheets("Cache")
    Set Data = Worksheets("Entry")
    Set userData = Data.Range("C" & dataRow & ":" & hFind("END") & dataRow)
    
    Cache.Rows(2).Insert (xlShiftDown)
    Cache.Range("A2").value = dataRow
    Cache.Range("B2").value = Data.Range(hFind("First Name") & dataRow).value & " " _
        & Data.Range(hFind("Last Name") & dataRow).value
    Cache.Range("C2").value = Now()
    Cache.Range("D2").Resize(columnSize:=userData.Columns.count).value = userData.value
End Sub

Sub clearCache()
    Dim Cache As Worksheet
    
    Set Cache = Worksheets("Cache")
    Cache.UsedRange.Offset(1).ClearContents
End Sub

Sub loadFromCache(cacheRow As Long)
    Dim Cache As Worksheet
    Dim Data As Worksheet
    Dim cachedData As Range
    Dim userData As Range
    Dim dataRow As Long
    Dim name As String
    Dim timeStamp As Date
    
    Set Cache = Worksheets("Cache")
    Set Data = Worksheets("Entry")
    Set cachedData = Cache.Range("D" & cacheRow, hFind("END") & cacheRow)
    dataRow = Cache.Range("A" & cacheRow).value
    name = Cache.Range("B" & cacheRow).value
    timeStamp = Cache.Range("C" & cacheRow).value
    
    If dataRow = 0 Then
        MsgBox "No record selected"
        Exit Sub
    End If
    
    If dataRow = 2 Then
        MsgBox "You were about to overwrite one of the first two rows... something is very wrong"
        Exit Sub
    End If
    
    Set userData = Data.Range("C" & dataRow).Resize(columnSize:=cachedData.Columns.count)
    userData.ClearContents
    userData.value = cachedData.value
    Cache.Range("A" & cacheRow).EntireRow.Delete
    
    MsgBox _
        "Record restored" & vbNewLine & _
        "Client name: " & name & vbNewLine & _
        "Row: " & dataRow & vbNewLine & _
        "Cached data timestamp: " & timeStamp
End Sub


