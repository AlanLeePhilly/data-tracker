Attribute VB_Name = "Calculate_Addresses"
Sub FillCoordinates(addr As String, zip As String, latCol As String, lonCol As String)
    'Opening data set
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    ActiveWindow.WindowState = xlMinimized
    Sheets("Entry").Activate
    
    'Go through all entries with addresses and find latitudes and longitudes
    'Only calculate location if the address and zipcode are both present
    Dim i As Integer
    For i = 3 To 2500
        If Not IsEmpty(Cells(i, addr)) And Not IsEmpty(Cells(i, zip)) Then
            'Finding address lat and lon
            Dim coords As Variant
            Dim currZip As String
            currZip = Cells(i, zip).value
            coords = FindLatLon(Cells(i, addr), Cells(i, zip), i)
            Cells(i, latCol).value = coords(1)
            Cells(i, lonCol).value = coords(2)
            Cells(i, zip).value = coords(3)
            
            If Not StrComp(currZip, coords(3)) = 0 Then
                MsgBox ("Row " & i & " || ALERT: The zipcode entered and zipcode found by geolocating services are different. Please check the new zipcode entered to make sure it is correct.")
            End If
        End If
    Next i
End Sub

Sub FillAddressCoordinates()
    'Opening data set
    ThisWorkbook.Activate
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    ActiveWindow.WindowState = xlMinimized
    Sheets("Entry").Activate
    
    'Find columns to use to look up and fill coordinates

    Dim addr As String
    Dim zip As String
    addr = headerFind("Address")
    zip = headerFind("Zipcode")
    
    Dim latCol As String
    latCol = headerFind("Latitude")
    Dim lonCol As String
    lonCol = headerFind("Longitude")
    
    FillCoordinates addr, zip, latCol, lonCol

    MsgBox ("All addresses (re)calculated")
End Sub

Sub FillIncidentAddressCoordinates()
    'Opening data set
    ThisWorkbook.Activate
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    ActiveWindow.WindowState = xlMinimized
    Sheets("Entry").Activate
    
    'Find columns to use to look up and fill coordinates
    
    Dim petitionHead As String
    petitionHead = hFind("PETITION")

    Dim addr As String
    Dim zip As String
    addr = headerFind("Incident Address", petitionHead)
    zip = headerFind("Incident Zipcode", petitionHead)
    
    Dim latCol As String
    latCol = headerFind("Latitude", petitionHead)
    Dim lonCol As String
    lonCol = headerFind("Longitude", petitionHead)
    
    FillCoordinates addr, zip, latCol, lonCol

    MsgBox ("All addresses (re)calculated")
End Sub
