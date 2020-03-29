Attribute VB_Name = "Geocoding"
Public Function FindLatLon(Address As String, Zipcode As String, rowNum As Integer)
    Dim hReq As Object
    Dim Json As Object
    Dim try As Object
    Dim strUrl As String
    Dim addressStr As String
    Dim cityStr As String

    On Error GoTo LatLonErr

    cityStr = ""

    'Check to see if address is "Homeless" or if zipcode is "19100"
    If StrComp(UCase(Address), "HOMELESS") = 0 Then
        MsgBox ("Row " & rowNum & " || ALERT: Address of 'homeless' entered was not mappable; no latitude or longitude coordinates added")
        Dim responseArr() As Variant
        responseArr = Array("", "", "", Zipcode)
        FindLatLon = responseArr
        Exit Function
    End If

    If StrComp(Zipcode, "19100") = 0 Then
        MsgBox ("Row " & rowNum & " || ALERT: Address entered with zipcode '19100'; city 'Philadelphia' used with no zipcode to attempt mapping instead")
        cityStr = "Philadelphia"
        Zipcode = ""
    End If

    'Probably want try-catch here
    strUrl = "https://nominatim.openstreetmap.org/search?format=json&addressdetails=1&limit=1&q=" & WorksheetFunction.EncodeURL(Address) & "%20" & WorksheetFunction.EncodeURL(cityStr) & "%20" & WorksheetFunction.EncodeURL(Zipcode)

    Set hReq = CreateObject("MSXML2.XMLHTTP")
    With hReq
        .Open "GET", strUrl, False
        .Send
    End With

    addressStr = "[{" & Mid(hReq.ResponseText, 107, Len(hReq.ResponseText))

    Set Json = JsonConverter.ParseJson(addressStr)

    Dim a(1 To 3) As Double
    a(1) = Json(1)("lat")
    a(2) = Json(1)("lon")
    a(3) = Json(1)("address")("postcode")
    FindLatLon = a
    Exit Function

LatLonErr:
    MsgBox ("Row " & rowNum & " || ALERT: Error occurred in finding location coordinates: " & err.Description & "; Setting coordinates to null. Please check address and zipcode and edit if location coordinates desired.")
    Dim errorArr() As Variant
    errorArr = Array("", "", "", Zipcode)
    FindLatLon = errorArr
    Exit Function

End Function

