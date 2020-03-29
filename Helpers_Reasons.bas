Attribute VB_Name = "Helpers_Reasons"
Function encodeReasons(re1 As String, re2 As String, re3 As String, re4 As String, re5 As String) As String
    encodeReasons = re1 + "*" + re2 + "*" + re3 + "*" + re4 + "*" + re5
End Function

Function decodeReasons(encoded As String) As String()
    decodeReasons = Split(encoded, "*")
End Function

