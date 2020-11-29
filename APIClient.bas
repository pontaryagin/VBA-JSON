Attribute VB_Name = "APIClient"
Option Explicit

Public Function FetchJsonAPI(ByVal request As String, ByVal url As String, Optional ByVal param As Object) As Object
    Dim json
    json = ConvertToJson(param)

    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    Dim res As Object: Set res = Nothing
    http.Open request, url, False
    http.SetRequestHeader "Content-Type", "application/json; charset=UTF-8"
    http.send json

    If http.ResponseText <> "" Then
        Set res = ParseJson(http.ResponseText)
    End If
    Set FetchJsonAPI = res
End Function
