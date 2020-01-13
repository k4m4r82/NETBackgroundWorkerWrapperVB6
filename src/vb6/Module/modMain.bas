Attribute VB_Name = "modMain"
Option Explicit

Public Const API_URL As String = "http://api.coding4ever.net:5000/api/buku"

Public Sub Main()
    frmMain.Show
End Sub

Public Function GetRequest(ByVal url As String) As String
    Dim http As MSXML2.XMLHTTP
    
    Set http = CreateObject("MSXML2.ServerXMLHTTP")
    
    http.Open "GET", url, False
    http.send

    GetRequest = http.responseText
    
    Set http = Nothing
End Function


