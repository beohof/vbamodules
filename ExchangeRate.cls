VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExchangeRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function EXCHANGERATE(ByVal key As String, cur As String, Optional QDate As String = "latest") As String
    ' fixer.io API key needs to be provided
    ' Return exchange rate of a currency in relation to 1 Euro
    Dim c As New WebClient
    Dim Resource As String
    Dim Response As WebResponse
    Dim keys As New Collection
    Dim Json As Object
    
    keys.Add "success"
    
    c.BaseUrl = "http://data.fixer.io/api/"
    
    Resource = QDate & _
    "?" & "access_key=" & key & _
    "&format=1" & _
    "&symbols=USD"
    
    Set Response = c.GetJson(Resource)
    Set Json = JsonConverter.ParseJson(Response.Content)
    
    
    If Json("success") Then
    
        EXCHANGERATE = Json("rates")(cur)
    
    End If
    
End Function

