VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BinanceWebQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public isPost As Boolean
Public api As String
Public weight As Long
Public isSigned As Boolean
Public params As New Dictionary

Public Sub AddBooleanParam(ByVal key As String, ByVal val As Variant)
    If Not IsMissing(val) Then
        params.Add key, IIf(CBool(val), "true", "false")
    End If
End Sub

Public Sub AddStringParam(ByVal key As String, ByVal val As String)
    If Len(val) <> 0 Then
        params.Add key, val
    End If
End Sub

Public Sub AddLongLongParam(ByVal key As String, ByVal val As LongLong)
    If val >= 0 Then
        params.Add key, val
    End If
End Sub

Public Sub AddLongParam(ByVal key As String, ByVal val As Long)
    If val >= 0 Then
        params.Add key, val
    End If
End Sub

Public Sub AddDateParam(ByVal key As String, ByVal val As Date)
    If val <> 0 Then
        params.Add key, xDateToUnixTimestamp(val)
    End If
End Sub

Public Sub AddListParam(ByVal key As String, ByVal vals As Collection)
    If vals Is Nothing Then
        Exit Sub
    End If
    If vals.count = 0 Then
        Exit Sub
    End If
    
    Dim val As String
    val = "[""" & vals(1) & """"
    Dim i As Long
    For i = 2 To vals.count
        val = val & ",""" & vals(i) & """"
    Next i
    val = val & "]"
    
    params.Add key, val
End Sub

Public Sub Sign()
    params.Add "timestamp", xGetBinanceServerTime()
    
    Dim val As String
    Dim key As Variant
    For Each key In params.Keys
        If Len(val) > 0 Then
            val = val & "&"
        End If
        val = val & CStr(key) & "=" & params(key)
    Next key
    
    params.Add "signature", HMACSHA256(val, GetApiSecret_Binance())
    
    isSigned = True
End Sub

Private Function xGetBinanceServerTime() As String
    Dim client As New WebClient
    client.BaseUrl = BinanceAPI.BASE_URL
    
    Dim response As WebResponse
    Set response = client.GetJson("/api/v3/time")
    
    If response.StatusCode = WebStatusCode.Ok Then
        xGetBinanceServerTime = format(response.data("serverTime"), "0")
    Else
        xGetBinanceServerTime = "0"
    End If
End Function

Public Function xDateToUnixTimestamp(ByVal dateTime As Date) As String
    Dim val As LongLong
    val = DateDiff("s", #1/1/1970#, dateTime)
    val = val * 1000
    xDateToUnixTimestamp = CStr(val)
End Function

