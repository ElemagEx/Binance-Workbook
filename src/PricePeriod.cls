VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PricePeriod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ticker As String
Public interval As String
Public startDate As Date
Public endDate As Date
Public path As String
Public data As New Dictionary

Public Function FindKline(ByVal symbol As String, ByVal stamp As Date) As PriceKline
    Set FindKline = Nothing
    
    If Not data.Exists(symbol) Then Exit Function
    
    Dim klines As Collection
    Set klines = data(symbol)
    
    Dim kline As Variant
    For Each kline In klines
        If kline.startTime <= stamp And stamp <= kline.endTime Then
            Set FindKline = kline
            Exit Function
        End If
    Next kline
End Function
