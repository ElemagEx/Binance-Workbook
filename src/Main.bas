Attribute VB_Name = "Main"
Option Explicit

Public Const SHEET_ASSETS As String = "Assets"
Public Const SHEET_CURRENCIES As String = "Currencies"

Public Function GetExchangeTimezone() As String
    Dim table As ListObject
    Set table = ThisWorkbook.Sheets(SHEET_CURRENCIES).ListObjects("Timezone")
    
    GetExchangeTimezone = table.ListColumns("Exchange Timezone").DataBodyRange(1).Value
End Function

Public Sub ResetExchangeTimezone(Optional ByVal tz As String = "")
    Dim table As ListObject
    Set table = ThisWorkbook.Sheets(SHEET_CURRENCIES).ListObjects("Timezone")
    
    table.ListColumns("Exchange Timezone").DataBodyRange(1).Value = tz
End Sub

Public Function Str2Dec(ByVal str As String) As Variant
    Dim val As String
    val = Replace(str, ".", Application.DecimalSeparator)
    Str2Dec = CDec(val)
End Function

