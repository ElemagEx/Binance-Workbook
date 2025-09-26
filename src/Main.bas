Attribute VB_Name = "Main"
Option Explicit
Option Private Module

Public Const SHEET_ASSETS As String = "Assets"
Public Const SHEET_CURRENCIES As String = "Currencies"

Public Const TABLE_TIMEZONE As String = "Timezone"

Public Const STR_FIAT As String = "Fiat"
Public Const STR_CRYPTO As String = "Crypto"

Public Function GetExchangeTimezone() As String
    Dim table As ListObject
    Set table = ThisWorkbook.Sheets(SHEET_CURRENCIES).ListObjects(TABLE_TIMEZONE)
    
    GetExchangeTimezone = table.ListColumns("Exchange Timezone").DataBodyRange(1).Value
End Function

Public Sub ResetExchangeTimezone(Optional ByVal tz As String = "")
    Dim table As ListObject
    Set table = ThisWorkbook.Sheets(SHEET_CURRENCIES).ListObjects(TABLE_TIMEZONE)
    
    table.ListColumns("Exchange Timezone").DataBodyRange(1).Value = tz
End Sub

Public Function Str2Dec(ByVal str As String) As Variant
    Dim val As String
    val = Replace(str, ".", Application.DecimalSeparator)
    Str2Dec = CDec(val)
End Function

