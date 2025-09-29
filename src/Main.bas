Attribute VB_Name = "Main"
Option Explicit
Option Private Module

Public Const SHEET_ASSETS As String = "Assets"
Public Const SHEET_CURRENCIES As String = "Currencies"
Public Const SHEET_LEDGER_TEMPLATE As String = "Ledger Template"

Public Const LEDGER_SHEET_SUFFIX As String = "Ledger"

Public Const TABLE_WALLET As String = "Wallet"
Public Const TABLE_USER_INFO As String = "UserInfo"
Public Const TABLE_LAST_UPDATE As String = "LastUpdate"

Public Const TABLE_TIMEZONE As String = "Timezone"
Public Const TABLE_CURRENCIES As String = "Currencies"

Public Const TABLE_TRIVIA As String = "Trivia"
Public Const TABLE_SUMMARY As String = "Summary"
Public Const TABLE_MARKETS As String = "Markets"
Public Const TABLE_DETAILS As String = "Details"

Public Const WALLET_SPOT As String = "Spot"
Public Const WALLET_FUNDING As String = "Funding"

Public Const OPERATION_BUY As String = "buy"
Public Const OPERATION_SELL As String = "sell"
Public Const OPERATION_INCOME As String = "income"
Public Const OPERATION_EXPENCE As String = "expence"
Public Const OPERATION_DUST_IN As String = "dust-in"
Public Const OPERATION_DUST_OUT As String = "dust-out"
Public Const OPERATION_FIAT_BUY As String = "fiat-buy"
Public Const OPERATION_FIAT_SELL As String = "fiat-sell"
Public Const OPERATION_DEPOSIT As String = "deposit"
Public Const OPERATION_WITHDRAW As String = "withdraw"
Public Const OPERATION_WALLET_IN As String = "wallet-in"
Public Const OPERATION_WALLET_OUT As String = "wallet-out"
Public Const OPERATION_CONVERT_IN As String = "convert-in"
Public Const OPERATION_CONVERT_OUT As String = "convert-out"
Public Const OPERATION_COMMISSION As String = "commission"
Public Const OPERATION_DISTRIBUTION As String = "distribution"

Public Const STR_NA As String = "n/a"
Public Const STR_BNB As String = "BNB"
Public Const STR_FIAT As String = "Fiat"
Public Const STR_CRYPTO As String = "Crypto"

Public Const ERR_ASSERTION_FAIL = 2000
Public Const ERR_UNKNOWN_WALLET = 2001

Public Enum MAX_PERIOD
    FIAT_OPERATIONS = 180
    
    SPOT_TRADING_GET_MY_TRADES = 1
    
    WALLET_DEPOSIT_HISTORY = 90
    WALLET_WITHDRAW_HISTORY = 90
    WALLET_DIVIDEND_HISTORY = 180
    WALLET_TRANSFER_HISTORY = 180
End Enum
Public Enum MAX_LIMIT
    FIAT_OPERATIONS = 500
    
    SPOT_TRADING_GET_MY_TRADES = 1000

    WALLET_DEPOSIT_HISTORY = 1000
    WALLET_WITHDRAW_HISTORY = 1000
    WALLET_DIVIDEND_HISTORY = 500
    WALLET_TRANSFER_HISTORY = 100
    
    SIMPLE_EARN_FLEXIBLE_POSITIONS = 100
    SIMPLE_EARN_LOCKED_POSITIONS = 100
End Enum

Public Const aMAX_PERIOD_WALLET_TRANSFERS As Long = 180

Public Const aMAX_PERIOD_SIMPLEEARN_FLEXIBLE_SUBSCRIPTIONS = 90
Public Const aMAX_PERIOD_SIMPLEEARN_FLEXIBLE_REDEMPTIONS = 90
Public Const aMAX_PERIOD_LOCKED_FLEXIBLE_SUBSCRIPTIONS = 90
Public Const aMAX_PERIOD_LOCKED_FLEXIBLE_REDEMPTIONS = 90

Public Const aMAX_LIMIT_WALLET_TRANSFERS As Long = 100

Public Const aMAX_LIMIT_SIMPLEEARN_FLEXIBLE_SUBSCRIPTIONS = 100
Public Const aMAX_LIMIT_SIMPLEEARN_FLEXIBLE_REDEMPTIONS = 100
Public Const aMAX_LIMIT_SIMPLEEARN_LOCKED_SUBSCRIPTIONS = 100
Public Const aMAX_LIMIT_SIMPLEEARN_LOCKED_REDEMPTIONS = 100

#Const IN_DEVELOPMENT = True

Public Sub Assert_Fail(Optional ByVal source As String = "", Optional ByVal desc As String = "")
#If IN_DEVELOPMENT Then
    Debug.Print "Assertion Failed: " & desc
    Debug.Print "Source: " & source
    Stop
#Else
    Err.Raise ERR_ASSERTION_FAIL, source, desc
#End If
End Sub

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

Public Function UnixTimestamp2Date(ByVal unixTime As LongLong) As Date
    UnixTimestamp2Date = DateAdd("s", unixTime / 1000, "1/1/1970")
End Function

Public Function UnixTimestampToDate(ByVal unixTime As Double) As Date
    UnixTimestampToDate = DateAdd("s", unixTime / 1000, "1/1/1970")
End Function

Public Function DateToUnixTimestamp(ByVal dateTime As Date) As String
    Dim val As LongLong
    val = DateDiff("s", #1/1/1970#, dateTime)
    val = val * 1000
    DateToUnixTimestamp = CStr(val)
End Function

Public Function Str2Dec(ByVal str As String) As Variant
    Dim val As String
    val = Replace(str, ".", Application.DecimalSeparator)
    Str2Dec = CDec(val)
End Function

