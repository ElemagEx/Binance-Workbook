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
Public Const WALLET_LOCKED_EARN As String = "Locked Earn"
Public Const WALLET_FLEXIBLE_EARN As String = "Flexible Earn"

Public Const OP_BUY As String = "buy"
Public Const OP_SELL As String = "sell"
Public Const OP_INCOME As String = "income"
Public Const OP_EXPENCE As String = "expence"
Public Const OP_DUST_IN As String = "dust-in"
Public Const OP_DUST_OUT As String = "dust-out"
Public Const OP_FIAT_BUY As String = "fiat-buy"
Public Const OP_FIAT_SELL As String = "fiat-sell"
Public Const OP_DEPOSIT As String = "deposit"
Public Const OP_WITHDRAW As String = "withdraw"
Public Const OP_WALLET_IN As String = "wallet-in"
Public Const OP_WALLET_OUT As String = "wallet-out"
Public Const OP_CONVERT_IN As String = "convert-in"
Public Const OP_CONVERT_OUT As String = "convert-out"
Public Const OP_COMMISSION As String = "commission"
Public Const OP_DISTRIBUTION As String = "distribution"

Public Const STR_NA As String = "n/a"
Public Const STR_BNB As String = "BNB"
Public Const STR_FIAT As String = "Fiat"
Public Const STR_CRYPTO As String = "Crypto"

Public Const ERR_ASSERTION_FAIL = 2000
Public Const ERR_UNKNOWN_WALLET = 2001

Public Enum MAX_PERIOD
    FIAT_OPERATIONS = 180
    SIMPLE_EARN_OPERATIONS = 90
    
    SPOT_TRADING_GET_MY_TRADES = 1
    
    WALLET_DEPOSIT_HISTORY = 90
    WALLET_WITHDRAW_HISTORY = 90
    WALLET_DIVIDEND_HISTORY = 180
    WALLET_TRANSFER_HISTORY = 180
End Enum
Public Enum MAX_LIMIT
    FIAT_OPERATIONS = 500
    SIMPLE_EARN_OPERATIONS = 100
    
    SPOT_TRADING_GET_MY_TRADES = 1000

    WALLET_DEPOSIT_HISTORY = 1000
    WALLET_WITHDRAW_HISTORY = 1000
    WALLET_DIVIDEND_HISTORY = 500
    WALLET_TRANSFER_HISTORY = 100
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

Public Sub HandleError(ByVal num As Long, ByVal src As String, ByVal desc As String, Optional ByVal showMsg As Boolean = True)
    Dim msg As String
    msg = "Error Source: " & src & vbCrLf & _
          "Error Number: " & num & vbCrLf & _
          "Description: " & desc
          
    Debug.Print msg
          
    If showMsg Then MsgBox msg, vbCritical
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

Public Function Str2Dec(ByVal str As String) As Variant
    Dim val As String
    val = Replace(str, ".", Application.DecimalSeparator)
    Str2Dec = CDec(val)
End Function

