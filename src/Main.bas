Attribute VB_Name = "Main"
Option Explicit
Option Private Module

Public Const SHEET_ASSETS As String = "Assets"
Public Const SHEET_CURRENCIES As String = "Currencies"
Public Const SHEET_MISC As String = "Misc"
Public Const SHEET_LEDGER_TEMPLATE As String = "Ledger Template"

Public Const LEDGER_SHEET_SUFFIX As String = "Ledger"

' Assets Sheet Tables
Public Const TABLE_WALLET As String = "Wallet"
Public Const TABLE_USER_INFO As String = "UserInfo"
Public Const TABLE_LAST_UPDATE As String = "LastUpdate"
' Currencies Sheet Tables
Public Const TABLE_TIMEZONE As String = "Timezone"
Public Const TABLE_CURRENCIES As String = "Currencies"
' Misc Sheet Tables
Public Const TABLE_EVALUATION As String = "Evaluation"
Public Const TABLE_CONVERSIONS As String = "Conversions"
Public Const TABLE_DUST_LOG As String = "DustLog"
' Ledger Template Sheet Tables
Public Const TABLE_TRIVIA As String = "Trivia"
Public Const TABLE_SUMMARY As String = "Summary"
Public Const TABLE_MARKETS As String = "Markets"
Public Const TABLE_DETAILS As String = "Details"

Public Const WALLET_SPOT As String = "Spot"
Public Const WALLET_FUNDING As String = "Funding"
Public Const WALLET_LOCKED_EARN As String = "Locked Earn"
Public Const WALLET_FLEXIBLE_EARN As String = "Flexible Earn"

'Public Const ID_PREFIX_DRIBBLET As String = "FF-"

Public Const ID_PREFIX_WALLET_IN As String = "WI-"  ' Universal Transfer In
Public Const ID_PREFIX_WALLET_OUT As String = "WO-" ' Universal Transfer Out

Public Const ID_PREFIX_WALLET_CVT As String = "WC-" ' Locked Earn to Flexible Earn (transfer)
Public Const ID_PREFIX_WALLET_SEI As String = "WS-" ' Spot to Flexible/Locked Earn (input)
Public Const ID_PREFIX_WALLET_SEO As String = "WT-" ' Spot to Flexible/Locked Earn (output)
Public Const ID_PREFIX_WALLET_FEI As String = "WF-" ' Funding to Flexible/Locked Earn (input)
Public Const ID_PREFIX_WALLET_FEO As String = "WG-" ' Funding to Flexible/Locked Earn (output)

Public Const ID_PREFIX_CONVERT_IN As String = "VI-"
Public Const ID_PREFIX_CONVERT_OUT As String = "VO-"
Public Const ID_PREFIX_CONVERT_FSO As String = "VF-" ' output transfer from funding to spot before spot asset conversion
Public Const ID_PREFIX_CONVERT_ESO As String = "VE-" ' output transfer from earn to spot before spot asset conversion
Public Const ID_PREFIX_CONVERT_EFO As String = "VN-" ' output transfer from earn to funding before funding asset conversion
Public Const ID_PREFIX_CONVERT_FSI As String = "VS-" ' input transfer to spot from funding before spot asset conversion
Public Const ID_PREFIX_CONVERT_ESI As String = "VT-" ' input transfer to spot from earn before spot asset conversion
Public Const ID_PREFIX_CONVERT_EFI As String = "VG-" ' input transfer to funding from earn before funding asset conversion

Public Const ID_PREFIX_DISTRIBUTION As String = "RI-"
Public Const ID_PREFIX_COMMISION As String = "XC-"
Public Const ID_PREFIX_TRADE As String = "XT-"
Public Const ID_PREFIX_FIAT_BUY As String = "TB-"
Public Const ID_PREFIX_FIAT_SELL As String = "TS-"
Public Const ID_PREFIX_FIAT_DEPOSIT As String = "TD-"
Public Const ID_PREFIX_FIAT_WITHDRAW As String = "TW-"
Public Const ID_PREFIX_CRYPTO_DEPOSIT As String = "ND-"
Public Const ID_PREFIX_CRYPTO_WITHDRAW As String = "NW-"

Public Const OP_BUY As String = "buy"
Public Const OP_SELL As String = "sell"
Public Const OP_INCOME As String = "income"
Public Const OP_EXPENCE As String = "expence"
Public Const OP_DUST_IN As String = "dust-in"
Public Const OP_DUST_OUT As String = "dust-out"
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

Public Const ERR_ASSERTION_FAIL As Long = 2000
Public Const ERR_UNKNOWN_WALLET As Long = 2001

Public Const KEY_EVAL_TICKER As String = "Ticker"
Public Const KEY_EVAL_METHOD As String = "Method"
Public Const KEY_EVAL_REFERS As String = "Refers"
Public Const KEY_EVAL_FORMAT As String = "Format"

Public Const MAIN_NUMBER_FORMAT As String = "#,##0.0???????"

Public Const LIST_EVAL_TICKERS As String = "MyEvalTickers"

Public Enum MAX_PERIOD
    FIAT_OPERATIONS = 180
    SIMPLE_EARN_OPERATIONS = 90
    
    SPOT_TRADING_MY_TRADES = 1
    
    CONVERT_TRADE_HISTORY = 30
    
    WALLET_DEPOSIT_HISTORY = 90
    WALLET_WITHDRAW_HISTORY = 90
    WALLET_DIVIDEND_HISTORY = 180
    WALLET_TRANSFER_HISTORY = 180
    WALLET_DRIBLETS_HISTORY = 180
End Enum
Public Enum MAX_LIMIT
    FIAT_OPERATIONS = 500
    SIMPLE_EARN_OPERATIONS = 100
    
    SPOT_TRADING_MY_TRADES = 1000

    CONVERT_TRADE_HISTORY = 1000
    
    WALLET_DEPOSIT_HISTORY = 1000
    WALLET_WITHDRAW_HISTORY = 1000
    WALLET_DIVIDEND_HISTORY = 500
    WALLET_TRANSFER_HISTORY = 100
    WALLET_DRIBLETS_HISTORY = 100
End Enum

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

