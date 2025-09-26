Attribute VB_Name = "KeysAndSecrets"
Option Explicit
Option Private Module

Private Const ENV_NAME_BINANCE_API_KEY As String = "BINANCE_API_KEY"
Private Const ENV_NAME_BINANCE_API_SECRET As String = "BINANCE_API_SECRET"

' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'
' IMPORTANT:
' Keep API Keys and API Secrets empty during development
'
' Note: Can be filled after deployment in deployed xlsm file
'
' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Private Const BINANCE_API_KEY As String = ""
Private Const BINANCE_API_SECRET As String = ""

Private ApiKey_Binance As String
Private ApiSecret_Binance As String

Public Function GetApiKey_Binance() As String
    If ApiKey_Binance = "" Then
        If BINANCE_API_KEY <> "" Then
            ApiKey_Binance = BINANCE_API_KEY
        Else
            ApiKey_Binance = Environ(ENV_NAME_BINANCE_API_KEY)
        End If
    End If
    GetApiKey_Binance = ApiKey_Binance
End Function

Public Function GetApiSecret_Binance() As String
    If ApiSecret_Binance = "" Then
        If BINANCE_API_SECRET <> "" Then
            ApiSecret_Binance = BINANCE_API_SECRET
        Else
            ApiSecret_Binance = Environ(ENV_NAME_BINANCE_API_SECRET)
        End If
    End If
    GetApiSecret_Binance = ApiSecret_Binance
End Function
