Attribute VB_Name = "BinanceAPI"
Option Explicit
Option Private Module

' --- Binance API base URL
Public Const BASE_URL As String = "https://api.binance.com"

Private s_CurrentWeight As Long
'
' Binance SpotTrading Public API: GET /api/v3/ticker/price
'
Public Function SpotTrading_GetPrice(ByVal symbol As String) As Dictionary
    Dim query As New BinanceWebQuery
    
    query.api = "/api/v3/ticker/price"
    query.weight = 2
    
    query.AddStringParam "symbol", symbol
    
    Set SpotTrading_GetPrice = xExecuteWebQuery(query)
End Function
'
' Binance SpotTrading Public API: GET /api/v3/ticker/price
'
Public Function SpotTrading_GetPrices(Optional ByVal symbols As Collection = Nothing) As Collection
    Dim query As New BinanceWebQuery
    
    query.api = "/api/v3/ticker/price"
    query.weight = 4
    
    query.AddListParam "symbols", symbols
    
    Set SpotTrading_GetPrices = xExecuteWebQuery(query)
End Function
'
' Binance SpotTrading Public API: GET /api/v3/exchangeInfo
'
Public Function SpotTrading_GetExchangeInfo(Optional ByVal showPermissionSets As Variant) As Dictionary
    Dim query As New BinanceWebQuery
    
    query.api = "/api/v3/exchangeInfo"
    query.weight = 20
    
    query.AddBooleanParam "showPermissionSets", showPermissionSets
    
    Set SpotTrading_GetExchangeInfo = xExecuteWebQuery(query)
End Function
'
' Binance SpotTrading Signed API: GET /api/v3/account
'
Public Function SpotTrading_GetAccountInfo(Optional ByVal omitZeroBalances As Variant) As Dictionary
    Dim query As New BinanceWebQuery
    
    query.api = "/api/v3/account"
    query.weight = 20
    
    query.AddBooleanParam "omitZeroBalances", omitZeroBalances
    query.Sign
    
    Set SpotTrading_GetAccountInfo = xExecuteWebQuery(query)
End Function
'
' Binance SpotTrading Signed API: GET /api/v3/myTrades
'
Public Function SpotTrading_GetMyTrades( _
    ByVal symbol As String, _
    Optional ByVal startTime As Date = 0, _
    Optional ByVal endTime As Date = 0, _
    Optional ByVal limit As Long = -1, _
    Optional ByVal fromId As LongLong = -1, _
    Optional ByVal orderId As LongLong = -1 _
    ) As Collection
    Dim query As New BinanceWebQuery
    
    query.api = "/api/v3/myTrades"
    query.weight = IIf(orderId >= 0, 5, 20)
    
    query.AddBooleanParam "omitZeroBalances", omitZeroBalances
    query.AddStringParam "symbol", symbol
    query.AddDateParam "startTime", startTime
    query.AddDateParam "endTime", endTime
    query.AddLongParam "limit", limit
    query.AddLongLongParam "fromId", fromId
    query.AddLongLongParam "orderId", orderId
    query.Sign

    Set SpotTrading_GetMyTrades = xExecuteWebQuery(query)
End Function
'
' Binance Wallet Signed API: GET /sapi/v1/capital/config/getall
'
Public Function Wallet_GetAllCoinsInfo() As Collection
    Dim query As New BinanceWebQuery
    
    query.api = "/sapi/v1/capital/config/getall"
    query.weight = 10
    
    query.Sign
    
    Set Wallet_GetAllCoinsInfo = xExecuteWebQuery(query)
End Function
'
' Binance Wallet Signed API: POST /sapi/v3/asset/getUserAsset
'
Public Function Wallet_GetUserAssets(Optional ByVal asset As String = "", Optional ByVal needBtcEvaluation As Variant) As Collection
    Dim query As New BinanceWebQuery
    
    query.isPost = True
    query.api = "/sapi/v3/asset/getUserAsset"
    query.weight = 5
    
    query.AddStringParam "asset", asset
    query.AddBooleanParam "needBtcEvaluation", needBtcEvaluation
    
    query.Sign
    
    Set Wallet_GetUserAssets = xExecuteWebQuery(query)
End Function
'
' Binance Wallet Signed API: POST /sapi/v1/asset/get-funding-asset
'
Public Function Wallet_GetFundingAssets(Optional ByVal asset As String = "", Optional ByVal needBtcEvaluation As Variant) As Collection
    Dim query As New BinanceWebQuery
    
    query.isPost = True
    query.api = "/sapi/v1/asset/get-funding-asset"
    query.weight = 1
    
    query.AddStringParam "asset", asset
    query.AddBooleanParam "needBtcEvaluation", needBtcEvaluation
    
    query.Sign

    Set Wallet_GetFundingAssets = xExecuteWebQuery(query)
End Function
'
' Binance Wallet Signed API: GET /sapi/v1/asset/dribblet
'
Public Function Wallet_GetDustLog( _
    Optional ByVal startTime As Date = 0, _
    Optional ByVal endTime As Date = 0 _
    ) As Dictionary
    Dim query As New BinanceWebQuery
    
    query.api = "/sapi/v1/asset/dribblet"
    query.weight = 1
    
    query.AddDateParam "startTime", startTime
    query.AddDateParam "endTime", endTime
    query.Sign

    Set Wallet_GetDustLog = xExecuteWebQuery(query)
End Function
'
' Binance Wallet Signed API: GET /sapi/v1/asset/transfer
'
Public Function Wallet_GetTransferHistory( _
    ByVal transferType As String, _
    Optional ByVal startTime As Date = 0, _
    Optional ByVal endTime As Date = 0, _
    Optional ByVal size As Long = -1, _
    Optional ByVal current As Long = -1, _
    Optional ByVal toSymbol As String = "", _
    Optional ByVal fromSymbol As String = "" _
    ) As Dictionary
    Dim query As New BinanceWebQuery
    
    query.api = "/sapi/v1/asset/transfer"
    query.weight = 1
    
    query.AddStringParam "type", transferType
    query.AddDateParam "startTime", startTime
    query.AddDateParam "endTime", endTime
    query.AddLongParam "size", size
    query.AddLongParam "current", current
    query.AddStringParam "toSymbol", toSymbol
    query.AddStringParam "fromSymbol", fromSymbol
    query.Sign
    
    Set Wallet_GetTransferHistory = xExecuteWebQuery(query)
End Function
'
' Binance Wallet Signed API: GET /sapi/v1/asset/assetDividend
'
Public Function Wallet_GetDividendHistory( _
    Optional ByVal asset As String = "", _
    Optional ByVal startTime As Date = 0, _
    Optional ByVal endTime As Date = 0, _
    Optional ByVal limit As Long = -1 _
    ) As Dictionary
    Dim query As New BinanceWebQuery
    
    query.api = "/sapi/v1/asset/assetDividend"
    query.weight = 10
    
    query.AddStringParam "asset", asset
    query.AddDateParam "startTime", startTime
    query.AddDateParam "endTime", endTime
    query.AddLongParam "limit", limit
    query.Sign
    
    Set Wallet_GetDividendHistory = xExecuteWebQuery(query)
End Function
'
' Binance Wallet Signed API: GET /sapi/v1/capital/deposit/hisrec
'
Public Function Wallet_GetDepositHistory( _
    Optional ByVal coin = "", _
    Optional ByVal startTime As Date = 0, _
    Optional ByVal endTime As Date = 0, _
    Optional ByVal limit As Long = -1, _
    Optional ByVal offset As Long = -1, _
    Optional ByVal status As Long = -1, _
    Optional ByVal includeSource As Variant _
    ) As Collection
    Dim query As New BinanceWebQuery
    
    query.api = "/sapi/v1/capital/deposit/hisrec"
    query.weight = 1
    
    query.AddStringParam "coin", coin
    query.AddDateParam "startTime", startTime
    query.AddDateParam "endTime", endTime
    query.AddLongParam "limit", limit
    query.AddLongParam "offset", offset
    query.AddLongParam "status", status
    query.AddBooleanParam "includeSource", includeSource
    query.Sign
    
    Set Wallet_GetDepositHistory = xExecuteWebQuery(query)
End Function
'
' Binance Wallet Signed API: GET /sapi/v1/capital/withdraw/history
'
Public Function Wallet_GetWithdrawHistory( _
    Optional ByVal coin = "", _
    Optional ByVal startTime As Date = 0, _
    Optional ByVal endTime As Date = 0, _
    Optional ByVal limit As Long = -1, _
    Optional ByVal offset As Long = -1, _
    Optional ByVal status As Long = -1 _
    ) As Collection
    Dim query As New BinanceWebQuery
    
    query.api = "/sapi/v1/capital/withdraw/history"
    query.weight = 18000
    
    query.AddStringParam "coin", coin
    query.AddDateParam "startTime", startTime
    query.AddDateParam "endTime", endTime
    query.AddLongParam "limit", limit
    query.AddLongParam "offset", offset
    query.AddLongParam "status", status
    query.Sign
    
    Set Wallet_GetWithdrawHistory = xExecuteWebQuery(query)
End Function
'
' Binance Fiat Signed API: GET /sapi/v1/fiat/orders
'
Public Function Fiat_Orders( _
    ByVal txType As String, _
    Optional ByVal beginTime As Date = 0, _
    Optional ByVal endTime As Date = 0, _
    Optional ByVal rows = -1, _
    Optional ByVal page = -1 _
    ) As Dictionary
    Dim query As New BinanceWebQuery
    
    query.api = "/sapi/v1/fiat/orders"
    query.weight = 45000
    
    query.AddStringParam "transactionType", txType
    query.AddDateParam "beginTime", beginTime
    query.AddDateParam "endTime", endTime
    query.AddLongParam "rows", rows
    query.AddLongParam "page", page
    query.Sign

    Set Fiat_Orders = xExecuteWebQuery(query)
End Function
Public Function Fiat_GetDeposits( _
    Optional ByVal beginTime As Date = 0, _
    Optional ByVal endTime As Date = 0, _
    Optional ByVal rows = -1, _
    Optional ByVal page = -1 _
    ) As Dictionary
    Set Fiat_GetDeposits = Fiat_Orders("0", beginTime, endTime, rows, page)
End Function
Public Function Fiat_GetWithdraws( _
    Optional ByVal beginTime As Date = 0, _
    Optional ByVal endTime As Date = 0, _
    Optional ByVal rows = -1, _
    Optional ByVal page = -1 _
    ) As Dictionary
    Set Fiat_GetWithdraws = Fiat_Orders("1", beginTime, endTime, rows, page)
End Function
'
' Binance Fiat Signed API: GET /sapi/v1/fiat/payments
'
Public Function Fiat_Payments( _
    ByVal txType As String, _
    Optional ByVal beginTime As Date = 0, _
    Optional ByVal endTime As Date = 0, _
    Optional ByVal rows = -1, _
    Optional ByVal page = -1 _
    ) As Dictionary
    Dim query As New BinanceWebQuery
    
    query.api = "/sapi/v1/fiat/payments"
    query.weight = 1
    
    query.AddStringParam "transactionType", txType
    query.AddDateParam "beginTime", beginTime
    query.AddDateParam "endTime", endTime
    query.AddLongParam "rows", rows
    query.AddLongParam "page", page
    query.Sign

    Set Fiat_Payments = xExecuteWebQuery(query)
End Function
Public Function Fiat_GetBuys( _
    Optional ByVal beginTime As Date = 0, _
    Optional ByVal endTime As Date = 0, _
    Optional ByVal rows = -1, _
    Optional ByVal page = -1 _
    ) As Dictionary
    Set Fiat_GetBuys = Fiat_Payments("0", beginTime, endTime, rows, page)
End Function
Public Function Fiat_GetSells( _
    Optional ByVal beginTime As Date = 0, _
    Optional ByVal endTime As Date = 0, _
    Optional ByVal rows = -1, _
    Optional ByVal page = -1 _
    ) As Dictionary
    Set Fiat_GetSells = Fiat_Payments("1", beginTime, endTime, rows, page)
End Function
'
' Binance Convert Public API: GET /sapi/v1/convert/tradeFlow
'
Public Function Convert_GetConvertHistory( _
    ByVal startTime As Date, _
    ByVal endTime As Date, _
    Optional ByVal limit As Long = -1 _
    ) As Dictionary
    Dim query As New BinanceWebQuery
    
    query.api = "/sapi/v1/convert/tradeFlow"
    query.weight = 3000
    
    query.AddDateParam "startTime", startTime
    query.AddDateParam "endTime", endTime
    query.AddLongParam "limit", limit
    query.Sign
    
    Set Convert_GetConvertHistory = xExecuteWebQuery(query)
End Function
'
' Binance Simple Earn Signed API: GET /sapi/v1/simple-earn/flexible/position
'
Public Function SimpleEarn_GetFlexiblePositions( _
    Optional ByVal asset As String = "", _
    Optional ByVal size As Long = -1, _
    Optional ByVal page As Long = -1, _
    Optional ByVal productId As String = "" _
    ) As Dictionary
    Dim query As New BinanceWebQuery
    
    query.api = "/sapi/v1/simple-earn/flexible/position"
    query.weight = 150
    
    query.AddStringParam "asset", asset
    query.AddLongParam "size", size
    query.AddLongParam "page", page
    query.AddStringParam "productId", productId
    query.Sign
    
    Set SimpleEarn_GetFlexiblePositions = xExecuteWebQuery(query)
End Function
'
' Binance Simple Earn Signed API: GET /sapi/v1/simple-earn/locked/position
'
Public Function SimpleEarn_GetLockedPositions( _
    Optional ByVal asset As String = "", _
    Optional ByVal size As Long = -1, _
    Optional ByVal page As Long = -1, _
    Optional ByVal productId As String = "" _
    ) As Dictionary
    Dim query As New BinanceWebQuery
    
    query.api = "/sapi/v1/simple-earn/locked/position"
    query.weight = 150
    
    query.AddStringParam "asset", asset
    query.AddLongParam "size", size
    query.AddLongParam "page", page
    query.AddStringParam "productId", productId
    query.Sign
    
    Set SimpleEarn_GetLockedPositions = xExecuteWebQuery(query)
End Function
'
' Binance Simple Earn Signed API: GET /sapi/v1/simple-earn/flexible/history/subscriptionRecord
'
Public Function SimpleEarn_GetFlexibleSubscriptions( _
    Optional ByVal asset As String = "", _
    Optional ByVal startTime As Date = 0, _
    Optional ByVal endTime As Date = 0, _
    Optional ByVal size As Long = -1, _
    Optional ByVal current As Long = -1, _
    Optional ByVal productId As String = "", _
    Optional ByVal purchaseId As String = "" _
    ) As Dictionary
    Dim query As New BinanceWebQuery
    
    query.api = "/sapi/v1/simple-earn/flexible/history/subscriptionRecord"
    query.weight = 150
    
    query.AddStringParam "asset", asset
    query.AddDateParam "startTime", startTime
    query.AddDateParam "endTime", endTime
    query.AddLongParam "size", size
    query.AddLongParam "current", current
    query.AddStringParam "productId", productId
    query.AddStringParam "purchaseId", purchaseId
    query.Sign
    
    Set SimpleEarn_GetFlexibleSubscriptions = xExecuteWebQuery(query)
End Function
'
' Binance Simple Earn Signed API: GET /sapi/v1/simple-earn/locked/history/subscriptionRecord
'
Public Function SimpleEarn_GetLockedSubscriptions( _
    Optional ByVal asset As String = "", _
    Optional ByVal startTime As Date = 0, _
    Optional ByVal endTime As Date = 0, _
    Optional ByVal size As Long = -1, _
    Optional ByVal current As Long = -1, _
    Optional ByVal purchaseId As String = "" _
    ) As Dictionary
    Dim query As New BinanceWebQuery
    
    query.api = "/sapi/v1/simple-earn/locked/history/subscriptionRecord"
    query.weight = 150
    
    query.AddStringParam "asset", asset
    query.AddDateParam "startTime", startTime
    query.AddDateParam "endTime", endTime
    query.AddLongParam "size", size
    query.AddLongParam "current", current
    query.AddStringParam "purchaseId", purchaseId
    query.Sign
    
    Set SimpleEarn_GetLockedSubscriptions = xExecuteWebQuery(query)
End Function
'
' Binance Simple Earn Signed API: GET /sapi/v1/simple-earn/flexible/history/redemptionRecord
'
Public Function SimpleEarn_GetFlexibleRedemptions( _
    Optional ByVal asset As String = "", _
    Optional ByVal startTime As Date = 0, _
    Optional ByVal endTime As Date = 0, _
    Optional ByVal size As Long = -1, _
    Optional ByVal current As Long = -1, _
    Optional ByVal redeemId As String = "", _
    Optional ByVal purchaseId As String = "" _
    ) As Dictionary
    Dim query As New BinanceWebQuery
    
    query.api = "/sapi/v1/simple-earn/flexible/history/redemptionRecord"
    query.weight = 150
    
    query.AddStringParam "asset", asset
    query.AddDateParam "startTime", startTime
    query.AddDateParam "endTime", endTime
    query.AddLongParam "size", size
    query.AddLongParam "current", current
    query.AddStringParam "redeemId", redeemId
    query.AddStringParam "purchaseId", purchaseId
    query.Sign
    
    Set SimpleEarn_GetFlexibleRedemptions = xExecuteWebQuery(query)
End Function
'
' Binance Simple Earn Signed API: GET /sapi/v1/simple-earn/locked/history/redemptionRecord
'
Public Function SimpleEarn_GetLockedRedemptions( _
    Optional ByVal asset As String = "", _
    Optional ByVal startTime As Date = 0, _
    Optional ByVal endTime As Date = 0, _
    Optional ByVal size As Long = -1, _
    Optional ByVal current As Long = -1, _
    Optional ByVal redeemId As String = "", _
    Optional ByVal positionId As String = "" _
    ) As Dictionary
    Dim query As New BinanceWebQuery
    
    query.api = "/sapi/v1/simple-earn/locked/history/redemptionRecord"
    query.weight = 150
    
    query.AddStringParam "asset", asset
    query.AddDateParam "startTime", startTime
    query.AddDateParam "endTime", endTime
    query.AddLongParam "size", size
    query.AddLongParam "current", current
    query.AddStringParam "redeemId", redeemId
    query.AddStringParam "positionId", positionId
    query.Sign
    
    Set SimpleEarn_GetLockedRedemptions = xExecuteWebQuery(query)
End Function

Private Function xExecuteWebQuery(ByVal query As BinanceWebQuery)
    Set xExecuteWebQuery = Nothing
    
    Debug.Print "Starting Binance API web query " & query.api & " ..."

    Dim client As New WebClient
    client.BaseUrl = BASE_URL
    
    Dim request As New WebRequest
    request.method = IIf(query.isPost, WebMethod.HttpPost, WebMethod.httpGet)
    request.Resource = query.api
    request.format = WebFormat.Json

    Dim key As Variant
    For Each key In query.params.Keys
        request.AddQuerystringParam CStr(key), query.params(key)
    Next key
    
    If query.isSigned Then
        request.Headers.Add WebHelpers.CreateKeyValue("X-MBX-APIKEY", GetApiKey_Binance())
    End If
    
    Dim response As WebResponse
    Set response = client.Execute(request)
    
    xCheckResponseHeaders response.Headers
    
    If response.StatusCode = WebStatusCode.Ok Then
        Debug.Print "Success! Received response from server."
        Set xExecuteWebQuery = response.data
    Else
        Debug.Print "Binance API Error Occurred!"
        Debug.Print "Status: " & response.StatusCode & " " & response.StatusDescription
        Debug.Print "Response: " & response.Content
        MsgBox "An API error occurred. Status: " & response.StatusCode & vbCrLf & "Response: " & response.StatusDescription, vbExclamation
    End If
End Function
    
Private Sub xCheckResponseHeaders(ByVal responseHeaders As Collection)
    Dim header As Variant
    For Each header In responseHeaders
        Dim name As String
        name = header("Key")
        If name = "x-mbx-used-weight" Or name = "x-mbx-used-weight-1m" Then
            Debug.Print name & ": " & header("Value")
        End If
    Next header
End Sub
