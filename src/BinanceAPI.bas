Attribute VB_Name = "BinanceAPI"
Option Explicit
Option Private Module

' --- Binance API base URL
Private Const BASE_URL As String = "https://api.binance.com"

Private s_CurrentWeight As Long
'
' Binance SpotTrading Public API: GET /api/v3/ticker/price
'
Public Function SpotTrading_GetPrice(ByVal symbol As String) As Dictionary
    xAddWeight 2
    
    Dim params As New Dictionary
    params.Add "symbol", symbol
    
    Set SpotTrading_GetPrice = xExecuteWebQuery(False, True, "/api/v3/ticker/price", params)
End Function
'
' Binance SpotTrading Public API: GET /api/v3/ticker/price
'
Public Function SpotTrading_GetPrices(Optional ByVal symbols As Collection = Nothing) As Collection
    xAddWeight 4
    
    Dim params As New Dictionary
        
    xAddListParam params, "symbols", symbols
    
    Set SpotTrading_GetPrices = xExecuteWebQuery(False, True, "/api/v3/ticker/price", params)
End Function
'
' Binance SpotTrading Public API: GET /api/v3/exchangeInfo
'
Public Function SpotTrading_GetExchangeInfo(Optional ByVal showPermissionSets As Variant) As Dictionary
    xAddWeight 20
    
    Dim params As New Dictionary

    If Not IsMissing(showPermissionSets) Then
        params.Add "showPermissionSets", IIf(CBool(showPermissionSets), "true", "false")
    End If

    Set SpotTrading_GetExchangeInfo = xExecuteWebQuery(False, True, "/api/v3/exchangeInfo", params)
End Function
'
' Binance SpotTrading Signed API: GET /api/v3/account
'
Public Function SpotTrading_GetAccountInfo(Optional ByVal omitZeroBalances As Variant) As Dictionary
    xAddWeight 20
    
    Dim params As New Dictionary
    
    If Not IsMissing(omitZeroBalances) Then
        params.Add "omitZeroBalances", IIf(CBool(omitZeroBalances), "true", "false")
    End If

    Set SpotTrading_GetAccountInfo = xExecuteWebQuery(True, True, "/api/v3/account", params)
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
    
    xAddWeight IIf(orderId >= 0, 5, 20)
    
    Dim params As New Dictionary
    
    params.Add "symbol", symbol
    
    If startTime <> 0 Then
        params.Add "startTime", DateToUnixTimestamp(startTime)
    End If
    If endTime <> 0 Then
        params.Add "endTime", DateToUnixTimestamp(endTime)
    End If
    If limit >= 0 Then
        params.Add "limit", limit
    End If
    If fromId >= 0 Then
        params.Add "fromId", fromId
    End If
    If fromId >= 0 Then
        params.Add "orderId", orderId
    End If

    Set SpotTrading_GetMyTrades = xExecuteWebQuery(True, True, "/api/v3/myTrades", params)
End Function
'
' Binance Wallet Signed API: GET /sapi/v1/capital/config/getall
'
Public Function Wallet_GetAllCoinsInfo() As Collection
    xAddWeight 10
    
    Dim params As New Dictionary

    Set Wallet_GetAllCoinsInfo = xExecuteWebQuery(True, True, "/sapi/v1/capital/config/getall", params)
End Function
'
' Binance Wallet Signed API: POST /sapi/v3/asset/getUserAsset
'
Public Function Wallet_GetUserAssets(Optional ByVal asset As String = "", Optional ByVal needBtcEvaluation As Variant) As Collection
    xAddWeight 5
    
    Dim params As New Dictionary
    If asset <> "" Then
        params.Add "asset", asset
    End If
    If Not IsMissing(needBtcEvaluation) Then
        params.Add "needBtcEvaluation", IIf(CBool(needBtcEvaluation), "true", "false")
    End If
    
    Set Wallet_GetUserAssets = xExecuteWebQuery(True, False, "/sapi/v3/asset/getUserAsset", params)
End Function
'
' Binance Wallet Signed API: POST /sapi/v1/asset/get-funding-asset
'
Public Function Wallet_GetFundingAssets(Optional ByVal asset As String = "", Optional ByVal needBtcEvaluation As Variant) As Collection
    xAddWeight 1
    
    Dim params As New Dictionary
    
    If asset <> "" Then
        params.Add "asset", asset
    End If
    If Not IsMissing(needBtcEvaluation) Then
        params.Add "needBtcEvaluation", IIf(CBool(needBtcEvaluation), "true", "false")
    End If

    Set Wallet_GetFundingAssets = xExecuteWebQuery(True, False, "/sapi/v1/asset/get-funding-asset", params)
End Function
'
' Binance Wallet Signed API: GET /sapi/v1/asset/dribblet
'
Public Function Wallet_GetDustLog( _
    Optional ByVal startTime As Date = 0, _
    Optional ByVal endTime As Date = 0 _
    ) As Dictionary

    Set Wallet_GetDustLog = Nothing

    Dim params As New Dictionary
    
    If startTime <> 0 Then
        params.Add "startTime", DateToUnixTimestamp(startTime)
    End If
    If endTime <> 0 Then
        params.Add "endTime", DateToUnixTimestamp(endTime)
    End If

    Dim responseText As String
    responseText = ExecuteBinanceSignedQuery("GET", "/sapi/v1/asset/dribblet", params)

    If responseText <> "" Then
        Set Wallet_GetDustLog = JsonConverter.ParseJson(responseText)
    End If
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
    '
    ' WEIGHT = 1
    '
    Set Wallet_GetTransferHistory = Nothing

    Dim params As New Dictionary
    params.Add "type", transferType
    
    If startTime <> 0 Then
        params.Add "startTime", DateToUnixTimestamp(startTime)
    End If
    If endTime <> 0 Then
        params.Add "endTime", DateToUnixTimestamp(endTime)
    End If
    If size >= 0 Then
        params.Add "size", size
    End If
    If current >= 0 Then
        params.Add "current", current
    End If
    If toSymbol <> "" Then
        params.Add "toSymbol", toSymbol
    End If
    If fromSymbol <> "" Then
        params.Add "fromSymbol", fromSymbol
    End If

    Dim responseText As String
    responseText = ExecuteBinanceSignedQuery("GET", "/sapi/v1/asset/transfer", params)

    If responseText <> "" Then
        Set Wallet_GetTransferHistory = JsonConverter.ParseJson(responseText)
    End If
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
    
    xAddWeight 10

    Dim params As New Dictionary
    
    If asset <> "" Then
        params.Add "asset", asset
    End If
    If startTime <> 0 Then
        params.Add "startTime", DateToUnixTimestamp(startTime)
    End If
    If endTime <> 0 Then
        params.Add "endTime", DateToUnixTimestamp(endTime)
    End If
    If limit >= 0 Then
        params.Add "limit", limit
    End If

    Set Wallet_GetDividendHistory = xExecuteWebQuery(True, True, "/sapi/v1/asset/assetDividend", params)
End Function
'
' Binance Wallet Signed API: GET /sapi/v1/capital/deposit/hisrec
'
Public Function Wallet_GetDepositHistory( _
    Optional ByVal includeSource = False, _
    Optional ByVal coin = "", _
    Optional ByVal startTime As Date = 0, _
    Optional ByVal endTime As Date = 0, _
    Optional ByVal limit As Long = -1, _
    Optional ByVal offset As Long = -1, _
    Optional ByVal status As Long = -1 _
    ) As Collection

    Set Wallet_GetDepositHistory = Nothing

    If beginTime <> 0 And endTime <> 0 Then
        If DateDiff("d", beginTime, endTime) > 90 Then
            MsgBox "Time interval must be less than 90 days"
            Exit Function
        End If
    End If

    Dim params As New Dictionary
    
    params.Add "includeSource", includeSource
    
    If coin <> "" Then
        params.Add "coin", coin
    End If
    If startTime <> 0 Then
        params.Add "startTime", DateToUnixTimestamp(startTime)
    End If
    If endTime <> 0 Then
        params.Add "endTime", DateToUnixTimestamp(endTime)
    End If
    If limit >= 0 Then
        params.Add "limit", limit
    End If
    If offset >= 0 Then
        params.Add "offset", offset
    End If
    If status >= 0 Then
        params.Add "status", status
    End If

    Dim responseText As String
    responseText = ExecuteBinanceSignedQuery("GET", "/sapi/v1/capital/deposit/hisrec", params)

    If responseText <> "" Then
        Set Wallet_GetDepositHistory = JsonConverter.ParseJson(responseText)
    End If
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

    Set Wallet_GetWithdrawHistory = Nothing

    If beginTime <> 0 And endTime <> 0 Then
        If DateDiff("d", beginTime, endTime) > 90 Then
            MsgBox "Time interval must be less than 90 days"
            Exit Function
        End If
    End If

    Dim params As New Dictionary
    
    If coin <> "" Then
        params.Add "coin", coin
    End If
    If startTime <> 0 Then
        params.Add "startTime", DateToUnixTimestamp(startTime)
    End If
    If endTime <> 0 Then
        params.Add "endTime", DateToUnixTimestamp(endTime)
    End If
    If limit >= 0 Then
        params.Add "limit", limit
    End If
    If offset >= 0 Then
        params.Add "offset", offset
    End If
    If status >= 0 Then
        params.Add "status", status
    End If

    Dim responseText As String
    responseText = ExecuteBinanceSignedQuery("GET", "/sapi/v1/capital/withdraw/history", params)

    If responseText <> "" Then
        Set Wallet_GetWithdrawHistory = JsonConverter.ParseJson(responseText)
    End If
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
    
    Set Fiat_Orders = Nothing

    Dim params As New Dictionary
    
    params.Add "transactionType", txType
    
    If beginTime <> 0 Then
        params.Add "beginTime", DateToUnixTimestamp(beginTime)
    End If
    If endTime <> 0 Then
        params.Add "endTime", DateToUnixTimestamp(endTime)
    End If
    If rows >= 0 Then
        params.Add "rows", rows
    End If
    If page >= 0 Then
        params.Add "page", page
    End If

    Set Fiat_Orders = xExecuteWebQuery(True, True, "/sapi/v1/fiat/orders", params)
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
    
    Set Fiat_Payments = Nothing

    Dim params As New Dictionary
    
    params.Add "transactionType", txType
    
    If beginTime <> 0 Then
        params.Add "beginTime", DateToUnixTimestamp(beginTime)
    End If
    If endTime <> 0 Then
        params.Add "endTime", DateToUnixTimestamp(endTime)
    End If
    If rows >= 0 Then
        params.Add "rows", rows
    End If
    If page >= 0 Then
        params.Add "page", page
    End If

    Set Fiat_Payments = xExecuteWebQuery(True, True, "/sapi/v1/fiat/payments", params)
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
    '
    ' WEIGHT=3000
    '
    Set Convert_GetConvertHistory = Nothing
    
    Dim params As New Dictionary
    
    params.Add "startTime", DateToUnixTimestamp(startTime)
    params.Add "endTime", DateToUnixTimestamp(endTime)
    
    If limit >= 0 Then
        params.Add "limit", limit
    End If

    Dim responseText As String
    responseText = ExecuteBinanceSignedQuery("GET", "/sapi/v1/convert/tradeFlow", params)

    If responseText <> "" Then
        Set Convert_GetConvertHistory = JsonConverter.ParseJson(responseText)
    End If
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
    
    xAddWeight 150
    
    Dim params As New Dictionary

    If asset <> "" Then
        params.Add "asset", asset
    End If
    If size >= 0 Then
        params.Add "size", size
    End If
    If page >= 0 Then
        params.Add "page", page
    End If
    If productId <> "" Then
        params.Add "productId", productId
    End If
    
    Set SimpleEarn_GetFlexiblePositions = xExecuteWebQuery(True, True, "/sapi/v1/simple-earn/flexible/position", params)
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
    
    xAddWeight 150

    Dim params As New Dictionary
    
    If asset <> "" Then
        params.Add "asset", asset
    End If
    If size >= 0 Then
        params.Add "size", size
    End If
    If page >= 0 Then
        params.Add "page", page
    End If
    If productId <> "" Then
        params.Add "productId", productId
    End If

    Set SimpleEarn_GetLockedPositions = xExecuteWebQuery(True, True, "/sapi/v1/simple-earn/locked/position", params)
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
    Optional ByVal productId As String = "" _
    ) As Dictionary
    
    Set SimpleEarn_GetFlexibleSubscriptions = Nothing

    Dim params As New Dictionary
    
    If asset <> "" Then
        params.Add "asset", asset
    End If
    If startTime <> 0 Then
        params.Add "startTime", DateToUnixTimestamp(startTime)
    End If
    If endTime <> 0 Then
        params.Add "endTime", DateToUnixTimestamp(endTime)
    End If
    If size >= 0 Then
        params.Add "size", size
    End If
    If current >= 0 Then
        params.Add "current", current
    End If
    If productId <> "" Then
        params.Add "productId", productId
    End If

    Dim responseText As String
    responseText = ExecuteBinanceSignedQuery("GET", "/sapi/v1/simple-earn/flexible/history/subscriptionRecord", params)

    If responseText <> "" Then
        Set SimpleEarn_GetFlexibleSubscriptions = JsonConverter.ParseJson(responseText)
    End If
End Function
'
' Binance Simple Earn Signed API: GET /sapi/v1/simple-earn/locked/history/subscriptionRecord
'
Public Function SimpleEarn_GetLockedSubscriptions( _
    Optional ByVal asset As String = "", _
    Optional ByVal startTime As Date = 0, _
    Optional ByVal endTime As Date = 0, _
    Optional ByVal size As Long = -1, _
    Optional ByVal current As Long = -1 _
    ) As Dictionary
    
    Set SimpleEarn_GetLockedSubscriptions = Nothing

    Dim params As New Dictionary
    
    If asset <> "" Then
        params.Add "asset", asset
    End If
    If startTime <> 0 Then
        params.Add "startTime", DateToUnixTimestamp(startTime)
    End If
    If endTime <> 0 Then
        params.Add "endTime", DateToUnixTimestamp(endTime)
    End If
    If size >= 0 Then
        params.Add "size", size
    End If
    If current >= 0 Then
        params.Add "current", current
    End If

    Dim responseText As String
    responseText = ExecuteBinanceSignedQuery("GET", "/sapi/v1/simple-earn/locked/history/subscriptionRecord", params)

    If responseText <> "" Then
        Set SimpleEarn_GetLockedSubscriptions = JsonConverter.ParseJson(responseText)
    End If
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
    Optional ByVal productId As String = "" _
    ) As Dictionary
    
    Set SimpleEarn_GetFlexibleRedemptions = Nothing

    Dim params As New Dictionary
    
    If asset <> "" Then
        params.Add "asset", asset
    End If
    If startTime <> 0 Then
        params.Add "startTime", DateToUnixTimestamp(startTime)
    End If
    If endTime <> 0 Then
        params.Add "endTime", DateToUnixTimestamp(endTime)
    End If
    If size >= 0 Then
        params.Add "size", size
    End If
    If current >= 0 Then
        params.Add "current", current
    End If
    If productId <> "" Then
        params.Add "productId", productId
    End If

    Dim responseText As String
    responseText = ExecuteBinanceSignedQuery("GET", "/sapi/v1/simple-earn/flexible/history/redemptionRecord", params)

    If responseText <> "" Then
        Set SimpleEarn_GetFlexibleRedemptions = JsonConverter.ParseJson(responseText)
    End If
End Function
'
' Binance Simple Earn Signed API: GET /sapi/v1/simple-earn/locked/history/redemptionRecord
'
Public Function SimpleEarn_GetLockedRedemptions( _
    Optional ByVal asset As String = "", _
    Optional ByVal startTime As Date = 0, _
    Optional ByVal endTime As Date = 0, _
    Optional ByVal size As Long = -1, _
    Optional ByVal current As Long = -1 _
    ) As Dictionary
    
    Set SimpleEarn_GetLockedRedemptions = Nothing

    Dim params As New Dictionary
    
    If asset <> "" Then
        params.Add "asset", asset
    End If
    If startTime <> 0 Then
        params.Add "startTime", DateToUnixTimestamp(startTime)
    End If
    If endTime <> 0 Then
        params.Add "endTime", DateToUnixTimestamp(endTime)
    End If
    If size >= 0 Then
        params.Add "size", size
    End If
    If current >= 0 Then
        params.Add "current", current
    End If

    Dim responseText As String
    responseText = ExecuteBinanceSignedQuery("GET", "/sapi/v1/simple-earn/locked/history/redemptionRecord", params)

    If responseText <> "" Then
        Set SimpleEarn_GetLockedRedemptions = JsonConverter.ParseJson(responseText)
    End If
End Function

Private Function xExecuteWebQuery( _
    ByVal signed As Boolean, _
    ByVal isGet As Boolean, _
    ByVal api As String, _
    ByVal params As Dictionary _
    )
    Set xExecuteWebQuery = Nothing
    
    Debug.Print "Starting Binance API web query " & api & " ..."

    Dim client As New WebClient
    client.BaseUrl = BASE_URL
    
    Dim request As New WebRequest
    request.method = IIf(isGet, WebMethod.httpGet, WebMethod.HttpPost)
    request.Resource = api
    request.format = WebFormat.Json
    
    Dim key As Variant
    For Each key In params.Keys
        request.AddQuerystringParam CStr(key), params(key)
    Next key
    
    If signed Then
        request.Headers.Add WebHelpers.CreateKeyValue("X-MBX-APIKEY", GetApiKey_Binance())
    
        request.AddQuerystringParam "timestamp", xGetBinanceServerTime()
        request.AddQuerystringParam "signature", HMACSHA256(xGetQueryFromFullUrl(client.GetFullUrl(request)), GetApiSecret_Binance())
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

' ===================================================================
' ===                     HELPER FUNCTIONS                        ===
' ===================================================================

Private Function ExecuteBinanceSignedQuery(ByVal method As String, ByVal api As String, params As Dictionary) As String
    
    ExecuteBinanceSignedQuery = ""
    
    Debug.Print "Starting Binance Signed API query " & method & " " & api & " ..."

    ' STEP 1: Get the official server time from Binance. This is critical to avoid "Timestamp for this request was ahead/behind..." errors.
    Dim serverTimestamp As String
    serverTimestamp = xGetBinanceServerTime()

    If serverTimestamp = "" Then
        MsgBox "Could not get the server time from Binance. The process will stop.", vbCritical
        Exit Function
    End If

    ' STEP 2: Prepare the parameters for the API request
    params.Add "timestamp", serverTimestamp
    Debug.Print "Using Binance Server Timestamp: " & serverTimestamp

    ' STEP 3: Build the query string and create the signature
    Dim queryString As String
    queryString = BuildQueryStringFromDict(params)

    Dim signature As String
    signature = "" ' xCreateHMACSHA256Signature(queryString)

    ' The final query string includes the signature
    queryString = queryString & "&signature=" & signature

    ' STEP 4: Make the signed API request ---
    Dim finalUrl As String
    finalUrl = BASE_URL & api & "?" & queryString
    
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    
    On Error GoTo HttpErrorHandler
    
    http.Open method, finalUrl, False
    http.SetRequestHeader "X-MBX-APIKEY", GetApiKey_Binance() ' The API key goes in the header
    http.send

    ' STEP 5: Process the response from Binance ---
    If http.status = 200 Then
        ExecuteBinanceSignedQuery = http.responseText
        Debug.Print "Success! Received response from server."
    Else
        ' If something went wrong, Binance sends an error message in the response
        Debug.Print "Binance API Error Occurred!"
        Debug.Print "Status: " & http.status & " " & http.statusText
        Debug.Print "Response: " & http.responseText
        MsgBox "An API error occurred. Status: " & http.status & vbCrLf & "Response: " & http.responseText, vbExclamation
    End If

    Set http = Nothing
    Exit Function

HttpErrorHandler:
    MsgBox "A network error occurred: " & Err.Description, vbCritical
    Set http = Nothing
End Function
Private Function xGetBinanceServerTime() As String
    Dim client As New WebClient
    client.BaseUrl = BASE_URL
    
    Dim response As WebResponse
    Set response = client.GetJson("/api/v3/time")
    
    If response.StatusCode = WebStatusCode.Ok Then
        xGetBinanceServerTime = format(response.data("serverTime"), "0")
    Else
        xGetBinanceServerTime = "0"
    End If
End Function

Private Function xGetQueryFromFullUrl(ByVal url As String) As String
    Dim pos As Long
    pos = InStr(url, "?")
    xGetQueryFromFullUrl = IIf(pos = 0, "", Mid(url, pos + 1))
End Function

' Converts a Dictionary of parameters into a URL query string (e.g., "key=val&key2=val2").
Private Function BuildQueryStringFromDict(params As Dictionary) As String
    Dim parts() As String
    ReDim parts(params.count - 1)
    Dim key As Variant, i As Long
    i = 0
    For Each key In params.Keys
        parts(i) = key & "=" & params(key)
        i = i + 1
    Next key
    BuildQueryStringFromDict = Join(parts, "&")
End Function

' Converts a Unix timestamp (in milliseconds) to a readable VBA date.
Public Function UnixTimestampToDate(ByVal unixTime As Double) As Date
    UnixTimestampToDate = DateAdd("s", unixTime / 1000, "1/1/1970")
End Function

Private Function DateToUnixTimestamp(ByVal dateTime As Date) As String
    Dim val As LongLong
    val = DateDiff("s", #1/1/1970#, dateTime)
    val = val * 1000
    DateToUnixTimestamp = CStr(val)
End Function

Private Function Str2Dec(ByVal str As String) As Variant
    Dim val As String
    val = Replace(str, ".", Application.DecimalSeparator)
    Str2Dec = CDec(val)
End Function

Private Sub xAddWeight(ByVal weight As Long)
    s_CurrentWeight = s_CurrentWeight + weight
End Sub

Private Sub xAddListParam(ByVal params As Dictionary, ByVal name As String, ByVal list As Collection)
    If list Is Nothing Then
        Exit Sub
    End If
    If list.count = 0 Then
        Exit Sub
    End If
    
    Dim val As String
    val = "[""" & list(1) & """"
    Dim i As Long
    For i = 2 To list.count
        val = val & ",""" & list(i) & """"
    Next i
    val = val & "]"
    params.Add name, val
End Sub

