VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PriceEvaluator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_mainFiatQuote As String
Private m_prices As Dictionary
Private m_quotes As Collection
Private m_paths As Dictionary
Private m_infos As Dictionary
Private m_quote As PriceQuote

Public Sub GetAllEvalSymbols()
    Assert_Check m_prices Is Nothing
    
    Set m_prices = xGetPrices()
    
    Set m_infos = TableEvaluation.GetInfos()
    
    m_mainFiatQuote = GetMainFiatCurrency()
    xCollectCommonQuotes
End Sub

Public Sub GetQuotePrices(ByVal quote As String, Optional ByVal paths As Dictionary = Nothing)
    Assert_Check m_prices Is Nothing
    
    xInitQuote quote, paths
    
    Dim symbols As Collection
    Set symbols = xCalcSymbols()
    
    Set m_prices = xGetPrices(symbols)
    
    Set m_infos = TableEvaluation.GetInfos()
    
    m_mainFiatQuote = GetMainFiatCurrency()
    xCollectCommonQuotes
End Sub

Public Function GetQuoteInfo() As PriceQuote
    Set GetQuoteInfo = m_quote
End Function

Public Function GetInfos() As Dictionary
    Assert_Check Not m_infos Is Nothing
    
    Set GetInfos = m_infos
End Function

Public Function EvaluatePath(ByVal isFiat As Boolean, ByVal base As String, ByVal quote As String) As String
    Assert_Check Not m_prices Is Nothing
    Assert_Check m_infos.Exists(m_mainFiatQuote)
    
    Dim path As String
    
    If isFiat Then
        path = xFindFiatPath(base, quote)
    Else
        path = xFindCryptoPath(base, quote)
    End If
    
    EvaluatePath = IIf(path = "", STR_NA, Mid(path, 2))
End Function

Public Function EvaluateCost(ByVal ticker As String, ByVal amount As Variant) As Variant
    Assert_Check Not m_prices Is Nothing

    EvaluateCost = Empty
    
    If amount = CDec(0) Then Exit Function

    If Not m_paths.Exists(ticker) Then Exit Function

    If IsEmpty(m_paths(ticker)) Then Exit Function

    EvaluateCost = xEvalCost(amount, m_paths(ticker))
End Function

Private Sub xInitQuote(ByVal quote As String, ByVal paths As Dictionary)
    Set m_paths = IIf(paths Is Nothing, New Dictionary, paths)
    
    Set m_quote = TableEvaluation.FindQuoteInfo(quote)
    
    If m_quote Is Nothing Then
        Set m_quote = New PriceQuote
        
        Dim coin As BinanceCoin
        Set coin = TableCurrencies.GetTickerInfo(quote)
        
        If coin Is Nothing Then
            m_quote.ticker = quote
            m_quote.isFiat = False
        Else
            m_quote.ticker = quote
            m_quote.isFiat = coin.isFiat
            m_quote.format = CalculateNumberFormat(coin.isFiat, coin.precision)
        End If
    End If
End Sub

Private Function xCalcSymbols() As Collection
    Dim ops As New Dictionary
    
    Dim ticker As Variant
    For Each ticker In m_paths.Keys
        Dim path As String
        path = m_paths(ticker)

        Dim parts() As String
        parts = Split(path, ",")
    
        Dim i As Long
        For i = LBound(parts) To UBound(parts)
            Dim op, bq As String
            op = Left(parts(i), 1)
            bq = Mid(parts(i), 2)
            
            If op = "*" Or op = "/" Then
                If Not ops.Exists(bq) Then ops.Add bq, op
            End If
        Next i
    Next ticker

    Dim symbols As New Collection
    
    Dim symbol As Variant
    For Each ticker In ops.Keys
        symbols.Add ticker
    Next ticker

    Set xCalcSymbols = symbols
End Function

Private Function xGetPrices(Optional ByVal symbols As Collection = Nothing) As Dictionary
    Dim prices As New Dictionary
    
    Dim list As Collection
    Set list = BinanceAPI.SpotTrading_GetPrices(symbols)
    
    Dim item As Dictionary
    For Each item In list
        prices.Add item("symbol"), Str2Dec(item("price"))
    Next item
    
    Set xGetPrices = prices
End Function

Private Sub xCollectCommonQuotes()
    Set m_quotes = New Collection
    
    Dim table As ListObject
    Set table = ThisWorkbook.Sheets(SHEET_MISC).ListObjects(TABLE_COMMON_QUOTES)
    
    Dim i As Long
    For i = 1 To table.ListRows.count
        m_quotes.Add table.ListColumns(1).DataBodyRange(i).Value
    Next i
End Sub

Private Function xFindFiatPath(ByVal base As String, ByVal quote As String) As String
    Dim path As String
    If base = quote Then
        xFindFiatPath = ",="
    ElseIf Not m_infos.Exists(base) Then
        path = xFindFiatPath(m_mainFiatQuote, quote)
        If path = "" Then
            xFindFiatPath = ""
        ElseIf path = ",=" Then
            xFindFiatPath = ",$" & base & "/" & m_mainFiatQuote
        Else
            xFindFiatPath = ",$" & base & "/" & m_mainFiatQuote & path
        End If
    Else
        Select Case m_infos(base).method
            Case STR_EVAL_METHOD_MARKET, STR_EVAL_METHOD_EXCHANGE
                xFindFiatPath = xFindCryptoPath(base, quote)
            Case STR_EVAL_METHOD_FOREX
                path = xFindFiatPath(m_infos(base).refers, quote)
                If path = "" Then
                    xFindFiatPath = ""
                ElseIf path = ",=" Then
                    xFindFiatPath = ",$" & m_infos(base).refers & "/" & base
                Else
                    xFindFiatPath = ",$" & m_infos(base).refers & "/" & base & path
                End If
            Case STR_EVAL_METHOD_STABLECOIN
                xFindFiatPath = xFindCryptoPath(m_infos(base).refers, quote)
            Case Else
                Assert_Fail
        End Select
    End If
End Function

Private Function xFindCryptoPath(ByVal base As String, ByVal quote As String) As String
    Dim path As String
    
    If Not m_infos.Exists(quote) Then
        xFindCryptoPath = xFindMarketPath(base, quote)
    Else
        Select Case m_infos(quote).method
            Case STR_EVAL_METHOD_MARKET, STR_EVAL_METHOD_EXCHANGE
                xFindCryptoPath = xFindMarketPath(base, quote)
            Case STR_EVAL_METHOD_FOREX
                path = xFindCryptoPath(base, m_infos(quote).refers)
                If path = "" Then
                    xFindCryptoPath = ""
                ElseIf path = ",=" Then
                    xFindCryptoPath = ",$" & m_infos(quote).refers & "/" & quote
                Else
                    xFindCryptoPath = path & ",$" & m_infos(quote).refers & "/" & quote
                End If
            Case STR_EVAL_METHOD_STABLECOIN
                xFindCryptoPath = xFindMarketPath(base, m_infos(quote).refers)
            Case Else
                Assert_Fail
        End Select
    End If
End Function

Private Function xFindMarketPath(ByVal base As String, ByVal quote As String) As String
    If base = quote Then
        xFindMarketPath = ",="
        Exit Function
    ElseIf m_prices.Exists(base & quote) Then
        xFindMarketPath = ",*" & base & quote
        Exit Function
    ElseIf m_prices.Exists(quote & base) Then
        xFindMarketPath = ",/" & quote & base
        Exit Function
    End If

    Dim mediator As Variant
    For Each mediator In m_quotes
        If m_prices.Exists(base & mediator) And m_prices.Exists(mediator & quote) Then
            xFindMarketPath = ",*" & base & mediator & ",*" & mediator & quote
            Exit Function
        ElseIf m_prices.Exists(base & mediator) And m_prices.Exists(quote & mediator) Then
            xFindMarketPath = ",*" & base & mediator & ",/" & quote & mediator
            Exit Function
        ElseIf m_prices.Exists(mediator & base) And m_prices.Exists(quote & mediator) Then
            xFindMarketPath = ",/" & mediator & base & ",*" & quote & mediator
            Exit Function
        ElseIf m_prices.Exists(mediator & base) And m_prices.Exists(mediator & quote) Then
            xFindMarketPath = ",/" & mediator & base & ",/" & mediator & quote
            Exit Function
        End If
    Next mediator
    
    xFindMarketPath = ""
End Function

Private Function xEvalCost(ByVal amount As Variant, ByVal path As String) As Variant
    xEvalCost = Empty
    
    If path = STR_NA Then Exit Function
    
    Dim parts() As String
    parts = Split(path, ",")
    
    Dim op, symbol As String
    Dim cost, price As Variant
    
    cost = CDec(amount)
    
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        If parts(i) <> "=" Then
            op = Left(parts(i), 1)
            symbol = Mid(parts(i), 2)
            Select Case op
                Case "*"
                    If Not m_prices.Exists(symbol) Then Exit Function
                    price = m_prices(symbol)
                    cost = cost * price
                Case "/"
                    If Not m_prices.Exists(symbol) Then Exit Function
                    price = m_prices(symbol)
                    If price = 0 Then Exit Function
                    cost = cost / price
                Case "$"
                    Exit Function
                Case Else
                    Assert_Fail
            End Select
        End If
    Next i
    
    xEvalCost = Round(cost, 8)
End Function

