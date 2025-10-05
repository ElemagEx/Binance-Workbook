Attribute VB_Name = "TableEvaluation"
Option Explicit
Option Private Module

Private Const COL_HEADER As String = "Ticker"
Private Const ROW_METHOD As Long = 1
Private Const ROW_REFERS As Long = 2
Private Const ROW_FORMAT As Long = 3

Public Sub RefreshEvalTickers()
    Dim table As ListObject
    Set table = xGetTable()
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_MISC)
    
    Dim list As Range
    Set list = ws.Range(table.HeaderRowRange(2), table.HeaderRowRange(table.ListColumns.count))
    
    Dim refs As String
    refs = "=" & ws.name & "!" & list.Address
    
    Dim namedRange As name
    For Each namedRange In ThisWorkbook.Names
        If namedRange.name = LIST_EVAL_TICKERS Then
            namedRange.Delete
            Exit For
        End If
    Next namedRange
    
    ThisWorkbook.Names.Add LIST_EVAL_TICKERS, refs, False
End Sub

Public Function GetInfo() As Dictionary
    Dim info As New Dictionary
        
    Set GetInfo = info
        
    Dim table As ListObject
    Set table = xGetTable()
    
    Dim i As Long
    For i = 2 To table.ListColumns.count
        Dim col As ListColumn
        Set col = table.ListColumns(i)
        
        Dim ticker As String
        ticker = col.name
        
        info.Add ticker, New Dictionary
        
        info(ticker).Add KEY_EVAL_TICKER, ticker
        info(ticker).Add KEY_EVAL_METHOD, col.DataBodyRange(ROW_METHOD).Value
        info(ticker).Add KEY_EVAL_REFERS, col.DataBodyRange(ROW_REFERS).Value
        info(ticker).Add KEY_EVAL_FORMAT, col.DataBodyRange(ROW_FORMAT).NumberFormat
    Next i
End Function

Public Function GetPrices(Optional ByVal symbols As Collection = Nothing) As Dictionary
    Dim prices As New Dictionary
    
    Dim list As Collection
    Set list = BinanceAPI.SpotTrading_GetPrices(symbols)
    
    Dim item As Dictionary
    For Each item In list
        prices.Add item("symbol"), Str2Dec(item("price"))
    Next item
    
    Set GetPrices = prices
End Function

Public Sub FillEvalPaths(ByVal cryptoTickers As Dictionary, ByVal fiatTickers As Dictionary)
    Dim prices As Dictionary
    Set prices = GetPrices()
    
    Dim info As Dictionary
    Set info = GetInfo()
    
    Dim quotes As Collection
    Set quotes = xCollectCommonQuotes()
    
    Dim path As String
    Dim ticker, quote As Variant
    For Each ticker In cryptoTickers.Keys
        For Each quote In info.Keys
            path = xFindCryptoPath(ticker, quote, quotes, prices, info)
            cryptoTickers(ticker).Add quote, IIf(Len(path) = 0, STR_NA, Mid(path, 2))
        Next quote
    Next ticker

    Dim defaultFiatTicker As String
    defaultFiatTicker = xGetDefaultFiatTicker()
    
    For Each ticker In fiatTickers.Keys
        For Each quote In info.Keys
            path = xFindFiatPath(defaultFiatTicker, ticker, quote, quotes, prices, info)
            fiatTickers(ticker).Add quote, IIf(Len(path) = 0, STR_NA, Mid(path, 2))
        Next quote
    Next ticker
End Sub

Private Function xGetTable() As ListObject
    Set xGetTable = ThisWorkbook.Sheets(SHEET_MISC).ListObjects(TABLE_EVALUATION)
End Function

Private Function xCollectCommonQuotes() As Collection
    Dim quotes As New Collection
    
    Dim table As ListObject
    Set table = ThisWorkbook.Sheets(SHEET_MISC).ListObjects(TABLE_COMMON_QUOTES)
    
    Dim i As Long
    For i = 1 To table.ListRows.count
        quotes.Add table.ListColumns(1).DataBodyRange(i).Value
    Next i
    
    Set xCollectCommonQuotes = quotes
End Function

Private Function xGetDefaultFiatTicker() As String
    Dim ticker As String
    
    ticker = ThisWorkbook.Sheets(SHEET_MISC).ListObjects(TABLE_DEAULT_FIAT).DataBodyRange(1).Value
    
    xGetDefaultFiatTicker = ticker
End Function

Private Function xFindFiatPath(ByVal fiat As String, ByVal ticker As String, ByVal quote As String, ByVal quotes As Collection, ByVal prices As Dictionary, ByVal info As Dictionary) As String
    Dim path As String
    If Not info.Exists(ticker) Then
        path = xFindFiatPath("", fiat, quote, quotes, prices, info)
        xFindFiatPath = IIf(Len(path) = 0, "", "," & "$" & ticker & "/" & fiat & path)
    Else
        Select Case info(ticker)(KEY_EVAL_METHOD)
            Case STR_EVAL_METHOD_MARKET
                path = xFindCryptoPath(ticker, quote, quotes, prices, info)
                xFindFiatPath = IIf(Len(path) = 0, "", "," & "!" & path)
            Case STR_EVAL_METHOD_FOREX
                path = xFindFiatPath(fiat, info(ticker)(KEY_EVAL_REFERS), quote, quotes, prices, info)
                xFindFiatPath = IIf(Len(path) = 0, "", "," & "$" & info(ticker)(KEY_EVAL_REFERS) & "/" & ticker & path)
                Exit Function
            Case STR_EVAL_METHOD_STABLECOIN
                path = xFindCryptoPath(info(ticker)(KEY_EVAL_REFERS), quote, quotes, prices, info)
                xFindFiatPath = IIf(Len(path) = 0, "", "," & "!" & path)
                Exit Function
            Case Else
                Assert_Fail
        End Select
    End If
End Function

Private Function xFindCryptoPath(ByVal ticker As String, ByVal quote As String, ByVal quotes As Collection, ByVal prices As Dictionary, ByVal info As Dictionary) As String
    Dim path As String
    
    If Not info.Exists(quote) Then
        xFindCryptoPath = xFindMarketPath(ticker, quote, quotes, prices)
    Else
        Select Case info(quote)(KEY_EVAL_METHOD)
            Case STR_EVAL_METHOD_MARKET
                path = xFindMarketPath(ticker, quote, quotes, prices)
                xFindCryptoPath = xFindCryptoPath = IIf(Len(path) = 0, "", path & "," & "!")
            Case STR_EVAL_METHOD_FOREX
                path = xFindCryptoPath(ticker, info(quote)(KEY_EVAL_REFERS), quotes, prices, info)
                xFindCryptoPath = IIf(Len(path) = 0, "", path & "," & "$" & info(quote)(KEY_EVAL_REFERS) & "/" & quote)
            Case STR_EVAL_METHOD_STABLECOIN
                path = xFindMarketPath(ticker, info(quote)(KEY_EVAL_REFERS), quotes, prices)
                xFindCryptoPath = IIf(Len(path) = 0, "", path & "," & "!")
            Case Else
                Assert_Fail
        End Select
    End If
End Function

Private Function xFindMarketPath(ByVal ticker As String, ByVal quote As String, ByVal quotes As Collection, ByVal prices As Dictionary) As String
    If ticker = quote Then
        xFindMarketPath = "," & "!"
        Exit Function
    ElseIf prices.Exists(ticker & quote) Then
        xFindMarketPath = "," & "*" & ticker & quote
        Exit Function
    ElseIf prices.Exists(quote & ticker) Then
        xFindMarketPath = "," & "/" & quote & ticker
        Exit Function
    End If

    Dim mediator As Variant
    For Each mediator In quotes
        If prices.Exists(ticker & mediator) And prices.Exists(quote & mediator) Then
            xFindMarketPath = "," & "*" & ticker & mediator & "," & "/" & quote & mediator
            Exit Function
        ElseIf prices.Exists(ticker & mediator) And prices.Exists(mediator & quote) Then
            xFindMarketPath = "," & "*" & ticker & mediator & "," & "*" & mediator & quote
            Exit Function
        ElseIf prices.Exists(mediator & ticker) And prices.Exists(quote & mediator) Then
            xFindMarketPath = "," & "/" & mediator & ticker & "," & "*" & quote & mediator
            Exit Function
        ElseIf prices.Exists(mediator & ticker) And prices.Exists(mediator & quote) Then
            xFindMarketPath = "," & "/" & mediator & ticker & "," & "/" & mediator & quote
            Exit Function
        End If
    Next mediator
    
    xFindMarketPath = ""
End Function

