Attribute VB_Name = "TableCurrencies"
Option Explicit
Option Private Module

Private Const COL_TICKER As String = "Ticker"
Private Const COL_NAME As String = "Name"
Private Const COL_TYPE As String = "Type"
Private Const COL_PRECISION As String = "Precision"
Private Const COL_BASES As String = "Bases"
Private Const COL_QUOTES As String = "Quotes"
Private Const COL_KEEP_ON_COMPACT As String = "Keep on Compact"
Private Const COL_MARKETS As String = "Markets"
Private Const COL_TRADES As String = "Trades"

Private Const COL_INSERT_MARKET_AFTER As String = COL_KEEP_ON_COMPACT
Private Const COL_INSERT_MARKET_BEFORE As String = COL_MARKETS
Private Const COL_INSERT_EVALUATION_AFTER As String = COL_MARKETS
Private Const COL_INSERT_EVALUATION_BEFORE As String = COL_TRADES

Private Const EVAL_PREFIX As String = "eval-"

Public Property Get tickers() As Collection
    Set tickers = New Collection

    Dim col As ListColumn
    Set col = xGetTable().ListColumns(COL_TICKER)
    
    If Not col.DataBodyRange Is Nothing Then
        Dim i As Long
        For i = 1 To col.DataBodyRange.count
            tickers.Add col.DataBodyRange(i).Value
        Next i
    End If
End Property

Public Function GetTradesLastUpdate(ByVal ticker As String, ByVal def As Date) As Date
    GetTradesLastUpdate = def

    Dim rowIndex As Long
    rowIndex = xFindTickerRowIndex(ticker, False)
    
    If rowIndex = 0 Then Exit Function
    
    Dim val As Variant
    val = xGetTable().ListColumns(COL_TRADES).DataBodyRange(rowIndex)
    
    If IsEmpty(val) Then Exit Function
    
    GetTradesLastUpdate = val
End Function

Public Sub SetTradesLastUpdate(ByVal ticker As String, ByVal val As Date)
    Dim rowIndex As Long
    rowIndex = xFindTickerRowIndex(ticker, False)
    
    If rowIndex = 0 Then Exit Sub
    
    xGetTable().ListColumns(COL_TRADES).DataBodyRange(rowIndex).Value = IIf(val = 0, Empty, val)
End Sub

Public Function GetSelectedTicker() As String
    GetSelectedTicker = ""
    
    If ActiveSheet.name = SHEET_CURRENCIES Then
        Dim table As ListObject
        Set table = xGetTable()

        If TypeName(Selection) = "Range" Then
            Dim rowIndex As Long
            rowIndex = Selection.Row - table.HeaderRowRange.Row
            
            If rowIndex >= 1 And rowIndex <= table.ListRows.count Then
                GetSelectedTicker = table.ListColumns(COL_TICKER).DataBodyRange(rowIndex)
            End If
        End If
    End If
End Function

Public Function GetTickerInfo(ByVal ticker As String) As BinanceCoin
    Set GetTickerInfo = Nothing

    Dim rowIndex As Long
    rowIndex = xFindTickerRowIndex(ticker, False)
    
    If rowIndex = 0 Then
        Exit Function
    End If

    Dim table As ListObject
    Set table = xGetTable()
    
    Dim coin As New BinanceCoin
    coin.ticker = table.ListColumns(COL_TICKER).DataBodyRange(rowIndex).Value
    coin.name = table.ListColumns(COL_NAME).DataBodyRange(rowIndex).Value
    coin.isFiat = (table.ListColumns(COL_TYPE).DataBodyRange(rowIndex).Value = STR_FIAT)
    coin.bases = table.ListColumns(COL_BASES).DataBodyRange(rowIndex).Value
    coin.quotes = table.ListColumns(COL_QUOTES).DataBodyRange(rowIndex).Value
    coin.precision = table.ListColumns(COL_PRECISION).DataBodyRange(rowIndex).Value
    
    Dim colIndexFirst, colIndexLast As Long
    colIndexFirst = table.ListColumns(COL_INSERT_MARKET_AFTER).index + 1
    colIndexLast = table.ListColumns(COL_INSERT_MARKET_BEFORE).index - 1
    
    Dim colIndex As Long
    For colIndex = colIndexFirst To colIndexLast
        Dim col As ListColumn
        Set col = table.ListColumns(colIndex)
        If VarType(col.DataBodyRange(rowIndex).Value) = vbBoolean Then
            coin.markets.Add col.name, col.DataBodyRange(rowIndex).Value
        End If
    Next colIndex
    
    Set GetTickerInfo = coin
End Function

Public Function IsDataCleanedUp()
    IsDataCleanedUp = False
    
    Dim table As ListObject
    Set table = xGetTable()
    
    Dim colIndexFirst, colIndexLast As Long
    colIndexFirst = table.ListColumns(COL_INSERT_MARKET_AFTER).index + 1
    colIndexLast = table.ListColumns(COL_INSERT_MARKET_BEFORE).index - 1
    
    If colIndexFirst <= colIndexLast Then Exit Function

    colIndexFirst = table.ListColumns(COL_INSERT_EVALUATION_AFTER).index + 1
    colIndexLast = table.ListColumns(COL_INSERT_EVALUATION_BEFORE).index - 1
    
    If colIndexFirst <= colIndexLast Then Exit Function

    Dim infos As Dictionary
    Set infos = TableEvaluation.GetInfos()
    
    Dim i As Long
    If table.ListRows.count > 0 Then
        For i = 1 To table.ListRows.count
            If Not infos.Exists(table.ListColumns(COL_TICKER).DataBodyRange(i).Value) Then
                Exit Function
            End If
            If Not IsEmpty(table.ListColumns(COL_TRADES).DataBodyRange(i).Value) Then
                Exit Function
            End If
        Next i
    End If
    
    IsDataCleanedUp = True
End Function

Public Sub CleanUpData()
    Dim f As Boolean
    f = IsDataCleanedUp()

    ResetExchangeTimezone
    
    Dim table As ListObject
    Set table = xGetTable()
    
    Dim i As Long
    
    Dim colIndexFirst, colIndexLast As Long
    colIndexFirst = table.ListColumns(COL_INSERT_EVALUATION_AFTER).index + 1
    colIndexLast = table.ListColumns(COL_INSERT_EVALUATION_BEFORE).index - 1
    
    For i = colIndexLast To colIndexFirst Step -1
        table.ListColumns(i).Delete
    Next i
    
    colIndexFirst = table.ListColumns(COL_INSERT_MARKET_AFTER).index + 1
    colIndexLast = table.ListColumns(COL_INSERT_MARKET_BEFORE).index - 1
    
    For i = colIndexLast To colIndexFirst Step -1
        table.ListColumns(i).Delete
    Next i
    
    Dim infos As Dictionary
    Set infos = TableEvaluation.GetInfos()
    
    If table.ListRows.count > 0 Then
        For i = table.ListRows.count To 1 Step -1
            If Not infos.Exists(table.ListColumns(COL_TICKER).DataBodyRange(i).Value) Then
                table.ListRows(i).Delete
            ElseIf Not table.ListColumns(COL_KEEP_ON_COMPACT).DataBodyRange(i) Then
                table.ListRows(i).Delete
            Else
                table.ListColumns(COL_BASES).DataBodyRange(i).Value = 0
                table.ListColumns(COL_QUOTES).DataBodyRange(i).Value = 0
                table.ListColumns(COL_MARKETS).DataBodyRange(i).formula = "=0"
            End If
        Next i
    End If
    
    table.ListColumns(COL_TRADES).DataBodyRange.ClearContents
End Sub

Public Sub UpdateData()
    xCollectData True, True
    
    Dim col As ListColumn
    Set col = xGetTable().ListColumns(COL_KEEP_ON_COMPACT)
    If Not col.DataBodyRange Is Nothing Then
        col.DataBodyRange.Value = True
    End If
End Sub

Public Sub CollectAllData()
    xCollectData False, False
End Sub

Public Sub CompactData()
    Dim table As ListObject
    Set table = xGetTable()
    
    Dim col As ListColumn
    Set col = table.ListColumns(COL_KEEP_ON_COMPACT)
    
    Dim i As Long
    
    If Not col.DataBodyRange Is Nothing Then
        For i = col.DataBodyRange.count To 1 Step -1
            If Not col.DataBodyRange(i).Value Then
                table.ListRows(i).Delete
            End If
        Next i
    End If
    
    Dim colIndexFirst, colIndexLast As Long
    colIndexFirst = table.ListColumns(COL_INSERT_MARKET_AFTER).index + 1
    colIndexLast = table.ListColumns(COL_INSERT_MARKET_BEFORE).index - 1
    
    For i = colIndexLast To colIndexFirst Step -1
        Set col = table.ListColumns(i)
    
        Dim cell As Range
        Set cell = table.ListColumns(COL_TICKER).DataBodyRange.Find(what:=col.name, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
        
        If cell Is Nothing Then
            col.Delete
        ElseIf WorksheetFunction.CountIf(col.DataBodyRange, True) > 0 Then
        ElseIf WorksheetFunction.CountIf(col.DataBodyRange, False) > 0 Then
        Else
            col.Delete
        End If
    Next i
    
    colIndexFirst = table.ListColumns(COL_INSERT_MARKET_AFTER).index + 1
    colIndexLast = table.ListColumns(COL_INSERT_MARKET_BEFORE).index - 1
    
    Dim formula As String
    If colIndexFirst > colIndexLast Then
        formula = "=0"
    Else
        formula = "=COUNTIF(" & TABLE_CURRENCIES & "[@[" & table.ListColumns(colIndexFirst).name & "]:[" & table.ListColumns(colIndexLast).name & "]],TRUE)"
    End If

    table.ListColumns(COL_MARKETS).DataBodyRange.formula = formula
End Sub

Public Sub CollectEvalPaths(ByVal paths As Dictionary, ByVal quote As String)
    Dim table As ListObject
    Set table = xGetTable()
    
    Dim colIndexFirst, colIndexLast As Long
    
    colIndexFirst = table.ListColumns(COL_INSERT_EVALUATION_AFTER).index + 1
    colIndexLast = table.ListColumns(COL_INSERT_EVALUATION_BEFORE).index - 1

    Dim colIndex As Long
    For colIndex = colIndexFirst To colIndexLast
        If table.ListColumns(colIndex).name = EVAL_PREFIX & quote Then
            Exit For
        End If
    Next colIndex
    
    If colIndex > colIndexLast Then Exit Sub
    
    Dim rowIndex As Long
    Dim ticker As Variant
    For Each ticker In paths.Keys
        rowIndex = xFindTickerRowIndex(ticker, False)
        
        If rowIndex > 0 Then
            paths.item(ticker) = table.ListColumns(colIndex).DataBodyRange(rowIndex).Value
        End If
    Next ticker
End Sub

Public Function xGetTable() As ListObject
    Set xGetTable = ThisWorkbook.Sheets(SHEET_CURRENCIES).ListObjects(TABLE_CURRENCIES)
End Function

Private Sub xCollectData(ByVal addSelfTickers As Boolean, ByVal addWalletTickers)
    Dim coins As New BinanceCoins
    Dim evaluator As New PriceEvaluator
    '
    ' Choosing tickers to collect
    '
    Dim ticker, quote As Variant
    If addSelfTickers Then
        For Each ticker In TableCurrencies.tickers
            coins.AddTicker ticker
        Next ticker
    End If
    If addWalletTickers Then
        For Each ticker In TableAssets.tickers
            coins.AddTicker ticker
        Next ticker
    End If
    '
    ' Collect data from Binance
    '
    coins.collectBasicInfo (addSelfTickers Or addWalletTickers)
    coins.collectExchangeInfo
    coins.RemoveUncollected
    
    evaluator.GetAllEvalSymbols
    
    Dim table As ListObject
    Set table = xGetTable()
    
    ResetExchangeTimezone coins.ExchangeTimezone
    
    Dim colIndex As Long
    Dim rowIndex As Long
    Dim coin As BinanceCoin
    '
    ' Add all collect currencies to table
    '
    For Each ticker In coins.tickers
        Set coin = coins.item(ticker)
        
        rowIndex = xFindTickerRowIndex(ticker, True)
        
        table.ListColumns(COL_NAME).DataBodyRange(rowIndex).Value = coin.name
        table.ListColumns(COL_TYPE).DataBodyRange(rowIndex).Value = IIf(coin.isFiat, STR_FIAT, STR_CRYPTO)
        table.ListColumns(COL_PRECISION).DataBodyRange(rowIndex).Value = coin.precision
    Next ticker

    Dim colIndexFirst, colIndexLast As Long
    '
    ' Add all new markets colums
    '
    For Each ticker In coins.tickers
        Set coin = coins.item(ticker)
        
        If coin.quotes > 0 Then
            colIndexFirst = table.ListColumns(COL_INSERT_MARKET_AFTER).index + 1
            colIndexLast = table.ListColumns(COL_INSERT_MARKET_BEFORE).index - 1
            
            For colIndex = colIndexFirst To colIndexLast
                If ticker = table.ListColumns(colIndex).name Then
                    Exit For
                End If
            Next colIndex
            If colIndex > colIndexLast Then
                table.ListColumns.Add(colIndexLast + 1).name = ticker
            End If
        End If
    Next ticker
    '
    ' Add missing evaluation tickers columns
    '
    Dim infos As Dictionary
    Set infos = evaluator.GetInfos()
    
    For Each quote In infos.Keys
        Dim name As String
        name = EVAL_PREFIX & quote
        
        colIndexFirst = table.ListColumns(COL_INSERT_EVALUATION_AFTER).index + 1
        colIndexLast = table.ListColumns(COL_INSERT_EVALUATION_BEFORE).index - 1

        For colIndex = colIndexFirst To colIndexLast
            If name = table.ListColumns(colIndex).name Then
                Exit For
            End If
        Next colIndex
        If colIndex > colIndexLast Then
            table.ListColumns.Add(colIndexLast + 1).name = name
        End If
    Next quote
    
    '
    ' Calc markets formula
    '
    colIndexFirst = table.ListColumns(COL_INSERT_MARKET_AFTER).index + 1
    colIndexLast = table.ListColumns(COL_INSERT_MARKET_BEFORE).index - 1
    
    Dim formula As String
    If colIndexFirst > colIndexLast Then
        formula = "=0"
    Else
        formula = "=COUNTIF(" & TABLE_CURRENCIES & "[@[" & table.ListColumns(colIndexFirst).name & "]:[" & table.ListColumns(colIndexLast).name & "]],TRUE)"
    End If
    '
    ' Set all markets availability for all currencies and set all evaluation paths
    '
    For rowIndex = 1 To table.ListRows.count
        ticker = table.ListColumns(COL_TICKER).DataBodyRange(rowIndex).Value
        
        Dim isFiat As Boolean
        
        If Not coins.Contains(ticker) Then
            table.ListColumns(COL_BASES).DataBodyRange(rowIndex).Value = 0
            table.ListColumns(COL_QUOTES).DataBodyRange(rowIndex).Value = 0
            table.ListColumns(COL_MARKETS).DataBodyRange(rowIndex).formula = formula
        
            For colIndex = colIndexFirst To colIndexLast
                table.ListColumns(colIndex).DataBodyRange(rowIndex).Value = " "
            Next colIndex
            
            isFiat = (table.ListColumns(COL_TYPE).DataBodyRange(rowIndex).Value = STR_FIAT)
        Else
            Set coin = coins.item(ticker)
            
            For colIndex = colIndexFirst To colIndexLast
                Dim col As ListColumn
                Set col = table.ListColumns(colIndex)
                If Not coin.markets.Exists(col.name) Then
                    col.DataBodyRange(rowIndex).Value = " "
                End If
            Next colIndex
        
            table.ListColumns(COL_BASES).DataBodyRange(rowIndex).Value = coin.bases
            table.ListColumns(COL_QUOTES).DataBodyRange(rowIndex).Value = coin.quotes
            table.ListColumns(COL_MARKETS).DataBodyRange(rowIndex).formula = formula
            
            isFiat = coin.isFiat
        End If
        
        For Each quote In infos.Keys
            Dim path As String
            path = evaluator.EvaluatePath(isFiat, ticker, quote)
            table.ListColumns(EVAL_PREFIX & quote).DataBodyRange(rowIndex).Value = path
        Next quote
    Next rowIndex
End Sub

Private Function xFindTickerRowIndex(ByVal ticker As String, ByVal addIfNotFound As Boolean) As Long
    Dim table As ListObject
    Set table = xGetTable()
    
    Dim cell As Range
    If table.ListRows.count > 0 Then
        Set cell = table.ListColumns(COL_TICKER).DataBodyRange.Find(what:=ticker, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    End If
    
    Dim rowIndex As Long
    
    If Not cell Is Nothing Then
        rowIndex = cell.Row - table.HeaderRowRange.Row
    ElseIf Not addIfNotFound Then
        rowIndex = 0
    Else
        rowIndex = table.ListRows.Add().index
        table.ListColumns(COL_TICKER).DataBodyRange(rowIndex).Value = ticker
    End If
    
    xFindTickerRowIndex = rowIndex
End Function

