Attribute VB_Name = "TableCurrencies"
Option Explicit
Option Private Module

Private Const COL_TICKER As String = "Ticker"
Private Const COL_NAME As String = "Name"
Private Const COL_TYPE As String = "Type"
Private Const COL_PRECISION As String = "Precision"
Private Const COL_BASES As String = "Bases"
Private Const COL_QUOTES As String = "Quotes"
Private Const COL_FORMAT As String = "Format"
Private Const COL_KEEP_ON_COMPACT As String = "Keep on Compact"
Private Const COL_MARKETS As String = "Markets"

Private Const FIRST_TRADE_COL_INDEX As Long = 11

Public Property Get Tickers() As Collection
    Set Tickers = New Collection

    Dim col As ListColumn
    Set col = xGetTable().ListColumns(COL_TICKER)
    
    If Not col.DataBodyRange Is Nothing Then
        Dim i As Long
        For i = 1 To col.DataBodyRange.count
            Tickers.Add col.DataBodyRange(i).Value
        Next i
    End If
End Property

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
    coin.format = table.ListColumns(COL_FORMAT).DataBodyRange(rowIndex).NumberFormat
    
    Dim colIndex As Long
    For colIndex = FIRST_TRADE_COL_INDEX To table.ListColumns.count
        Dim col As ListColumn
        Set col = table.ListColumns(colIndex)
        If VarType(col.DataBodyRange(rowIndex).Value) = vbBoolean Then
            coin.markets.Add col.name, col.DataBodyRange(rowIndex).Value
        End If
    Next colIndex
    
    Set GetTickerInfo = coin
End Function

Public Function IsDataCleanedUp()
    IsDataCleanedUp = True
End Function

Public Sub CleanUpData()
    ResetExchangeTimezone
    
    Dim table As ListObject
    Set table = xGetTable()
    
    Dim i As Long
    For i = table.ListColumns.count To FIRST_TRADE_COL_INDEX Step -1
        table.ListColumns(i).Delete
    Next i
    
    If table.ListRows.count > 0 Then
        For i = table.ListRows.count To 1 Step -1
            If Not table.ListColumns(COL_KEEP_ON_COMPACT).DataBodyRange(i).Value Then
                table.ListRows(i).Delete
            ElseIf table.ListColumns(COL_FORMAT).DataBodyRange(i).NumberFormat = "General" Then
                table.ListRows(i).Delete
            Else
                table.ListColumns(COL_BASES).DataBodyRange(i).Value = 0
                table.ListColumns(COL_QUOTES).DataBodyRange(i).Value = 0
                table.ListColumns(COL_MARKETS).DataBodyRange(i).formula = "=0"
            End If
        Next i
    End If
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
    For i = table.ListColumns.count To FIRST_TRADE_COL_INDEX Step -1
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
    
    Dim formula As String
    If table.ListColumns.count < FIRST_TRADE_COL_INDEX Then
        formula = "=0"
    Else
        formula = "=COUNTIF(" & TABLE_NAME & "[@[" & table.ListColumns(FIRST_TRADE_COL_INDEX).name & "]:[" & table.ListColumns(table.ListColumns.count).name & "]],TRUE)"
    End If

    table.ListColumns(COL_MARKETS).DataBodyRange.formula = formula
End Sub

Public Function xGetTable() As ListObject
    Set xGetTable = ThisWorkbook.Sheets(SHEET_CURRENCIES).ListObjects(TABLE_CURRENCIES)
End Function

Private Sub xCollectData(ByVal addSelfTickers As Boolean, ByVal addWalletTickers)
    Dim coins As New BinanceCoins

    Dim ticker As Variant
    If addSelfTickers Then
        For Each ticker In TableCurrencies.Tickers
            coins.AddTicker ticker
        Next ticker
    End If
    If addWalletTickers Then
        For Each ticker In TableWallet.Tickers
            coins.AddTicker ticker
        Next ticker
    End If
    
    coins.collectBasicInfo (addSelfTickers Or addWalletTickers)
    coins.collectExchangeInfo
    coins.RemoveUncollected
    
    Dim table As ListObject
    Set table = xGetTable()
    
    ResetExchangeTimezone coins.ExchangeTimezone
    
    Dim colIndex As Long
    Dim rowIndex As Long
    Dim coin As BinanceCoin
    
    For Each ticker In coins.Tickers
        Set coin = coins.item(ticker)
        
        rowIndex = xFindTickerRowIndex(ticker, True)
        
        table.ListColumns(COL_NAME).DataBodyRange(rowIndex).Value = coin.name
        table.ListColumns(COL_TYPE).DataBodyRange(rowIndex).Value = IIf(coin.isFiat, STR_FIAT, STR_CRYPTO)
        table.ListColumns(COL_PRECISION).DataBodyRange(rowIndex).Value = coin.precision
    Next ticker

    For Each ticker In coins.Tickers
        Set coin = coins.item(ticker)
        
        If coin.quotes > 0 Then
            For colIndex = FIRST_TRADE_COL_INDEX To table.ListColumns.count
                If ticker = table.ListColumns(colIndex).name Then
                    Exit For
                End If
            Next colIndex
            If colIndex > table.ListColumns.count Then
                table.ListColumns.Add(table.ListColumns.count + 1).name = ticker
            End If
        End If
    Next ticker
    
    Dim formula As String
    If table.ListColumns.count < FIRST_TRADE_COL_INDEX Then
        formula = "=0"
    Else
        formula = "=COUNTIF(" & TABLE_NAME & "[@[" & table.ListColumns(FIRST_TRADE_COL_INDEX).name & "]:[" & table.ListColumns(table.ListColumns.count).name & "]],TRUE)"
    End If

    For rowIndex = 1 To table.ListRows.count
        ticker = table.ListColumns(COL_TICKER).DataBodyRange(rowIndex).Value
        
        If Not coins.Contains(ticker) Then
            table.ListColumns(COL_BASES).DataBodyRange(rowIndex).Value = 0
            table.ListColumns(COL_QUOTES).DataBodyRange(rowIndex).Value = 0
            table.ListColumns(COL_MARKETS).DataBodyRange(rowIndex).formula = formula
        
            For colIndex = FIRST_TRADE_COL_INDEX To table.ListColumns.count
                table.ListColumns(colIndex).DataBodyRange(rowIndex).Value = " "
            Next colIndex
        Else
            Set coin = coins.item(ticker)
            
            For colIndex = FIRST_TRADE_COL_INDEX To table.ListColumns.count
                Dim col As ListColumn
                Set col = table.ListColumns(colIndex)
                If Not coin.markets.Exists(col.name) Then
                    col.DataBodyRange(rowIndex).Value = " "
                End If
            Next colIndex
        
            table.ListColumns(COL_BASES).DataBodyRange(rowIndex).Value = coin.bases
            table.ListColumns(COL_QUOTES).DataBodyRange(rowIndex).Value = coin.quotes
            table.ListColumns(COL_MARKETS).DataBodyRange(rowIndex).formula = formula
        End If
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
        table.ListColumns(COL_FORMAT).DataBodyRange(rowIndex).NumberFormat = "General"
    End If
    
    xFindTickerRowIndex = rowIndex
End Function

