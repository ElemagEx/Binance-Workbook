Attribute VB_Name = "TableCurrencies"
Option Explicit
Option Private Module

Private Const SHEET_NAME As String = SHEET_CURRENCIES
Private Const TABLE_NAME As String = "Currencies"

Private Const COL_TICKER As String = "Ticker"
Private Const COL_NAME As String = "Name"
Private Const COL_TYPE As String = "Type"
Private Const COL_PRECISION As String = "Precision"
Private Const COL_BASES As String = "Bases"
Private Const COL_QUOTES As String = "Quotes"
Private Const COL_FORMAT As String = "Format"
Private Const COL_KEEP_ON_COMPACT As String = "Keep on Compact"
Private Const COL_TRADES As String = "Trades"
Private Const COL_SYMBOLS As String = "Symbols"

Private Const FIRST_TRADE_COL_INDEX As Long = 11

Public Property Get Tickers() As collection
    Set Tickers = New collection

    Dim col As ListColumn
    Set col = ThisWorkbook.Sheets(SHEET_NAME).ListObjects(TABLE_NAME).ListColumns(COL_TICKER)
    
    If Not col.DataBodyRange Is Nothing Then
        Dim i As Long
        For i = 1 To col.DataBodyRange.Count
            Tickers.Add col.DataBodyRange(i).Value
        Next i
    End If
End Property

Public Function IsDataCleanedUp()
    IsCleanedUp = True
End Function

Public Sub CleanUpData()
    ResetExchangeTimezone
    
    Dim table As ListObject
    Set table = ThisWorkbook.Sheets(SHEET_NAME).ListObjects(TABLE_NAME)
    
    Dim i As Long
    For i = table.ListColumns.Count To FIRST_TRADE_COL_INDEX Step -1
        table.ListColumns(i).Delete
    Next i
    
    If table.ListRows.Count > 0 Then
        For i = table.ListRows.Count To 1 Step -1
            If Not table.ListColumns(COL_KEEP_ON_COMPACT).DataBodyRange(i).Value Then
                table.ListRows(i).Delete
            ElseIf table.ListColumns(COL_FORMAT).DataBodyRange(i).NumberFormat = "General" Then
                table.ListRows(i).Delete
            Else
                table.ListColumns(COL_BASES).DataBodyRange(i).Value = 0
                table.ListColumns(COL_QUOTES).DataBodyRange(i).Value = 0
                table.ListColumns(COL_SYMBOLS).DataBodyRange(i).formula = "=0"
            End If
        Next i
    End If
End Sub

Public Sub UpdateData()
    xCollectData True, True
    
    Dim col As ListColumn
    Set col = ThisWorkbook.Sheets(SHEET_NAME).ListObjects(TABLE_NAME).ListColumns(COL_KEEP_ON_COMPACT)
    If Not col.DataBodyRange Is Nothing Then
        col.DataBodyRange.Value = True
    End If
End Sub

Public Sub CollectAllData()
    xCollectData False, False
End Sub

Public Sub CompactData()
    Dim table As ListObject
    Set table = ThisWorkbook.Sheets(SHEET_NAME).ListObjects(TABLE_NAME)
    
    Dim col As ListColumn
    Set col = table.ListColumns(COL_KEEP_ON_COMPACT)
    
    If Not col.DataBodyRange Is Nothing Then
        Dim i As Long
        For i = col.DataBodyRange.Count To 1 Step -1
            If Not col.DataBodyRange(i).Value Then
                table.ListRows(i).Delete
            End If
        Next i
    End If
End Sub

Private Sub xCollectData(ByVal addSelfTickers As Boolean, ByVal addWalletTickers)
    Dim coins As New BinanceCoins
    Dim c As BinanceCoin
    'coins.AddTicker "ABC"
    Set c = coins("ABC")

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
    
    Dim table As ListObject
    Set table = ThisWorkbook.Sheets(SHEET_NAME).ListObjects(TABLE_NAME)
    
    Dim col As ListColumn
    Set col = table.ListColumns(COL_FORMAT)
    
    ResetExchangeTimezone coins.ExchangeTimezone
    
    For Each ticker In coins.Tickers
        Dim coin As BinanceCoin
        Set coin = coins(ticker)
        
        Dim rowIndex As Long
        rowIndex = xFindTickerRowIndex(ticker, True)
        
        table.ListColumns(COL_NAME).DataBodyRange.Value = coin.name
        table.ListColumns(COL_TYPE).DataBodyRange.Value = IIf(coin.isFiat, STR_FIAT, STR_CRYPTO)
        table.ListColumns(COL_BASES).DataBodyRange.Value = coin.bases
        table.ListColumns(COL_QUOTES).DataBodyRange.Value = coin.quotes
        table.ListColumns(COL_PRECISION).DataBodyRange.Value = coin.precision
    Next ticker
End Sub

Private Function xFindTickerRowIndex(ByVal ticker As String, ByVal addIfNotFound As Boolean) As Long
    Dim table As ListObject
    Set table = ThisWorkbook.Sheets(SHEET_NAME).ListObjects(TABLE_NAME)
    
    Dim cell As range
    If table.ListRows.Count > 0 Then
        Set cell = table.ListColumns(COL_TICKER).DataBodyRange.Find(what:=ticker, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    End If
    
    Dim index As Long
    
    If Not cell Is Nothing Then
        index = cell.row - table.HeaderRowRange.row
    ElseIf Not addIfNotFound Then
        index = 0
    Else
        index = table.ListRows.Add().index
        table.ListColumns(COL_TICKER).DataBodyRange(index).Value = ticker
        table.ListColumns(COL_FORMAT).DataBodyRange(index).NumberFormat = "General"
        table.ListColumns(COL_SYMBOLS).DataBodyRange(index).formula = "=0"
    End If
    
    xFindTickerRowIndex = index
End Function

