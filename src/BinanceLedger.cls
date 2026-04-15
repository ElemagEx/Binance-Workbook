VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BinanceLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const COL_TRIVIA_PROPERTY As Long = 1
Private Const COL_TRIVIA_VALUE As Long = 2

Private Const ROW_TRIVIA_TICKER As Long = 1
Private Const ROW_TRIVIA_NAME As Long = 2
Private Const ROW_TRIVIA_TYPE As Long = 3
Private Const ROW_TRIVIA_QUOTE As Long = 4
Private Const ROW_TRIVIA_PRICE As Long = 5
Private Const ROW_TRIVIA_COST As Long = 6
Private Const ROW_TRIVIA_TOTAL As Long = 7
Private Const ROW_TRIVIA_BALANCE As Long = 8
Private Const ROW_TRIVIA_DYNAMIC As Long = 9
Private Const ROW_TRIVIA_DIVIATION As Long = 10
Private Const ROW_TRIVIA_ACTIVITY_START_DATE As Long = 11
Private Const ROW_TRIVIA_KLINES_INTERVAL As Long = 12

Private Const COL_SUMMARY_ACQUIRED As String = "Acquired"
Private Const COL_SUMMARY_SPENT As String = "Spent"
Private Const COL_SUMMARY_CHARGE As String = "Charge"
Private Const COL_SUMMARY_BALANCE As String = "Balance"
Private Const COL_SUMMARY_BNB_FEES As String = "BNB fees"
Private Const COL_SUMMARY_DEBIT As String = "Debit"
Private Const COL_SUMMARY_CREDIT As String = "Credit"
Private Const COL_SUMMARY_COMMISSION As String = "Commission"
Private Const COL_SUMMARY_INVESTED As String = "Invested"
Private Const COL_SUMMARY_UNREALIZED_LP As String = "Unrealized L/P"

Private Const COL_MARKETS_TICKER As String = "Ticker"
Private Const COL_MARKETS_OBSERVED As String = "Observed"

Private Const COL_DETAILS_STAMP As String = "Stamp"
Private Const COL_DETAILS_WALLET As String = "Wallet"
Private Const COL_DETAILS_OPERATION As String = "Operation"
Private Const COL_DETAILS_ACQUIRED As String = "Acquired"
Private Const COL_DETAILS_SPENT As String = "Spent"
Private Const COL_DETAILS_TICKER As String = "Ticker"
Private Const COL_DETAILS_PRICE As String = "Price"
Private Const COL_DETAILS_AMOUNT As String = "Amount"
Private Const COL_DETAILS_CHARGE As String = "Charge"
Private Const COL_DETAILS_BNB_FEE As String = "BNB Fee"
Private Const COL_DETAILS_DEBIT As String = "Debit"
Private Const COL_DETAILS_CREDIT As String = "Credit"
Private Const COL_DETAILS_COMMISSION As String = "Commission"
Private Const COL_DETAILS_NOTE As String = "Note"
Private Const COL_DETAILS_ID As String = "ID"

Private m_ws As Worksheet
Private m_ticker As String

Public Property Get ActivityStartDate() As Date
    ActivityStartDate = m_ws.ListObjects(xTableName(TABLE_TRIVIA)).ListColumns(COL_TRIVIA_VALUE).DataBodyRange(ROW_TRIVIA_ACTIVITY_START_DATE).Value
End Property

Public Function GetTicker() As String
    GetTicker = m_ticker
End Function

Public Function GetQuote() As String
    GetQuote = m_ws.ListObjects(xTableName(TABLE_TRIVIA)).ListColumns(COL_TRIVIA_VALUE).DataBodyRange(ROW_TRIVIA_QUOTE).Value
End Function

Public Function IsValid() As Boolean
    IsValid = True
End Function

Public Sub Activate()
    m_ws.Activate
End Sub

Public Sub Init(ByVal ticker As String, ByVal ws As Worksheet)
    Set m_ws = ws
    m_ticker = ticker
End Sub

Public Sub Create(ByVal ticker As String, ByVal name As String)
    Dim wsLast As Worksheet
    Set wsLast = ThisWorkbook.Sheets(ThisWorkbook.Sheets.count())

    Dim wsTemplate As Worksheet
    Set wsTemplate = ThisWorkbook.Sheets(SHEET_LEDGER_TEMPLATE)

    wsTemplate.Copy After:=wsLast
    
    Set m_ws = ThisWorkbook.Sheets(ThisWorkbook.Sheets.count())
    
    m_ws.name = name
    m_ticker = ticker
    
    Dim table As ListObject
    For Each table In m_ws.ListObjects
        If InStr(table.name, TABLE_TRIVIA) > 0 Then
            table.name = xTableName(TABLE_TRIVIA)
        ElseIf InStr(table.name, TABLE_SUMMARY) > 0 Then
            table.name = xTableName(TABLE_SUMMARY)
        ElseIf InStr(table.name, TABLE_MARKETS) > 0 Then
            table.name = xTableName(TABLE_MARKETS)
        ElseIf InStr(table.name, TABLE_DETAILS) > 0 Then
            table.name = xTableName(TABLE_DETAILS)
        End If
    Next table

    xReset True
    ApplyMainQuote
End Sub

Public Sub RefreshCoinInfo()
    xReset False
End Sub

Public Sub ClearCosts()
    Dim info As PriceQuote
    Set info = TableEvaluation.FindQuoteInfo(GetQuote())
    
    Dim numberFormat As String
    If info Is Nothing Then
        numberFormat = NUMBER_FORMAT_MAIN_CRYPTO
    Else
        numberFormat = CalculateNumberFormat(info.isFiat)
    End If
    
    Dim table As ListObject
    
    Set table = m_ws.ListObjects(xTableName(TABLE_TRIVIA))
    
    table.ListColumns(COL_TRIVIA_VALUE).DataBodyRange(ROW_TRIVIA_COST).numberFormat = numberFormat
    
    Set table = m_ws.ListObjects(xTableName(TABLE_SUMMARY))
    
    table.ListColumns(COL_SUMMARY_DEBIT).DataBodyRange.numberFormat = numberFormat
    table.ListColumns(COL_SUMMARY_CREDIT).DataBodyRange.numberFormat = numberFormat
    table.ListColumns(COL_SUMMARY_COMMISSION).DataBodyRange.numberFormat = numberFormat
    table.ListColumns(COL_SUMMARY_INVESTED).DataBodyRange.numberFormat = numberFormat
    table.ListColumns(COL_SUMMARY_UNREALIZED_LP).DataBodyRange.numberFormat = numberFormat
    
    table.ListColumns(COL_SUMMARY_DEBIT).Total.numberFormat = numberFormat
    table.ListColumns(COL_SUMMARY_CREDIT).Total.numberFormat = numberFormat
    table.ListColumns(COL_SUMMARY_COMMISSION).Total.numberFormat = numberFormat
    table.ListColumns(COL_SUMMARY_INVESTED).Total.numberFormat = numberFormat
    table.ListColumns(COL_SUMMARY_UNREALIZED_LP).Total.numberFormat = numberFormat
    
    Set table = m_ws.ListObjects(xTableName(TABLE_DETAILS))

    Dim rowIndex As Long
    If table.ListRows.count = 0 Then rowIndex = table.ListRows.Add().index

    table.ListColumns(COL_DETAILS_DEBIT).DataBodyRange.numberFormat = numberFormat
    table.ListColumns(COL_DETAILS_CREDIT).DataBodyRange.numberFormat = numberFormat
    table.ListColumns(COL_DETAILS_COMMISSION).DataBodyRange.numberFormat = numberFormat
        
    table.ListColumns(COL_DETAILS_DEBIT).DataBodyRange.ClearContents
    table.ListColumns(COL_DETAILS_CREDIT).DataBodyRange.ClearContents
    table.ListColumns(COL_DETAILS_COMMISSION).DataBodyRange.ClearContents
    
    table.ListColumns(COL_DETAILS_DEBIT).Total.numberFormat = numberFormat
    table.ListColumns(COL_DETAILS_CREDIT).Total.numberFormat = numberFormat
    table.ListColumns(COL_DETAILS_COMMISSION).Total.numberFormat = numberFormat
    
    If rowIndex > 0 Then table.ListRows(rowIndex).Delete
End Sub

Public Sub ApplyMainQuote()
    m_ws.ListObjects(xTableName(TABLE_TRIVIA)).ListColumns(COL_TRIVIA_VALUE).DataBodyRange(ROW_TRIVIA_QUOTE).Value = TableUserInfo.quote
    ClearCosts
End Sub

Public Sub OnChange(ByVal Target As Range)
    Dim cells As Range
    Debug.Print "Tick"
    
    Set cells = m_ws.ListObjects(xTableName(TABLE_MARKETS)).DataBodyRange
    If Not Intersect(Target, cells) Is Nothing Then
        TableCurrencies.SetTradesLastUpdate m_ticker, 0
    End If
End Sub

Public Sub UpdateTrades()
    Dim markets As Dictionary
    Set markets = xCollectTradeMarkets()
    
    If markets.count = 0 Then
        MsgBox "Cannot find selected markets for this ticker: " & m_ticker
        Exit Sub
    End If
    
    Dim lastUpdateTime As Date
    lastUpdateTime = TableCurrencies.GetTradesLastUpdate(m_ticker, ActivityStartDate)
    
    Dim ops As New BinanceOps
    
    Dim quote As Variant
    For Each quote In markets.Keys
        ops.Collect_MySpotTrades lastUpdateTime, m_ticker, quote
    Next quote
    
    BinanceLedgers.PopulateOperations ops

    TableCurrencies.SetTradesLastUpdate m_ticker, ops.endDate
End Sub

Public Sub Evaluate()
    Dim quote As String
    quote = GetQuote()
    
    If quote = "" Then Exit Sub
    
    Dim periods As Dictionary
    Set periods = xCollectPeriods(quote)

    Dim evaluator As New PriceEvaluator
    evaluator.GetQuotePrices quote, Nothing, periods
    
    Dim table As ListObject
    Set table = m_ws.ListObjects(xTableName(TABLE_DETAILS))
    
    Dim rowIndex As Long
    For rowIndex = 1 To table.ListRows.count
        If _
            IsEmpty(table.ListColumns(COL_DETAILS_DEBIT).DataBodyRange(rowIndex).Value) Or _
            IsEmpty(table.ListColumns(COL_DETAILS_CREDIT).DataBodyRange(rowIndex).Value) _
        Then
            Dim stamp As Date
            stamp = table.ListColumns(COL_DETAILS_STAMP).DataBodyRange(rowIndex).Value
            
            Dim ticker As String
            ticker = table.ListColumns(COL_DETAILS_TICKER).DataBodyRange(rowIndex).Value
        
            Dim amount1, amount2 As Variant
            
            If ticker = "" Then
                amount1 = table.ListColumns(COL_DETAILS_ACQUIRED).DataBodyRange(rowIndex).Value
                amount1 = evaluator.EvaluateHistoricalCost(stamp, m_ticker, amount1)
            
                amount2 = table.ListColumns(COL_DETAILS_SPENT).DataBodyRange(rowIndex).Value
                amount2 = evaluator.EvaluateHistoricalCost(stamp, m_ticker, amount2)
                
                table.ListColumns(COL_DETAILS_DEBIT).DataBodyRange(rowIndex).Value = amount1
                table.ListColumns(COL_DETAILS_CREDIT).DataBodyRange(rowIndex).Value = amount2
            Else
                If IsInOp(table.ListColumns(COL_DETAILS_OPERATION).DataBodyRange(rowIndex).Value) Then
                    amount1 = table.ListColumns(COL_DETAILS_ACQUIRED).DataBodyRange(rowIndex).Value
                    amount2 = table.ListColumns(COL_DETAILS_AMOUNT).DataBodyRange(rowIndex).Value
                    
                    amount1 = evaluator.EvaluateHistoricalCostEx(stamp, m_ticker, ticker, amount1, amount2)
                    
                    amount2 = table.ListColumns(COL_DETAILS_SPENT).DataBodyRange(rowIndex).Value
                    amount2 = evaluator.EvaluateHistoricalCost(stamp, m_ticker, amount2)
                    
                    ticker = m_ticker
                Else
                    amount1 = table.ListColumns(COL_DETAILS_SPENT).DataBodyRange(rowIndex).Value
                    amount2 = table.ListColumns(COL_DETAILS_AMOUNT).DataBodyRange(rowIndex).Value
                
                    amount2 = evaluator.EvaluateHistoricalCostEx(stamp, m_ticker, ticker, amount1, amount2)
                    
                    amount1 = table.ListColumns(COL_DETAILS_ACQUIRED).DataBodyRange(rowIndex).Value
                    amount1 = evaluator.EvaluateHistoricalCost(stamp, m_ticker, amount1)
                End If

                table.ListColumns(COL_DETAILS_DEBIT).DataBodyRange(rowIndex).Value = amount1
                table.ListColumns(COL_DETAILS_CREDIT).DataBodyRange(rowIndex).Value = amount2
            
                amount1 = table.ListColumns(COL_DETAILS_CHARGE).DataBodyRange(rowIndex).Value
                amount2 = table.ListColumns(COL_DETAILS_BNB_FEE).DataBodyRange(rowIndex).Value
                
                amount1 = evaluator.EvaluateHistoricalCost(stamp, ticker, amount1)
                amount2 = evaluator.EvaluateHistoricalCost(stamp, STR_BNB, amount2)
                
                table.ListColumns(COL_DETAILS_COMMISSION).DataBodyRange(rowIndex).Value = CDec(amount1) + CDec(amount2)
            End If
        End If
    Next rowIndex
End Sub

Public Sub PopulateOperations(ByVal ops As Collection, Optional ByVal override As Boolean = False)
    Dim table As ListObject
    Set table = m_ws.ListObjects(xTableName(TABLE_DETAILS))
    
    Dim op As BinanceOp
    For Each op In ops
        Dim forDelete As Boolean
        forDelete = op.acquired = 0 And op.spent = 0
        
        Dim rowIndex As Long
        rowIndex = xFindOpRowIndex(table, op.id)
        
        If rowIndex = 0 Then
            If Not forDelete Then rowIndex = table.ListRows.Add().index
        ElseIf forDelete Then
            table.ListRows(rowIndex).Delete
            rowIndex = 0
        ElseIf Not override Then
            rowIndex = 0
        End If
        
        If rowIndex > 0 Then
            table.ListColumns(COL_DETAILS_STAMP).DataBodyRange(rowIndex).Value = op.stamp
            table.ListColumns(COL_DETAILS_WALLET).DataBodyRange(rowIndex).Value = op.wallet
            table.ListColumns(COL_DETAILS_OPERATION).DataBodyRange(rowIndex).Value = op.operation
            table.ListColumns(COL_DETAILS_ACQUIRED).DataBodyRange(rowIndex).Value = op.acquired
            table.ListColumns(COL_DETAILS_SPENT).DataBodyRange(rowIndex).Value = op.spent
            If Len(op.ticker) > 0 Then
                table.ListColumns(COL_DETAILS_TICKER).DataBodyRange(rowIndex).Value = op.ticker
                table.ListColumns(COL_DETAILS_PRICE).DataBodyRange(rowIndex).Value = op.price
                table.ListColumns(COL_DETAILS_AMOUNT).DataBodyRange(rowIndex).Value = op.amount
                table.ListColumns(COL_DETAILS_CHARGE).DataBodyRange(rowIndex).Value = op.charge
                table.ListColumns(COL_DETAILS_BNB_FEE).DataBodyRange(rowIndex).Value = op.bnb_fee
            End If
            table.ListColumns(COL_DETAILS_NOTE).DataBodyRange(rowIndex).Value = op.note
            table.ListColumns(COL_DETAILS_ID).DataBodyRange(rowIndex).Value = op.id
        End If
    Next op

    xSort
End Sub

Private Function xTableName(ByVal name As String) As String
    xTableName = m_ticker & "_" & name
End Function

Private Function xFindOpRowIndex(ByVal table As ListObject, ByVal id As String) As Long
    xFindOpRowIndex = 0
    
    If table.ListRows.count > 0 Then
        Dim cell As Range
        Set cell = table.ListColumns(COL_DETAILS_ID).DataBodyRange.Find(what:=id, LookIn:=xlValues, LookAt:=xlWhole)
        
        If Not cell Is Nothing Then
            xFindOpRowIndex = cell.Row - table.HeaderRowRange.Row
        End If
    End If
End Function

Private Function xCollectTradeMarkets() As Dictionary
    Dim table As ListObject
    Set table = m_ws.ListObjects(xTableName(TABLE_MARKETS))

    Dim markets As New Dictionary

    Dim i As Long
    For i = 2 To table.ListColumns.count - 1
        If table.ListColumns(i).DataBodyRange(1).Value = True Then
            Dim quote, symbol As String
            
            quote = table.ListColumns(i).name
            symbol = m_ticker & quote
            
            markets.Add quote, symbol
        End If
    Next i
    
    Set xCollectTradeMarkets = markets
End Function

Private Sub xSort()
    Dim table As ListObject
    Set table = m_ws.ListObjects(xTableName(TABLE_DETAILS))
    
    Dim col As ListColumn
    Set col = table.ListColumns(COL_DETAILS_STAMP)
    
    If Not col.DataBodyRange Is Nothing Then
        With table.Sort
            .SortFields.Clear
            .SortFields.Add key:=col.DataBodyRange, SortOn:=xlSortOnValues, order:=xlAscending, DataOption:=xlSortNormal
            .header = xlYes
            .Apply
        End With
    End If
End Sub

Private Function xCollectPeriods(ByVal quote As String) As Dictionary
    Dim periods As New Dictionary
    
    Set xCollectPeriods = periods
    
    Dim interval As String
    interval = CalcKlineInterval(m_ws.ListObjects(xTableName(TABLE_TRIVIA)).ListColumns(COL_TRIVIA_VALUE).DataBodyRange(ROW_TRIVIA_KLINES_INTERVAL).Value)
    
    Dim table As ListObject
    Set table = m_ws.ListObjects(xTableName(TABLE_DETAILS))

    Dim rowIndex As Long
    For rowIndex = 1 To table.ListRows.count
        If _
            IsEmpty(table.ListColumns(COL_DETAILS_DEBIT).DataBodyRange(rowIndex).Value) Or _
            IsEmpty(table.ListColumns(COL_DETAILS_CREDIT).DataBodyRange(rowIndex).Value) _
        Then
            Dim stamp As Date
            stamp = table.ListColumns(COL_DETAILS_STAMP).DataBodyRange(rowIndex).Value
            
            xAddTickerToPeriods periods, m_ticker, stamp, interval, quote
            
            Dim ticker As String
            ticker = table.ListColumns(COL_DETAILS_TICKER).DataBodyRange(rowIndex).Value
            
            If ticker <> "" Then
                xAddTickerToPeriods periods, ticker, stamp, interval, quote
            End If
            
            If Not IsEmpty(table.ListColumns(COL_DETAILS_BNB_FEE).DataBodyRange(rowIndex).Value) Then
                xAddTickerToPeriods periods, STR_BNB, stamp, interval, quote
            End If
        End If
    Next rowIndex
End Function

Private Sub xAddTickerToPeriods(ByVal periods As Dictionary, ByVal ticker As String, ByVal stamp As Date, ByVal interval As String, ByVal quote As String)
    Dim period As PricePeriod
    
    If periods.Exists(ticker) Then
        Set period = periods(ticker)
        If period.startDate > stamp Then period.startDate = stamp
        If period.endDate < stamp Then period.endDate = stamp
    Else
        Set period = New PricePeriod
        period.ticker = ticker
        period.interval = interval
        period.startDate = stamp
        period.endDate = stamp
        period.path = TableCurrencies.GetEvalPath(ticker, quote)
        periods.Add ticker, period
    End If
End Sub

Private Sub xReset(ByVal clearDetails)
    Dim coin As BinanceCoin
    Set coin = TableCurrencies.GetTickerInfo(m_ticker)
    
    Dim numberFormat As String
    If coin Is Nothing Then
        numberFormat = NUMBER_FORMAT_MAIN_CRYPTO
    Else
        numberFormat = CalculateNumberFormat(coin.isFiat, coin.precision)
    End If
    
    If clearDetails Then
        TableCurrencies.SetTradesLastUpdate m_ticker, 0
    End If
    
    Dim table As ListObject
    
    Set table = m_ws.ListObjects(xTableName(TABLE_TRIVIA))
    
    If clearDetails Then
        table.ListColumns(COL_TRIVIA_VALUE).DataBodyRange(ROW_TRIVIA_TICKER).Value = m_ticker
    End If
    If Not coin Is Nothing Then
        table.ListColumns(COL_TRIVIA_VALUE).DataBodyRange(ROW_TRIVIA_NAME).Value = coin.name
        table.ListColumns(COL_TRIVIA_VALUE).DataBodyRange(ROW_TRIVIA_TYPE).Value = IIf(coin.isFiat, STR_FIAT, STR_CRYPTO)
    End If
    
    table.ListColumns(COL_TRIVIA_VALUE).DataBodyRange(ROW_TRIVIA_TOTAL).numberFormat = numberFormat
    table.ListColumns(COL_TRIVIA_VALUE).DataBodyRange(ROW_TRIVIA_BALANCE).numberFormat = numberFormat
    table.ListColumns(COL_TRIVIA_VALUE).DataBodyRange(ROW_TRIVIA_DYNAMIC).numberFormat = numberFormat
    table.ListColumns(COL_TRIVIA_VALUE).DataBodyRange(ROW_TRIVIA_DIVIATION).numberFormat = numberFormat
    
    Set table = m_ws.ListObjects(xTableName(TABLE_SUMMARY))

    table.ListColumns(COL_SUMMARY_ACQUIRED).DataBodyRange.numberFormat = numberFormat
    table.ListColumns(COL_SUMMARY_SPENT).DataBodyRange.numberFormat = numberFormat
    table.ListColumns(COL_SUMMARY_CHARGE).DataBodyRange.numberFormat = numberFormat
    table.ListColumns(COL_SUMMARY_BALANCE).DataBodyRange.numberFormat = numberFormat

    table.ListColumns(COL_SUMMARY_ACQUIRED).Total.numberFormat = numberFormat
    table.ListColumns(COL_SUMMARY_SPENT).Total.numberFormat = numberFormat
    table.ListColumns(COL_SUMMARY_CHARGE).Total.numberFormat = numberFormat
    table.ListColumns(COL_SUMMARY_BALANCE).Total.numberFormat = numberFormat
    
    Set table = m_ws.ListObjects(xTableName(TABLE_DETAILS))
    
    If Not table.DataBodyRange Is Nothing Then
        table.ListColumns(COL_DETAILS_ACQUIRED).DataBodyRange.numberFormat = numberFormat
        table.ListColumns(COL_DETAILS_SPENT).DataBodyRange.numberFormat = numberFormat

        If clearDetails Then table.DataBodyRange.Delete
    End If

    table.ListColumns(COL_DETAILS_ACQUIRED).Total.numberFormat = numberFormat
    table.ListColumns(COL_DETAILS_SPENT).Total.numberFormat = numberFormat
    
    Dim col As ListColumn
    
    Set table = m_ws.ListObjects(xTableName(TABLE_MARKETS))
    
    Application.EnableEvents = False
    Do While table.ListColumns.count > 2
        Set col = table.ListColumns(2)
        If Not coin Is Nothing Then
            If coin.markets.Exists(col.name) Then coin.markets.item(col.name) = col.DataBodyRange(1).Value
        End If
        col.Delete
    Loop
    If Not coin Is Nothing Then
        Dim market As Variant
        For Each market In coin.markets.Keys
            Set col = table.ListColumns.Add(table.ListColumns.count)
            col.name = market
            col.DataBodyRange(1).Value = coin.markets(market)
        Next market
    End If
    Application.EnableEvents = True
    
    Dim observedFormula As String
    If table.ListColumns.count <= 2 Then
        observedFormula = "=0"
    Else
        observedFormula = _
            "=COUNTIF(" & xTableName(TABLE_MARKETS) & "[@[" & _
            table.ListColumns(2).name & _
            "]:[" & _
            table.ListColumns(table.ListColumns.count - 1).name & _
            "]],TRUE)"
    End If
    
    table.ListColumns(COL_MARKETS_TICKER).DataBodyRange(1).Value = m_ticker
    table.ListColumns(COL_MARKETS_OBSERVED).DataBodyRange(1).formula = observedFormula
End Sub
