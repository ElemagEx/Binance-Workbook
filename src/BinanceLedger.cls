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
Private Const ROW_TRIVIA_TOTAL As Long = 4
Private Const ROW_TRIVIA_BALANCE As Long = 7
Private Const ROW_TRIVIA_DYNAMIC As Long = 8
Private Const ROW_TRIVIA_DIVIATION As Long = 9
Private Const ROW_TRIVIA_ACTIVITY_START_DATE As Long = 10

Private Const COL_SUMMARY_ACQUIRED As String = "Acquired"
Private Const COL_SUMMARY_SPENT As String = "Spent"
Private Const COL_SUMMARY_CHARGE As String = "Charge"
Private Const COL_SUMMARY_BALANCE As String = "Balance"

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
Private Const COL_DETAILS_NOTE As String = "Note"
Private Const COL_DETAILS_ID As String = "ID"

Private m_ws As Worksheet
Private m_ticker As String

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
End Sub

Public Sub PopulateOperations(ByVal ops As Collection, Optional ByVal override As Boolean = False)
    Dim table As ListObject
    Set table = m_ws.ListObjects(xTableName(TABLE_DETAILS))
    
    Dim op As BinanceOp
    For Each op In ops
        Dim rowIndex As Long
        rowIndex = xFindOpRowIndex(table, op.id)
        
        If rowIndex = 0 Then
            rowIndex = table.ListRows.Add().index
        ElseIf Not override Then
            rowIndex = 0
        End If
        
        If rowIndex > 0 Then
            table.ListColumns(COL_DETAILS_STAMP).DataBodyRange(rowIndex).Value = op.stamp
            table.ListColumns(COL_DETAILS_WALLET).DataBodyRange(rowIndex).Value = op.wallet
            table.ListColumns(COL_DETAILS_OPERATION).DataBodyRange(rowIndex).Value = op.operation
            table.ListColumns(COL_DETAILS_ACQUIRED).DataBodyRange(rowIndex).Value = op.acquired
            table.ListColumns(COL_DETAILS_SPENT).DataBodyRange(rowIndex).Value = op.spent
            table.ListColumns(COL_DETAILS_TICKER).DataBodyRange(rowIndex).Value = op.ticker
            table.ListColumns(COL_DETAILS_PRICE).DataBodyRange(rowIndex).Value = op.price
            table.ListColumns(COL_DETAILS_AMOUNT).DataBodyRange(rowIndex).Value = op.amount
            table.ListColumns(COL_DETAILS_CHARGE).DataBodyRange(rowIndex).Value = op.charge
            table.ListColumns(COL_DETAILS_BNB_FEE).DataBodyRange(rowIndex).Value = op.bnb_fee
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

Private Sub xSort()
    Dim table As ListObject
    Set table = m_ws.ListObjects(xTableName(TABLE_DETAILS))
    
    Dim col As ListColumn
    Set col = table.ListColumns(COL_DETAILS_STAMP)
    
    If Not col.DataBodyRange Is Nothing Then
        With table.Sort
            .SortFields.Clear
            .SortFields.Add key:=col.DataBodyRange, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .header = xlYes
            .Apply
        End With
    End If
End Sub

Private Sub xReset(ByVal clearDetails)
    Dim coin As BinanceCoin
    Set coin = TableCurrencies.GetTickerInfo(m_ticker)
    
    Dim numFormat As String
    numFormat = IIf(coin.isFiat, "#,##0.00", "#,##0.0" + String(coin.precision - 1, "?"))
    
    Dim table As ListObject
    
    Set table = m_ws.ListObjects(xTableName(TABLE_TRIVIA))
    
    table.ListColumns(COL_TRIVIA_VALUE).DataBodyRange(ROW_TRIVIA_TICKER).Value = coin.ticker
    table.ListColumns(COL_TRIVIA_VALUE).DataBodyRange(ROW_TRIVIA_NAME).Value = coin.name
    table.ListColumns(COL_TRIVIA_VALUE).DataBodyRange(ROW_TRIVIA_TYPE).Value = IIf(coin.isFiat, STR_FIAT, STR_CRYPTO)
    
    table.ListColumns(COL_TRIVIA_VALUE).DataBodyRange(ROW_TRIVIA_TOTAL).NumberFormat = numFormat
    table.ListColumns(COL_TRIVIA_VALUE).DataBodyRange(ROW_TRIVIA_BALANCE).NumberFormat = numFormat
    table.ListColumns(COL_TRIVIA_VALUE).DataBodyRange(ROW_TRIVIA_DYNAMIC).NumberFormat = numFormat
    table.ListColumns(COL_TRIVIA_VALUE).DataBodyRange(ROW_TRIVIA_DIVIATION).NumberFormat = numFormat

    Set table = m_ws.ListObjects(xTableName(TABLE_SUMMARY))

    table.ListColumns(COL_SUMMARY_ACQUIRED).DataBodyRange.NumberFormat = numFormat
    table.ListColumns(COL_SUMMARY_SPENT).DataBodyRange.NumberFormat = numFormat
    table.ListColumns(COL_SUMMARY_CHARGE).DataBodyRange.NumberFormat = numFormat
    table.ListColumns(COL_SUMMARY_BALANCE).DataBodyRange.NumberFormat = numFormat

    Set table = m_ws.ListObjects(xTableName(TABLE_DETAILS))
    
    If clearDetails And Not table.DataBodyRange Is Nothing Then
        table.DataBodyRange.Delete
    End If
    
    table.ListColumns(COL_DETAILS_ACQUIRED).Range.NumberFormat = numFormat
    table.ListColumns(COL_DETAILS_SPENT).Range.NumberFormat = numFormat

    Set table = m_ws.ListObjects(xTableName(TABLE_MARKETS))
    
    Do While table.ListColumns.count > 2
        table.ListColumns(2).Delete
    Loop

    Dim market As Variant
    For Each market In coin.markets.Keys
        Dim col As ListColumn
        Set col = table.ListColumns.Add(table.ListColumns.count)
        col.name = market
        col.DataBodyRange(1).Value = coin.markets(market)
    Next market
    
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
