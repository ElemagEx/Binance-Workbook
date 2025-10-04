Attribute VB_Name = "BinanceLedgers"
Option Explicit
Option Private Module

Public Function HasExistingLedgers()
    HasExistingLedgers = True
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If xIsLedgerWorksheet(ws) Then
            Exit Function
        End If
    Next ws
    HasExistingLedgers = False
End Function

Public Sub RemoveAll()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If xIsLedgerWorksheet(ws) Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
    Next ws
End Sub

Public Sub RefreshCoinsInfo()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If xIsLedgerWorksheet(ws) Then
            xOpenLedger(ws).RefreshCoinInfo
        End If
    Next ws
End Sub

Public Sub OpenCurrentTickerLedger()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim ticker As String
    If ticker = "" Then ticker = TableWallet.GetSelectedTicker()
    If ticker = "" Then ticker = TableCurrencies.GetSelectedTicker()
    
    If ticker = "" Then
        MsgBox "Please select cell(s) in single row in Assets or Currencies Sheets"
        Exit Sub
    End If
    
    Dim ledger As BinanceLedger
    Set ledger = xFindLedger(ticker, True)
    ledger.Activate
End Sub

Public Sub UpdateCurrentLedgerTrades()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim ledger As BinanceLedger
    Set ledger = xOpenLedger(ws)
    
    If ledger Is Nothing Then
        MsgBox "Open Ledger worksheet first"
        Exit Sub
    End If
    
    ledger.UpdateTrades
End Sub

Public Sub PopulateOperations(ByVal ops As BinanceOps, Optional ByVal override As Boolean)
    Dim ticker As Variant
    For Each ticker In ops.tickers
        Dim onlyDelOps As Boolean
        onlyDelOps = ops.HasOnlyDelOps(ticker)
    
        If Not onlyDelOps Then TableWallet.CheckTicker ticker
        
        Dim ledger As BinanceLedger
        Set ledger = xFindLedger(ticker, Not onlyDelOps)
        
        If Not ledger Is Nothing Then
            ledger.Activate
            ledger.PopulateOperations ops.item(ticker), override
        End If
    Next ticker
End Sub

Public Function GetOrderString() As String
    Dim order As String
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If xIsLedgerWorksheet(ws) Then
            order = order & "," & xOpenLedger(ws).ticker
        End If
    Next ws
    GetOrderString = IIf(Len(order) = 0, "", Mid(order, 2))
End Function

Public Function xLedgerName(ByVal ticker) As String
    xLedgerName = ticker & " " & LEDGER_SHEET_SUFFIX
End Function

Private Function xOpenLedger(ByVal ws As Worksheet) As BinanceLedger
    Set xOpenLedger = Nothing
    
    If Not xIsLedgerWorksheet(ws) Then Exit Function
    
    Dim parts() As String
    parts = Split(ws.name, " ")
    
    Dim ticker As String
    ticker = parts(LBound(parts))
    
    Dim ledger As New BinanceLedger
    ledger.Init ticker, ws
    
    Set xOpenLedger = ledger
End Function

Private Function xFindLedger(ByVal ticker As String, ByVal createIfDoesNotExists As Boolean) As BinanceLedger
    Set xFindLedger = Nothing
    
    Dim name As String
    name = xLedgerName(ticker)
    
    Dim ledger As BinanceLedger
    
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If name = ws.name Then
            Set ledger = New BinanceLedger
            ledger.Init ticker, ws
            Set xFindLedger = ledger
            Exit Function
        End If
    Next ws

    If createIfDoesNotExists Then
        Set ledger = New BinanceLedger
        ledger.Create ticker, name
        Set xFindLedger = ledger
    End If
End Function

Private Function xIsLedgerWorksheet(Optional ByVal ws As Worksheet = Nothing)
    xIsLedgerWorksheet = False
    
    If ws Is Nothing Then
        Set ws = ActiveSheet
    End If
    
    Dim parts() As String
    parts = Split(ws.name, " ")
    
    Dim count As Long
    count = UBound(parts) - LBound(parts) + 1

    If count <> 2 Or parts(UBound(parts)) <> LEDGER_SHEET_SUFFIX Then
        Exit Function
    End If
    
    Dim ticker As String
    ticker = parts(LBound(parts))

    Dim ledger As BinanceLedger
    Set ledger = New BinanceLedger
    ledger.Init ticker, ws
    
    xIsLedgerWorksheet = ledger.IsValid()
End Function
