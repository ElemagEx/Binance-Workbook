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

Public Sub PopulateOperations(ByVal ops As BinanceOps, Optional ByVal override As Boolean)
    Dim ticker As Variant
    For Each ticker In ops.Tickers
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

Public Function xLedgerName(ByVal ticker) As String
    xLedgerName = ticker & " " & LEDGER_SHEET_SUFFIX
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
