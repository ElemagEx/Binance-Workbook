Attribute VB_Name = "TableCurrencies"
Option Explicit

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
    
    Dim coins As New BinanceCoins
    coins.collectBasicInfo False
    coins.collectExchangeInfo

    ResetExchangeTimezone coins.ExchangeTimezone

End Sub
