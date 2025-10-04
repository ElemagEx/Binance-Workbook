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

Public Function xGetTable() As ListObject
    Set xGetTable = ThisWorkbook.Sheets(SHEET_MISC).ListObjects(TABLE_EVALUATION)
End Function

