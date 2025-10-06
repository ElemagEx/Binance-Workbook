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

Public Function FindQuoteInfo(ByVal quote As String) As PriceQuote
    Set FindQuoteInfo = Nothing
    
    Dim table As ListObject
    Set table = xGetTable()
    
    Dim i As Long
    For i = 2 To table.ListColumns.count
        Dim col As ListColumn
        Set col = table.ListColumns(i)
        
        Dim ticker As String
        ticker = col.name
    
        If ticker = quote Then
            Dim info As New PriceQuote
            info.ticker = ticker
            info.method = col.DataBodyRange(ROW_METHOD).Value
            info.refers = col.DataBodyRange(ROW_REFERS).Value
            info.format = col.DataBodyRange(ROW_FORMAT).NumberFormat
            Set FindQuoteInfo = info
            Exit Function
        End If
    Next i
End Function

Public Function GetInfos() As Dictionary
    Dim infos As New Dictionary
        
    Dim table As ListObject
    Set table = xGetTable()
    
    Dim i As Long
    For i = 2 To table.ListColumns.count
        Dim col As ListColumn
        Set col = table.ListColumns(i)
        
        Dim ticker As String
        ticker = col.name
        
        Dim info As PriceQuote
        Set info = New PriceQuote
        info.ticker = ticker
        info.method = col.DataBodyRange(ROW_METHOD).Value
        info.refers = col.DataBodyRange(ROW_REFERS).Value
        info.format = col.DataBodyRange(ROW_FORMAT).NumberFormat
        
        infos.Add ticker, info
    Next i
        
    Set GetInfos = infos
End Function

Public Sub EvaluatePrices(ByVal tickers As Dictionary)
    Dim prices As Dictionary
    Set prices = GetPrices()
    
    Dim path As String
    Dim price As Variant
    Dim ticker As Variant
    For Each ticker In tickers.Keys
        path = tickers(ticker)
        price = xEvalPath(path, prices)
        tickers.item(ticker) = price
    Next ticker
End Sub

Private Function xGetTable() As ListObject
    Set xGetTable = ThisWorkbook.Sheets(SHEET_MISC).ListObjects(TABLE_EVALUATION)
End Function

Private Function xEvalPath(ByVal path As String, ByVal prices As Dictionary) As Variant
    xEvalPath = Empty
    
    If path = STR_NA Then Exit Function
    
    Dim parts() As String
    parts = Split(path, ",")
    
    Dim op, symbol As String
    Dim price As Variant
    
    price = CDec(1)
    
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        If parts(i) <> "=" Then
            op = Left(parts(i), 1)
            symbol = Mid(parts(i), 2)
            Select Case op
                Case "*"
                    If Not prices.Exists(symbol) Then Exit Function
                    price = price * prices(symbol)
                Case "/"
                    If Not prices.Exists(symbol) Then Exit Function
                    price = price / prices(symbol)
                Case "$"
                    Exit Function
                Case Else
                    Assert_Fail
            End Select
        End If
    Next i
    
    xEvalPath = price
End Function
