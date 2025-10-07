Attribute VB_Name = "TableWallet"
Option Explicit
Option Private Module

Private Const COL_TICKER As String = "Ticker"
Private Const COL_PRICE As String = "Price"
Private Const COL_TOTAL_AMOUNT As String = "Total Amount"
Private Const COL_SPOT_AMOUNT As String = "Spot Amount"
Private Const COL_FUNDING_AMOUNT As String = "Funding Amount"
Private Const COL_EARN_AMOUNT As String = "Earn Amount"
Private Const COL_DYNAMIC_AMOUNT As String = "Dynamic Amount"
Private Const COL_STATIC_AMOUNT As String = "Static Amount"
Private Const COL_TOTAL_COST As String = "Total Cost"
Private Const COL_SPOT_COST As String = "Spot Cost"
Private Const COL_FUNDING_COST As String = "Funding Cost"
Private Const COL_EARN_COST As String = "Earn Cost"
Private Const COL_DYNAMIC_COST As String = "Dynamic Cost"
Private Const COL_STATIC_COST As String = "Static Cost"

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

Public Function GetSelectedTicker() As String
    GetSelectedTicker = ""
    
    If ActiveSheet.name = SHEET_ASSETS Then
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

Public Sub CheckTicker(ByVal ticker As String)
    xFindTickerRowIndex ticker, True
End Sub

Public Function IsDataCleanedUp() As Boolean
    IsDataCleanedUp = (xGetTable().ListRows.count = 0)
End Function

Public Sub CleanUpData()
    Dim table As ListObject
    Set table = xGetTable()
    
    If Not table.DataBodyRange Is Nothing Then
        table.DataBodyRange.Delete
    End If

    TableLastUpdate.wallet = 0
End Sub

Public Sub ClearData()
    Dim table As ListObject
    Set table = xGetTable()
    
    If table.ListRows.count > 0 Then
        table.ListColumns(COL_SPOT_AMOUNT).DataBodyRange.ClearContents
        table.ListColumns(COL_FUNDING_AMOUNT).DataBodyRange.ClearContents
        table.ListColumns(COL_EARN_AMOUNT).DataBodyRange.ClearContents
        table.ListColumns(COL_DYNAMIC_AMOUNT).DataBodyRange.ClearContents
    End If

    TableLastUpdate.wallet = 0
End Sub

Public Sub ClearCosts()
    Dim table As ListObject
    Set table = xGetTable()
    
    If table.ListRows.count = 0 Then Exit Sub
    
    table.ListColumns(COL_PRICE).DataBodyRange.ClearContents
    table.ListColumns(COL_SPOT_COST).DataBodyRange.ClearContents
    table.ListColumns(COL_FUNDING_COST).DataBodyRange.ClearContents
    table.ListColumns(COL_EARN_COST).DataBodyRange.ClearContents
    table.ListColumns(COL_DYNAMIC_COST).DataBodyRange.ClearContents

    table.ListColumns(COL_TOTAL_COST).Range.numberFormat = "General"
    table.ListColumns(COL_SPOT_COST).Range.numberFormat = "General"
    table.ListColumns(COL_FUNDING_COST).Range.numberFormat = "General"
    table.ListColumns(COL_EARN_COST).Range.numberFormat = "General"
    table.ListColumns(COL_DYNAMIC_COST).Range.numberFormat = "General"
    table.ListColumns(COL_STATIC_COST).Range.numberFormat = "General"
End Sub

Public Sub UpdateData()
    ClearData
    
    Dim assets As New BinanceAssets
    
    assets.CollectSpotWalletAssets COL_SPOT_AMOUNT
    assets.CollectFundingWalletAssets COL_FUNDING_AMOUNT
    assets.CollectSimpleEarnLockedAssets COL_EARN_AMOUNT
    assets.CollectSimpleEarnFlexibleAssets COL_DYNAMIC_AMOUNT
    
    Dim table As ListObject
    Set table = xGetTable()
    
    Dim ticker As Variant
    For Each ticker In assets.wallet.Keys
        
        Dim rowIndex As Long
        rowIndex = xFindTickerRowIndex(ticker, True)
        
        Dim name As Variant
        For Each name In assets.wallet(ticker).Keys
            table.ListColumns(name).DataBodyRange(rowIndex) = assets.wallet(ticker)(name)
        Next name
        
    Next ticker
    
    TableLastUpdate.wallet = Now()
End Sub

Public Sub Evaluate()
    Dim quote As String
    quote = TableUserInfo.quote
    
    Dim paths As Dictionary
    Set paths = xCollectPaths(quote)

    If paths.count = 0 Then Exit Sub
    
    Dim table As ListObject
    Set table = xGetTable()
    
    Dim evaluator As New PriceEvaluator
    evaluator.GetQuotePrices quote, paths
    
    Dim info As PriceQuote
    Set info = evaluator.GetQuoteInfo()
    
    Dim numberFormat As String
    numberFormat = CalculateNumberFormat(info.isFiat)
    table.ListColumns(COL_TOTAL_COST).DataBodyRange.numberFormat = numberFormat
    table.ListColumns(COL_SPOT_COST).DataBodyRange.numberFormat = numberFormat
    table.ListColumns(COL_FUNDING_COST).DataBodyRange.numberFormat = numberFormat
    table.ListColumns(COL_EARN_COST).DataBodyRange.numberFormat = numberFormat
    table.ListColumns(COL_DYNAMIC_COST).DataBodyRange.numberFormat = numberFormat
    table.ListColumns(COL_STATIC_COST).DataBodyRange.numberFormat = numberFormat
    
    table.ListColumns(COL_TOTAL_COST).Total.numberFormat = numberFormat
    table.ListColumns(COL_SPOT_COST).Total.numberFormat = numberFormat
    table.ListColumns(COL_FUNDING_COST).Total.numberFormat = numberFormat
    table.ListColumns(COL_EARN_COST).Total.numberFormat = numberFormat
    table.ListColumns(COL_DYNAMIC_COST).Total.numberFormat = numberFormat
    table.ListColumns(COL_STATIC_COST).Total.numberFormat = numberFormat
    
    Dim rowIndex As Long
    For rowIndex = 1 To table.ListRows.count
        Dim ticker As String
        ticker = table.ListColumns(COL_TICKER).DataBodyRange(rowIndex).Value
        
        table.ListColumns(COL_PRICE).DataBodyRange(rowIndex).Value = evaluator.EvaluateCost(ticker, CDec(1))
        
        Dim amount As Variant
        amount = table.ListColumns(COL_SPOT_AMOUNT).DataBodyRange(rowIndex).Value
        If Not IsEmpty(amount) Then table.ListColumns(COL_SPOT_COST).DataBodyRange(rowIndex).Value = evaluator.EvaluateCost(ticker, CDec(amount))
        
        amount = table.ListColumns(COL_FUNDING_AMOUNT).DataBodyRange(rowIndex).Value
        If Not IsEmpty(amount) Then table.ListColumns(COL_FUNDING_COST).DataBodyRange(rowIndex).Value = evaluator.EvaluateCost(ticker, CDec(amount))
        
        amount = table.ListColumns(COL_EARN_AMOUNT).DataBodyRange(rowIndex).Value
        If Not IsEmpty(amount) Then table.ListColumns(COL_EARN_COST).DataBodyRange(rowIndex).Value = evaluator.EvaluateCost(ticker, CDec(amount))
        
        amount = table.ListColumns(COL_DYNAMIC_AMOUNT).DataBodyRange(rowIndex).Value
        If Not IsEmpty(amount) Then table.ListColumns(COL_DYNAMIC_COST).DataBodyRange(rowIndex).Value = evaluator.EvaluateCost(ticker, CDec(amount))
    Next rowIndex
End Sub

Public Sub Sort(ByVal tickers As String)
    Dim table As ListObject
    Set table = xGetTable()
    
    If table.ListRows.count > 0 Then
        Dim cells As Range
        Set cells = table.ListColumns(COL_TICKER).DataBodyRange
        
        With table.Sort
            .SortFields.Clear
            .SortFields.Add cells, xlSortOnValues, xlAscending, tickers, xlSortNormal
            .header = xlYes
            .Apply
        End With
    End If
End Sub

Public Function xGetTable() As ListObject
    Set xGetTable = ThisWorkbook.Sheets(SHEET_ASSETS).ListObjects(TABLE_WALLET)
End Function

Private Function xFindTickerRowIndex(ByVal ticker As String, ByVal addIfNotFound As Boolean) As Long
    Dim table As ListObject
    Set table = xGetTable()
    
    Dim cell As Range
    If table.ListRows.count > 0 Then
        Set cell = table.ListColumns(COL_TICKER).DataBodyRange.Find(what:=ticker, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    End If
    
    Dim index As Long
    
    If Not cell Is Nothing Then
        index = cell.Row - table.HeaderRowRange.Row
    ElseIf Not addIfNotFound Then
        index = 0
    Else
        index = table.ListRows.Add().index
        table.ListColumns(COL_TICKER).DataBodyRange(index).Value = ticker
    End If
    
    xFindTickerRowIndex = index
End Function

Public Function xCollectPaths(ByVal quote As String) As Dictionary
    Dim paths As New Dictionary
    
    Set xCollectPaths = paths
    
    If quote = "" Then Exit Function
    
    Dim table As ListObject
    Set table = xGetTable()

    If table.ListRows.count = 0 Then Exit Function
    
    Dim col As ListColumn
    Set col = table.ListColumns(COL_TICKER)
    
    Dim cell As Range
    For Each cell In col.DataBodyRange
        Dim ticker As String
        ticker = cell.Value
        paths.Add ticker, Empty
    Next cell
    
    TableCurrencies.CollectEvalPaths paths, quote
End Function
