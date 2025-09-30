Attribute VB_Name = "TableWallet"
Option Explicit
Option Private Module

Private Const SHEET_NAME As String = SHEET_ASSETS
Private Const TABLE_NAME As String = "Wallet"

Private Const COL_TICKER As String = "Ticker"
Private Const COL_SPOT_AMOUNT As String = "Spot Amount"
Private Const COL_FUNDING_AMOUNT As String = "Funding Amount"
Private Const COL_EARN_AMOUNT As String = "Earn Amount"
Private Const COL_DYNAMIC_AMOUNT As String = "Dynamic Amount"

Public Property Get Tickers() As Collection
    Set Tickers = New Collection

    Dim col As ListColumn
    Set col = ThisWorkbook.Sheets(SHEET_NAME).ListObjects(TABLE_NAME).ListColumns(COL_TICKER)
    
    If Not col.DataBodyRange Is Nothing Then
        Dim i As Long
        For i = 1 To col.DataBodyRange.count
            Tickers.Add col.DataBodyRange(i).Value
        Next i
    End If
End Property

Public Sub CheckTicker(ByVal ticker As String)
    xFindTickerRowIndex ticker, True
End Sub

Public Function IsDataCleanedUp() As Boolean
    Dim table As ListObject
    Set table = ThisWorkbook.Sheets(SHEET_NAME).ListObjects(TABLE_NAME)

    IsDataCleanedUp = (table.ListRows.count = 0)
End Function

Public Sub CleanUpData()
    Dim table As ListObject
    Set table = ThisWorkbook.Sheets(SHEET_NAME).ListObjects(TABLE_NAME)
    
    If Not table.DataBodyRange Is Nothing Then
        table.DataBodyRange.Delete
    End If

    TableLastUpdate.wallet = 0
End Sub

Public Sub ClearData()
    Dim table As ListObject
    Set table = ThisWorkbook.Sheets(SHEET_NAME).ListObjects(TABLE_NAME)
    
    If table.ListRows.count > 0 Then
        table.ListColumns(COL_SPOT_AMOUNT).DataBodyRange.ClearContents
        table.ListColumns(COL_FUNDING_AMOUNT).DataBodyRange.ClearContents
        table.ListColumns(COL_EARN_AMOUNT).DataBodyRange.ClearContents
        table.ListColumns(COL_DYNAMIC_AMOUNT).DataBodyRange.ClearContents
    End If

    TableLastUpdate.wallet = 0
End Sub

Public Sub UpdateData()
    ClearData
    
    Dim assets As New BinanceAssets
    
    assets.CollectSpotWalletAssets COL_SPOT_AMOUNT
    assets.CollectFundingWalletAssets COL_FUNDING_AMOUNT
    assets.CollectSimpleEarnLockedAssets COL_EARN_AMOUNT
    assets.CollectSimpleEarnFlexibleAssets COL_DYNAMIC_AMOUNT
    
    Dim table As ListObject
    Set table = ThisWorkbook.Sheets(SHEET_NAME).ListObjects(TABLE_NAME)
    
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

Private Function xFindTickerRowIndex(ByVal ticker As String, ByVal addIfNotFound As Boolean) As Long
    Dim table As ListObject
    Set table = ThisWorkbook.Sheets(SHEET_NAME).ListObjects(TABLE_NAME)
    
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

