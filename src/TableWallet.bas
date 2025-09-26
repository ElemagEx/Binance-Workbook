Attribute VB_Name = "TableWallet"
Option Explicit

Private Const SHEET_NAME As String = SHEET_ASSETS
Private Const TABLE_NAME As String = "Wallet"

Private Const COL_TICKER As String = "Ticker"
Private Const COL_SPOT_AMOUNT As String = "Spot Amount"
Private Const COL_FUNDING_AMOUNT As String = "Funding Amount"
Private Const COL_EARN_AMOUNT As String = "Earn Amount"
Private Const COL_DYNAMIC_AMOUNT As String = "Dynamic Amount"

Public Function IsDataCleanedUp()
    Dim table As ListObject
    Set table = ThisWorkbook.Sheets(SHEET_NAME).ListObjects(TABLE_NAME)

    IsCleanedUp = (table.ListRows.Count = 0)
End Function

Public Sub CleanUpData()
    Dim table As ListObject
    Set table = ThisWorkbook.Sheets(SHEET_NAME).ListObjects(TABLE_NAME)
    
    If Not table.DataBodyRange Is Nothing Then
        table.DataBodyRange.Delete
    End If

    TableLastUpdate.Wallet = 0
End Sub

Public Sub ClearData()
    Dim table As ListObject
    Set table = ThisWorkbook.Sheets(SHEET_NAME).ListObjects(TABLE_NAME)
    
    table.ListColumns(COL_SPOT_AMOUNT).DataBodyRange.ClearContents
    table.ListColumns(COL_FUNDING_AMOUNT).DataBodyRange.ClearContents
    table.ListColumns(COL_EARN_AMOUNT).DataBodyRange.ClearContents
    table.ListColumns(COL_DYNAMIC_AMOUNT).DataBodyRange.ClearContents

    TableLastUpdate.Wallet = 0
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
    For Each ticker In assets.Wallet.Keys
        
        Dim rowIndex As Long
        rowIndex = xFindTickerRowIndex(ticker, True)
        
        Dim name As Variant
        For Each name In assets.Wallet(ticker).Keys
            table.ListColumns(name).DataBodyRange(rowIndex) = assets.Wallet(ticker)(name)
        Next name
        
    Next ticker
    
    TableLastUpdate.Wallet = now()
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
    End If
    
    xFindTickerRowIndex = index
End Function

