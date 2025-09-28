Attribute VB_Name = "Actions"
Option Explicit

Public Sub Action_ClearWallet()
    ThisWorkbook.Sheets(SHEET_ASSETS).Activate
    TableWallet.ClearData
End Sub

Public Sub Action_UpdateWallet()
    ThisWorkbook.Sheets(SHEET_ASSETS).Activate
    TableWallet.UpdateData
End Sub

Public Sub Action_UpdateUsedCurrencies()
    ThisWorkbook.Sheets(SHEET_CURRENCIES).Activate
    TableCurrencies.UpdateData
End Sub

Public Sub Action_CollectAllCurrencies()
    ThisWorkbook.Sheets(SHEET_CURRENCIES).Activate
    TableCurrencies.CollectAllData
End Sub

Public Sub Action_CompactCurrencies()
    ThisWorkbook.Sheets(SHEET_CURRENCIES).Activate
    TableCurrencies.CompactData
End Sub
Public Sub Action_UpdateFiatBuys()
    Dim ops As New BinanceOps
    
    ops.Collect_FiatBuys TableLastUpdate.FiatBuy
    
    BinanceLedgers.PopulateOperations ops
End Sub

