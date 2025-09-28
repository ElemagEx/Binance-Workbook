Attribute VB_Name = "Actions"
Option Explicit

Public Sub Action_ClearAssets()
    ThisWorkbook.Sheets(SHEET_ASSETS).Activate
    TableUserInfo.ClearData
    TableWallet.ClearData
End Sub

Public Sub Action_UpdateAssets()
    ThisWorkbook.Sheets(SHEET_ASSETS).Activate
    TableUserInfo.UpdateData
    TableWallet.UpdateData
End Sub

Public Sub Action_CollectUsedCurrencies()
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

Public Sub Action_UpdateTransfers()

End Sub

Public Sub Action_UpdateConversions()

End Sub

Public Sub Action_UpdateDustLog()

End Sub

Public Sub Action_UpdateFiatBuys()
    Dim ops As New BinanceOps
    
    ops.Collect_FiatBuys TableLastUpdate.FiatBuy
    
    BinanceLedgers.PopulateOperations ops
    
    TableLastUpdate.FiatBuy = ops.endDate
End Sub

Public Sub Action_UpdateFiatSells()
    Dim ops As New BinanceOps
    
    ops.Collect_FiatSells TableLastUpdate.FiatSell
    
    BinanceLedgers.PopulateOperations ops
    
    TableLastUpdate.FiatSell = ops.endDate
End Sub

Public Sub Action_UpdateFiatDeposits()
    Dim ops As New BinanceOps
    
    ops.Collect_FiatDeposits TableLastUpdate.FiatDeposit
    
    BinanceLedgers.PopulateOperations ops
    
    TableLastUpdate.FiatDeposit = ops.endDate
End Sub

Public Sub Action_UpdateFiatWithdraws()
    Dim ops As New BinanceOps
    
    ops.Collect_FiatWithdraws TableLastUpdate.FiatWithdraw
    
    BinanceLedgers.PopulateOperations ops
    
    TableLastUpdate.FiatWithdraw = ops.endDate
End Sub

Public Sub Action_UpdateCryptoDeposits()
    Dim ops As New BinanceOps
    
    ops.Collect_CryptoDeposits TableLastUpdate.CryptoDeposit
    
    BinanceLedgers.PopulateOperations ops
    
    TableLastUpdate.CryptoDeposit = ops.endDate
End Sub

Public Sub Action_UpdateCryptoWithdraws()
    Dim ops As New BinanceOps
    
    ops.Collect_CryptoWithdraws TableLastUpdate.CryptoWithdraw
    
    BinanceLedgers.PopulateOperations ops
    
    TableLastUpdate.CryptoWithdraw = ops.endDate
End Sub

Public Sub Action_UpdateCryptoDistributions()
    Dim ops As New BinanceOps
    
    ops.Collect_Distributions TableLastUpdate.Distribution
    
    BinanceLedgers.PopulateOperations ops
    
    TableLastUpdate.Distribution = ops.endDate
End Sub

Public Sub Action_UpdateFlexibleEarn()

End Sub

Public Sub Action_UpdateLockedEarn()

End Sub
