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
    Dim ops As New BinanceOps
    
    ops.Collect_Transfers TableLastUpdate.Transfers
    
    BinanceLedgers.PopulateOperations ops
    
    TableLastUpdate.Transfers = ops.endDate
End Sub

Public Sub Action_UpdateConversions()

End Sub

Public Sub Action_UpdateDustLog()

End Sub

Public Sub Action_UpdateFiatBuys()
    Dim ops As New BinanceOps
    
    ops.Collect_FiatBuys TableLastUpdate.FiatBuys
    
    BinanceLedgers.PopulateOperations ops
    
    TableLastUpdate.FiatBuys = ops.endDate
End Sub

Public Sub Action_UpdateFiatSells()
    Dim ops As New BinanceOps
    
    ops.Collect_FiatSells TableLastUpdate.FiatSells
    
    BinanceLedgers.PopulateOperations ops
    
    TableLastUpdate.FiatSells = ops.endDate
End Sub

Public Sub Action_UpdateFiatDeposits()
    Dim ops As New BinanceOps
    
    ops.Collect_FiatDeposits TableLastUpdate.FiatDeposits
    
    BinanceLedgers.PopulateOperations ops
    
    TableLastUpdate.FiatDeposits = ops.endDate
End Sub

Public Sub Action_UpdateFiatWithdraws()
    Dim ops As New BinanceOps
    
    ops.Collect_FiatWithdraws TableLastUpdate.FiatWithdraws
    
    BinanceLedgers.PopulateOperations ops
    
    TableLastUpdate.FiatWithdraws = ops.endDate
End Sub

Public Sub Action_UpdateCryptoDeposits()
    Dim ops As New BinanceOps
    
    ops.Collect_CryptoDeposits TableLastUpdate.CryptoDeposits
    
    BinanceLedgers.PopulateOperations ops
    
    TableLastUpdate.CryptoDeposits = ops.endDate
End Sub

Public Sub Action_UpdateCryptoWithdraws()
    Dim ops As New BinanceOps
    
    ops.Collect_CryptoWithdraws TableLastUpdate.CryptoWithdraws
    
    BinanceLedgers.PopulateOperations ops
    
    TableLastUpdate.CryptoWithdraws = ops.endDate
End Sub

Public Sub Action_UpdateCryptoDistributions()
    Dim ops As New BinanceOps
    
    ops.Collect_Distributions TableLastUpdate.Distributions
    
    BinanceLedgers.PopulateOperations ops
    
    TableLastUpdate.Distributions = ops.endDate
End Sub

Public Sub Action_UpdateFlexibleEarn()
    Dim ops As New BinanceOps
    
    ops.Collect_FlexibleEarnSubscriptions TableLastUpdate.FlexibleEarns
    ops.Collect_FlexibleEarnRedemptions TableLastUpdate.FlexibleEarns
    
    BinanceLedgers.PopulateOperations ops
    
    TableLastUpdate.FlexibleEarns = ops.endDate
End Sub

Public Sub Action_UpdateLockedEarn()
    Dim ops As New BinanceOps
    
    ops.Collect_LockedEarnSubscriptions TableLastUpdate.LockedEarns
    ops.Collect_LockedEarnRedemptions TableLastUpdate.LockedEarns
    
    BinanceLedgers.PopulateOperations ops
    
    TableLastUpdate.LockedEarns = ops.endDate
End Sub
