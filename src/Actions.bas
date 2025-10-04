Attribute VB_Name = "Actions"
Option Explicit

Public Sub Action_ClearAssets()
    xActionHeader
    On Error GoTo ErrHandler
    
    ThisWorkbook.Sheets(SHEET_ASSETS).Activate
    TableUserInfo.ClearData
    TableWallet.ClearData
    
    On Error GoTo 0

Finalize:
    xActionFooter
    Exit Sub
ErrHandler:
    HandleError Err.Number, Err.source, Err.Description
    Resume Finalize
End Sub

Public Sub Action_UpdateAssets()
    xActionHeader
    On Error GoTo ErrHandler
    
    ThisWorkbook.Sheets(SHEET_ASSETS).Activate
    TableUserInfo.UpdateData
    TableWallet.UpdateData

    On Error GoTo 0

Finalize:
    xActionFooter
    Exit Sub
ErrHandler:
    HandleError Err.Number, Err.source, Err.Description
    Resume Finalize
End Sub

Public Sub Action_SortAssets()
    xActionHeader
    On Error GoTo ErrHandler
    
    ThisWorkbook.Sheets(SHEET_ASSETS).Activate
    TableWallet.Sort BinanceLedgers.GetOrderString()

    On Error GoTo 0

Finalize:
    xActionFooter
    Exit Sub
ErrHandler:
    HandleError Err.Number, Err.source, Err.Description
    Resume Finalize
End Sub

Public Sub Action_ShowAssets()
    xActionHeader
    On Error GoTo ErrHandler
    
    ThisWorkbook.Sheets(SHEET_ASSETS).Activate

    On Error GoTo 0

Finalize:
    xActionFooter
    Exit Sub
ErrHandler:
    HandleError Err.Number, Err.source, Err.Description
    Resume Finalize
End Sub

Public Sub Action_CollectUsedCurrencies()
    xActionHeader
    On Error GoTo ErrHandler
    
    ThisWorkbook.Sheets(SHEET_CURRENCIES).Activate
    TableCurrencies.UpdateData
    
    BinanceLedgers.RefreshCoinsInfo

    On Error GoTo 0

Finalize:
    xActionFooter
    Exit Sub
ErrHandler:
    HandleError Err.Number, Err.source, Err.Description
    Resume Finalize
End Sub

Public Sub Action_CollectAllCurrencies()
    xActionHeader
    On Error GoTo ErrHandler
    
    ThisWorkbook.Sheets(SHEET_CURRENCIES).Activate
    TableCurrencies.CollectAllData

    BinanceLedgers.RefreshCoinsInfo
    
    On Error GoTo 0

Finalize:
    xActionFooter
    Exit Sub
ErrHandler:
    HandleError Err.Number, Err.source, Err.Description
    Resume Finalize
End Sub

Public Sub Action_CompactCurrencies()
    xActionHeader
    On Error GoTo ErrHandler
    
    ThisWorkbook.Sheets(SHEET_CURRENCIES).Activate
    TableCurrencies.CompactData

    On Error GoTo 0

Finalize:
    xActionFooter
    Exit Sub
ErrHandler:
    HandleError Err.Number, Err.source, Err.Description
    Resume Finalize
End Sub

Public Sub Action_OpenCurrentTickerLedger()
    xActionHeader
    On Error GoTo ErrHandler
    
    BinanceLedgers.OpenCurrentTickerLedger

    On Error GoTo 0

Finalize:
    xActionFooter
    Exit Sub
ErrHandler:
    HandleError Err.Number, Err.source, Err.Description
    Resume Finalize
End Sub

Public Sub Action_UpdateTransfers()
    xActionHeader
    On Error GoTo ErrHandler
    
    Dim ops As New BinanceOps
    ops.Collect_Transfers TableLastUpdate.Transfers
    
    BinanceLedgers.PopulateOperations ops
    
    TableLastUpdate.Transfers = ops.endDate

    On Error GoTo 0

Finalize:
    xActionFooter
    Exit Sub
ErrHandler:
    HandleError Err.Number, Err.source, Err.Description
    Resume Finalize
End Sub

Public Sub Action_UpdateConversions()
    xActionHeader
    On Error GoTo ErrHandler

    Dim ops As New BinanceOps
    ops.Collect_Conversions TableLastUpdate.Convertions
    
    BinanceLedgers.PopulateOperations ops
    
    TableLastUpdate.Convertions = ops.endDate

    xCheckForUnappliedOps ops

    On Error GoTo 0

Finalize:
    xActionFooter
    Exit Sub
ErrHandler:
    HandleError Err.Number, Err.source, Err.Description
    Resume Finalize
End Sub

Public Sub Action_UpdateDustLog()
    xActionHeader
    On Error GoTo ErrHandler

    Dim ops As New BinanceOps
    ops.Collect_DustLog TableLastUpdate.DustLog
    
    BinanceLedgers.PopulateOperations ops
    
    xCheckForUnappliedOps ops
    
    TableLastUpdate.DustLog = ops.endDate

    On Error GoTo 0

Finalize:
    xActionFooter
    Exit Sub
ErrHandler:
    HandleError Err.Number, Err.source, Err.Description
    Resume Finalize
End Sub

Public Sub Action_UpdateFiatBuys()
    xActionHeader
    On Error GoTo ErrHandler
    
    Dim ops As New BinanceOps
    ops.Collect_FiatBuys TableLastUpdate.FiatBuys
    
    BinanceLedgers.PopulateOperations ops
    
    TableLastUpdate.FiatBuys = ops.endDate

    On Error GoTo 0

Finalize:
    xActionFooter
    Exit Sub
ErrHandler:
    HandleError Err.Number, Err.source, Err.Description
    Resume Finalize
End Sub

Public Sub Action_UpdateFiatSells()
    xActionHeader
    On Error GoTo ErrHandler
    
    Dim ops As New BinanceOps
    ops.Collect_FiatSells TableLastUpdate.FiatSells
    
    BinanceLedgers.PopulateOperations ops
    
    TableLastUpdate.FiatSells = ops.endDate

    On Error GoTo 0

Finalize:
    xActionFooter
    Exit Sub
ErrHandler:
    HandleError Err.Number, Err.source, Err.Description
    Resume Finalize
End Sub

Public Sub Action_UpdateFiatDeposits()
    xActionHeader
    On Error GoTo ErrHandler
    
    Dim ops As New BinanceOps
    ops.Collect_FiatDeposits TableLastUpdate.FiatDeposits
    
    BinanceLedgers.PopulateOperations ops
    
    TableLastUpdate.FiatDeposits = ops.endDate

    On Error GoTo 0

Finalize:
    xActionFooter
    Exit Sub
ErrHandler:
    HandleError Err.Number, Err.source, Err.Description
    Resume Finalize
End Sub

Public Sub Action_UpdateFiatWithdraws()
    xActionHeader
    On Error GoTo ErrHandler
    
    Dim ops As New BinanceOps
    ops.Collect_FiatWithdraws TableLastUpdate.FiatWithdraws
    
    BinanceLedgers.PopulateOperations ops
    
    TableLastUpdate.FiatWithdraws = ops.endDate

    On Error GoTo 0

Finalize:
    xActionFooter
    Exit Sub
ErrHandler:
    HandleError Err.Number, Err.source, Err.Description
    Resume Finalize
End Sub

Public Sub Action_UpdateCryptoDeposits()
    xActionHeader
    On Error GoTo ErrHandler
    
    Dim ops As New BinanceOps
    ops.Collect_CryptoDeposits TableLastUpdate.CryptoDeposits
    
    BinanceLedgers.PopulateOperations ops
    
    TableLastUpdate.CryptoDeposits = ops.endDate

    On Error GoTo 0

Finalize:
    xActionFooter
    Exit Sub
ErrHandler:
    HandleError Err.Number, Err.source, Err.Description
    Resume Finalize
End Sub

Public Sub Action_UpdateCryptoWithdraws()
    xActionHeader
    On Error GoTo ErrHandler
    
    Dim ops As New BinanceOps
    ops.Collect_CryptoWithdraws TableLastUpdate.CryptoWithdraws
    
    BinanceLedgers.PopulateOperations ops
    
    TableLastUpdate.CryptoWithdraws = ops.endDate

    On Error GoTo 0

Finalize:
    xActionFooter
    Exit Sub
ErrHandler:
    HandleError Err.Number, Err.source, Err.Description
    Resume Finalize
End Sub

Public Sub Action_UpdateCryptoDistributions()
    xActionHeader
    On Error GoTo ErrHandler
    
    Dim ops As New BinanceOps
    ops.Collect_Distributions TableLastUpdate.Distributions
    
    BinanceLedgers.PopulateOperations ops
    
    TableLastUpdate.Distributions = ops.endDate

    On Error GoTo 0

Finalize:
    xActionFooter
    Exit Sub
ErrHandler:
    HandleError Err.Number, Err.source, Err.Description
    Resume Finalize
End Sub

Public Sub Action_UpdateFlexibleEarn()
    xActionHeader
    On Error GoTo ErrHandler
    
    Dim ops As New BinanceOps
    ops.Collect_FlexibleEarnSubscriptions TableLastUpdate.FlexibleEarns
    ops.Collect_FlexibleEarnRedemptions TableLastUpdate.FlexibleEarns
    
    BinanceLedgers.PopulateOperations ops
    
    TableLastUpdate.FlexibleEarns = ops.endDate

    On Error GoTo 0

Finalize:
    xActionFooter
    Exit Sub
ErrHandler:
    HandleError Err.Number, Err.source, Err.Description
    Resume Finalize
End Sub

Public Sub Action_UpdateLockedEarn()
    xActionHeader
    On Error GoTo ErrHandler
    
    Dim ops As New BinanceOps
    ops.Collect_LockedEarnSubscriptions TableLastUpdate.LockedEarns
    ops.Collect_LockedEarnRedemptions TableLastUpdate.LockedEarns
    
    BinanceLedgers.PopulateOperations ops
    
    TableLastUpdate.LockedEarns = ops.endDate

    On Error GoTo 0

Finalize:
    xActionFooter
    Exit Sub
ErrHandler:
    HandleError Err.Number, Err.source, Err.Description
    Resume Finalize
End Sub

Public Sub Action_UpdateTrades()
    xActionHeader
    On Error GoTo ErrHandler
    
    BinanceLedgers.UpdateCurrentLedgerTrades
    
    On Error GoTo 0

Finalize:
    xActionFooter
    Exit Sub
ErrHandler:
    HandleError Err.Number, Err.source, Err.Description
    Resume Finalize
End Sub

Public Sub Action_CheckUnappliedOperations()
    xActionHeader
    On Error GoTo ErrHandler
    
    Dim ops As New BinanceOps
    
    TableConversions.CollectUnappliedOps ops
    TableDustLog.CollectUnappliedOps ops
    
    BinanceLedgers.PopulateOperations ops, True
    
    TableLastUpdate.LockedEarns = ops.endDate

    On Error GoTo 0

Finalize:
    xActionFooter
    Exit Sub
ErrHandler:
    HandleError Err.Number, Err.source, Err.Description
    Resume Finalize
End Sub

Private Sub xActionHeader()

End Sub

Private Sub xActionFooter()

End Sub

Private Sub xCheckForUnappliedOps(ByVal ops As BinanceOps)
    If ops.unappliedOps Then
        ThisWorkbook.Sheets(SHEET_MISC).Activate
        
        TableConversions.Sort
        TableDustLog.Sort
        
        MsgBox "There is unapplied operations. Take care of them."
    End If
End Sub

