Attribute VB_Name = "Actions"
Option Explicit

#Const IN_DEVELOPMENT = True

Private Const HANDLE_ERRORS As Boolean = False

Public Sub Action_ClearAssets()
    xActionHeader
    If HANDLE_ERRORS Then On Error GoTo ErrHandler

    ThisWorkbook.Sheets(SHEET_ASSETS).Activate
    TableUserInfo.ClearData
    TableAssets.ClearData
    
    If HANDLE_ERRORS Then On Error GoTo 0

Finalize:
    xActionFooter
    Exit Sub
ErrHandler:
    HandleError Err.Number, Err.source, Err.Description
    Resume Finalize
End Sub

Public Sub Action_UpdateAssets()
    xActionHeader
    If HANDLE_ERRORS Then On Error GoTo ErrHandler
    
    ThisWorkbook.Sheets(SHEET_ASSETS).Activate
    TableUserInfo.UpdateData
    TableAssets.UpdateData
    TableAssets.Evaluate

    If HANDLE_ERRORS Then On Error GoTo 0

Finalize:
    xActionFooter
    Exit Sub
ErrHandler:
    HandleError Err.Number, Err.source, Err.Description
    Resume Finalize
End Sub

Public Sub Action_SortAssets()
    xActionHeader
    If HANDLE_ERRORS Then On Error GoTo ErrHandler
    
    ThisWorkbook.Sheets(SHEET_ASSETS).Activate
    TableAssets.Sort BinanceLedgers.GetOrderString()

    If HANDLE_ERRORS Then On Error GoTo 0

Finalize:
    xActionFooter
    Exit Sub
ErrHandler:
    HandleError Err.Number, Err.source, Err.Description
    Resume Finalize
End Sub

Public Sub Action_ShowAssets()
    xActionHeader
    If HANDLE_ERRORS Then On Error GoTo ErrHandler
    
    ThisWorkbook.Sheets(SHEET_ASSETS).Activate

    If HANDLE_ERRORS Then On Error GoTo 0

Finalize:
    xActionFooter
    Exit Sub
ErrHandler:
    HandleError Err.Number, Err.source, Err.Description
    Resume Finalize
End Sub

Public Sub Action_CollectUsedCurrencies()
    xActionHeader
    If HANDLE_ERRORS Then On Error GoTo ErrHandler
    
    ThisWorkbook.Sheets(SHEET_CURRENCIES).Activate
    TableCurrencies.UpdateData
    
    BinanceLedgers.RefreshCoinsInfo

    If HANDLE_ERRORS Then On Error GoTo 0

Finalize:
    xActionFooter
    Exit Sub
ErrHandler:
    HandleError Err.Number, Err.source, Err.Description
    Resume Finalize
End Sub

Public Sub Action_CollectAllCurrencies()
    xActionHeader
    If HANDLE_ERRORS Then On Error GoTo ErrHandler
    
    ThisWorkbook.Sheets(SHEET_CURRENCIES).Activate
    TableCurrencies.CollectAllData

    BinanceLedgers.RefreshCoinsInfo
    
    If HANDLE_ERRORS Then On Error GoTo 0

Finalize:
    xActionFooter
    Exit Sub
ErrHandler:
    HandleError Err.Number, Err.source, Err.Description
    Resume Finalize
End Sub

Public Sub Action_CompactCurrencies()
    xActionHeader
    If HANDLE_ERRORS Then On Error GoTo ErrHandler
    
    ThisWorkbook.Sheets(SHEET_CURRENCIES).Activate
    TableCurrencies.CompactData

    If HANDLE_ERRORS Then On Error GoTo 0

Finalize:
    xActionFooter
    Exit Sub
ErrHandler:
    HandleError Err.Number, Err.source, Err.Description
    Resume Finalize
End Sub

Public Sub Action_OpenCurrentTickerLedger()
    xActionHeader
    If HANDLE_ERRORS Then On Error GoTo ErrHandler
    
    BinanceLedgers.OpenCurrentTickerLedger

    If HANDLE_ERRORS Then On Error GoTo 0

Finalize:
    xActionFooter
    Exit Sub
ErrHandler:
    HandleError Err.Number, Err.source, Err.Description
    Resume Finalize
End Sub

Public Sub Action_UpdateTransfers()
    xActionHeader
    If HANDLE_ERRORS Then On Error GoTo ErrHandler
    
    Dim ops As New BinanceOps
    ops.Collect_Transfers TableLastUpdate.Transfers
    
    BinanceLedgers.PopulateOperations ops
    
    TableLastUpdate.Transfers = ops.endDate

    If HANDLE_ERRORS Then On Error GoTo 0

Finalize:
    xActionFooter
    Exit Sub
ErrHandler:
    HandleError Err.Number, Err.source, Err.Description
    Resume Finalize
End Sub

Public Sub Action_UpdateConversions()
    xActionHeader
    If HANDLE_ERRORS Then On Error GoTo ErrHandler

    Dim ops As New BinanceOps
    ops.Collect_Conversions TableLastUpdate.Convertions
    
    BinanceLedgers.PopulateOperations ops
    
    TableLastUpdate.Convertions = ops.endDate

    xCheckForUnappliedOps ops

    If HANDLE_ERRORS Then On Error GoTo 0

Finalize:
    xActionFooter
    Exit Sub
ErrHandler:
    HandleError Err.Number, Err.source, Err.Description
    Resume Finalize
End Sub

Public Sub Action_UpdateDustLog()
    xActionHeader
    If HANDLE_ERRORS Then On Error GoTo ErrHandler

    Dim ops As New BinanceOps
    ops.Collect_DustLog TableLastUpdate.DustLog
    
    BinanceLedgers.PopulateOperations ops
    
    xCheckForUnappliedOps ops
    
    TableLastUpdate.DustLog = ops.endDate

    If HANDLE_ERRORS Then On Error GoTo 0

Finalize:
    xActionFooter
    Exit Sub
ErrHandler:
    HandleError Err.Number, Err.source, Err.Description
    Resume Finalize
End Sub

Public Sub Action_UpdateFiatBuys()
    xActionHeader
    If HANDLE_ERRORS Then On Error GoTo ErrHandler
    
    Dim ops As New BinanceOps
    ops.Collect_FiatBuys TableLastUpdate.FiatBuys
    
    BinanceLedgers.PopulateOperations ops
    
    TableLastUpdate.FiatBuys = ops.endDate

    If HANDLE_ERRORS Then On Error GoTo 0

Finalize:
    xActionFooter
    Exit Sub
ErrHandler:
    HandleError Err.Number, Err.source, Err.Description
    Resume Finalize
End Sub

Public Sub Action_UpdateFiatSells()
    xActionHeader
    If HANDLE_ERRORS Then On Error GoTo ErrHandler
    
    Dim ops As New BinanceOps
    ops.Collect_FiatSells TableLastUpdate.FiatSells
    
    BinanceLedgers.PopulateOperations ops
    
    TableLastUpdate.FiatSells = ops.endDate

    If HANDLE_ERRORS Then On Error GoTo 0

Finalize:
    xActionFooter
    Exit Sub
ErrHandler:
    HandleError Err.Number, Err.source, Err.Description
    Resume Finalize
End Sub

Public Sub Action_UpdateFiatDeposits()
    xActionHeader
    If HANDLE_ERRORS Then On Error GoTo ErrHandler
    
    Dim ops As New BinanceOps
    ops.Collect_FiatDeposits TableLastUpdate.FiatDeposits
    
    BinanceLedgers.PopulateOperations ops
    
    TableLastUpdate.FiatDeposits = ops.endDate

    If HANDLE_ERRORS Then On Error GoTo 0

Finalize:
    xActionFooter
    Exit Sub
ErrHandler:
    HandleError Err.Number, Err.source, Err.Description
    Resume Finalize
End Sub

Public Sub Action_UpdateFiatWithdraws()
    xActionHeader
    If HANDLE_ERRORS Then On Error GoTo ErrHandler
    
    Dim ops As New BinanceOps
    ops.Collect_FiatWithdraws TableLastUpdate.FiatWithdraws
    
    BinanceLedgers.PopulateOperations ops
    
    TableLastUpdate.FiatWithdraws = ops.endDate

    If HANDLE_ERRORS Then On Error GoTo 0

Finalize:
    xActionFooter
    Exit Sub
ErrHandler:
    HandleError Err.Number, Err.source, Err.Description
    Resume Finalize
End Sub

Public Sub Action_UpdateCryptoDeposits()
    xActionHeader
    If HANDLE_ERRORS Then On Error GoTo ErrHandler
    
    Dim ops As New BinanceOps
    ops.Collect_CryptoDeposits TableLastUpdate.CryptoDeposits
    
    BinanceLedgers.PopulateOperations ops
    
    TableLastUpdate.CryptoDeposits = ops.endDate

    If HANDLE_ERRORS Then On Error GoTo 0

Finalize:
    xActionFooter
    Exit Sub
ErrHandler:
    HandleError Err.Number, Err.source, Err.Description
    Resume Finalize
End Sub

Public Sub Action_UpdateCryptoWithdraws()
    xActionHeader
    If HANDLE_ERRORS Then On Error GoTo ErrHandler
    
    Dim ops As New BinanceOps
    ops.Collect_CryptoWithdraws TableLastUpdate.CryptoWithdraws
    
    BinanceLedgers.PopulateOperations ops
    
    TableLastUpdate.CryptoWithdraws = ops.endDate

    If HANDLE_ERRORS Then On Error GoTo 0

Finalize:
    xActionFooter
    Exit Sub
ErrHandler:
    HandleError Err.Number, Err.source, Err.Description
    Resume Finalize
End Sub

Public Sub Action_UpdateCryptoDistributions()
    xActionHeader
    If HANDLE_ERRORS Then On Error GoTo ErrHandler
    
    Dim ops As New BinanceOps
    ops.Collect_Distributions TableLastUpdate.Distributions
    
    BinanceLedgers.PopulateOperations ops
    
    TableLastUpdate.Distributions = ops.endDate

    If HANDLE_ERRORS Then On Error GoTo 0

Finalize:
    xActionFooter
    Exit Sub
ErrHandler:
    HandleError Err.Number, Err.source, Err.Description
    Resume Finalize
End Sub

Public Sub Action_UpdateFlexibleEarn()
    xActionHeader
    If HANDLE_ERRORS Then On Error GoTo ErrHandler
    
    Dim ops As New BinanceOps
    ops.Collect_FlexibleEarnSubscriptions TableLastUpdate.FlexibleEarns
    ops.Collect_FlexibleEarnRedemptions TableLastUpdate.FlexibleEarns
    
    BinanceLedgers.PopulateOperations ops
    
    TableLastUpdate.FlexibleEarns = ops.endDate

    If HANDLE_ERRORS Then On Error GoTo 0

Finalize:
    xActionFooter
    Exit Sub
ErrHandler:
    HandleError Err.Number, Err.source, Err.Description
    Resume Finalize
End Sub

Public Sub Action_UpdateLockedEarn()
    xActionHeader
    If HANDLE_ERRORS Then On Error GoTo ErrHandler
    
    Dim ops As New BinanceOps
    ops.Collect_LockedEarnSubscriptions TableLastUpdate.LockedEarns
    ops.Collect_LockedEarnRedemptions TableLastUpdate.LockedEarns
    
    BinanceLedgers.PopulateOperations ops
    
    TableLastUpdate.LockedEarns = ops.endDate

    If HANDLE_ERRORS Then On Error GoTo 0

Finalize:
    xActionFooter
    Exit Sub
ErrHandler:
    HandleError Err.Number, Err.source, Err.Description
    Resume Finalize
End Sub

Public Sub Action_UpdateTrades()
    xActionHeader
    If HANDLE_ERRORS Then On Error GoTo ErrHandler
    
    BinanceLedgers.UpdateCurrentLedgerTrades
    
    If HANDLE_ERRORS Then On Error GoTo 0

Finalize:
    xActionFooter
    Exit Sub
ErrHandler:
    HandleError Err.Number, Err.source, Err.Description
    Resume Finalize
End Sub

Public Sub Action_EvaluateLedger()
    xActionHeader
    If HANDLE_ERRORS Then On Error GoTo ErrHandler
    
    BinanceLedgers.EvaluateCurrentLedgerCosts
    
    If HANDLE_ERRORS Then On Error GoTo 0

Finalize:
    xActionFooter
    Exit Sub
ErrHandler:
    HandleError Err.Number, Err.source, Err.Description
    Resume Finalize
End Sub

Public Sub Action_CheckUnappliedOperations()
    xActionHeader
    If HANDLE_ERRORS Then On Error GoTo ErrHandler
    
    Dim ops As New BinanceOps
    
    TableConversions.CollectUnappliedOps ops
    TableDustLog.CollectUnappliedOps ops
    
    BinanceLedgers.PopulateOperations ops, True
    
    If HANDLE_ERRORS Then On Error GoTo 0

Finalize:
    xActionFooter
    Exit Sub
ErrHandler:
    HandleError Err.Number, Err.source, Err.Description
    Resume Finalize
End Sub

Public Sub Action_RebuildEvalTickers()
    xActionHeader
    If HANDLE_ERRORS Then On Error GoTo ErrHandler

    If HANDLE_ERRORS Then On Error GoTo 0

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

