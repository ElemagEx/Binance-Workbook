Attribute VB_Name = "Actions"
Option Explicit

Public Sub Action_ClearWallet()
    TableWallet.ClearData
End Sub

Public Sub Action_UpdateWallet()
    TableWallet.UpdateData
End Sub

Public Sub Action_UpdateFiatBuys()
    Dim ops As New BinanceOps
    
    ops.Collect_FiatBuys TableLastUpdate.FiatBuy
    
    BinanceLedgers.PopulateOperations ops
End Sub
