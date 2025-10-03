Attribute VB_Name = "TableConversions"
Option Explicit
Option Private Module

Private Const ID_PREFIX As String = "XX-"

Private Const COL_STAMP As String = "Stamp"
Private Const COL_APPLIED As String = "Applied"
Private Const COL_CONVERSION_ID As String = "Conversion ID"
Private Const COL_FROM_ASSET = "From Asset"
Private Const COL_FROM_AMOUNT = "From Amount"
Private Const COL_SPOT_WALLET = "Spot Wallet"
Private Const COL_FUNDING_WALLET = "Funding Wallet"
Private Const COL_EARN_WALLET = "Earn Wallet"
Private Const COL_IS_VALID = "Is Valid"
Private Const COL_TO_ASSET = "To Asset"
Private Const COL_TO_AMOUNT = "To Amount"

Public Function IsDataCleanedUp() As Boolean
    IsDataCleanedUp = (xGetTable().ListRows.count = 0)
End Function

Public Sub CleanUpData()
    Dim table As ListObject
    Set table = xGetTable()
    
    If Not table.DataBodyRange Is Nothing Then
        table.DataBodyRange.Delete
    End If
End Sub

Public Sub HandleComboConversion(ByVal ops As BinanceOps, ByVal item As Dictionary, ByVal stamp As Date, ByVal spotWallet As Boolean, ByVal fundingWallet As Boolean, ByVal earnWallet As Boolean)
    Dim table As ListObject
    Set table = xGetTable()
    
    Dim id As String
    id = ID_PREFIX & item("orderId")
    
    Dim rowIndex As Long
    rowIndex = xFindConversionRowIndex(id)
    
    If rowIndex <> 0 Then
        xCollectWalletsTransfers ops, rowIndex, True
    Else
        xAddNewComboConversion id, item, stamp, spotWallet, fundingWallet, earnWallet
        ops.unappliedOps = True
    End If
End Sub

Public Sub CollectUnappliedOps(ByVal ops As BinanceOps)
    Dim table As ListObject
    Set table = xGetTable()
    
    Dim rowIndex As Long
    For rowIndex = 1 To table.ListRows.count
        xCollectWalletsTransfers ops, rowIndex, False
    Next rowIndex
End Sub

Public Sub Sort()
    Dim table As ListObject
    Set table = xGetTable()
    
    Dim col As ListColumn
    Set col = table.ListColumns(COL_STAMP)
    
    If Not col.DataBodyRange Is Nothing Then
        With table.Sort
            .SortFields.Clear
            .SortFields.Add key:=col.DataBodyRange, SortOn:=xlSortOnValues, order:=xlAscending, DataOption:=xlSortNormal
            .header = xlYes
            .Apply
        End With
    End If
End Sub

Private Function xGetTable() As ListObject
    Set xGetTable = ThisWorkbook.Sheets(SHEET_MISC).ListObjects(TABLE_CONVERSIONS)
End Function

Private Function xFindConversionRowIndex(ByVal id As String) As Long
    xFindConversionRowIndex = 0
    
    Dim table As ListObject
    Set table = xGetTable()
    
    If table.ListRows.count > 0 Then
        Dim cell As Range
        Set cell = table.ListColumns(COL_CONVERSION_ID).DataBodyRange.Find(what:=id, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    
        If Not cell Is Nothing Then
            xFindConversionRowIndex = cell.Row - table.HeaderRowRange.Row
        End If
    End If
End Function

Private Sub xAddNewComboConversion(ByVal id As String, ByVal item As Dictionary, ByVal stamp As Date, ByVal spotWallet As Boolean, ByVal fundingWallet As Boolean, ByVal earnWallet As Boolean)
    Dim table As ListObject
    Set table = xGetTable()

    Dim rowIndex As Long
    rowIndex = table.ListRows.Add().index
    
    Dim amount As Variant
    amount = item("fromAmount")

    table.ListColumns(COL_STAMP).DataBodyRange(rowIndex).Value = stamp
    table.ListColumns(COL_CONVERSION_ID).DataBodyRange(rowIndex).Value = id
    table.ListColumns(COL_FROM_ASSET).DataBodyRange(rowIndex).Value = item("fromAsset")
    table.ListColumns(COL_FROM_AMOUNT).DataBodyRange(rowIndex).Value = amount
    table.ListColumns(COL_SPOT_WALLET).DataBodyRange(rowIndex).Value = IIf(spotWallet, amount, " ")
    table.ListColumns(COL_FUNDING_WALLET).DataBodyRange(rowIndex).Value = IIf(fundingWallet, IIf(spotWallet, 0, amount), " ")
    table.ListColumns(COL_EARN_WALLET).DataBodyRange(rowIndex).Value = IIf(earnWallet, 0, " ")
    table.ListColumns(COL_TO_ASSET).DataBodyRange(rowIndex).Value = item("toAsset")
    table.ListColumns(COL_TO_AMOUNT).DataBodyRange(rowIndex).Value = item("toAmount")
End Sub

Private Sub xCollectWalletsTransfers(ByVal ops As BinanceOps, ByVal rowIndex As Long, ByVal checkApplience As Boolean)
    Dim table As ListObject
    Set table = xGetTable()
    
    If checkApplience <> table.ListColumns(COL_APPLIED).DataBodyRange(rowIndex).Value Then
        ops.unappliedOps = checkApplience And True
        Exit Sub
    End If
    If Not table.ListColumns(COL_IS_VALID).DataBodyRange(rowIndex).Value Then
        ops.unappliedOps = checkApplience And True
        Exit Sub
    End If
    
    If Not IsNumeric(table.ListColumns(COL_SPOT_WALLET).DataBodyRange(rowIndex).Value) Then
        xAddWalletTransfer ops, rowIndex, WALLET_FLEXIBLE_EARN, WALLET_FUNDING, ID_PREFIX_CONVERT_EFO, ID_PREFIX_CONVERT_EFI, COL_EARN_WALLET
    ElseIf Not IsNumeric(table.ListColumns(COL_FUNDING_WALLET).DataBodyRange(rowIndex).Value) Then
        xAddWalletTransfer ops, rowIndex, WALLET_FLEXIBLE_EARN, WALLET_SPOT, ID_PREFIX_CONVERT_ESO, ID_PREFIX_CONVERT_ESI, COL_EARN_WALLET
    ElseIf Not IsNumeric(table.ListColumns(COL_EARN_WALLET).DataBodyRange(rowIndex).Value) Then
        xAddWalletTransfer ops, rowIndex, WALLET_FUNDING, WALLET_SPOT, ID_PREFIX_CONVERT_FSO, ID_PREFIX_CONVERT_FSI, COL_FUNDING_WALLET
    Else
        xAddWalletTransfer ops, rowIndex, WALLET_FUNDING, WALLET_SPOT, ID_PREFIX_CONVERT_FSO, ID_PREFIX_CONVERT_FSI, COL_FUNDING_WALLET
        xAddWalletTransfer ops, rowIndex, WALLET_FLEXIBLE_EARN, WALLET_SPOT, ID_PREFIX_CONVERT_ESO, ID_PREFIX_CONVERT_ESI, COL_EARN_WALLET
    End If
    
    If Not checkApplience Then
        table.ListColumns(COL_APPLIED).DataBodyRange(rowIndex).Value = True
    End If
End Sub

Private Sub xAddWalletTransfer( _
    ByVal ops As BinanceOps, _
    ByVal rowIndex As Long, _
    ByVal walletOut As String, _
    ByVal walletIn As String, _
    ByVal prefixOut As String, _
    ByVal prefixIn As String, _
    ByVal amountColName As String _
    )
    Dim table As ListObject
    Set table = xGetTable()

    Dim amount As Variant
    amount = table.ListColumns(amountColName).DataBodyRange(rowIndex).Value

    Dim ticker As String
    ticker = table.ListColumns(COL_FROM_ASSET).DataBodyRange(rowIndex).Value
    
    Dim id As String
    id = Mid(table.ListColumns(COL_CONVERSION_ID).DataBodyRange(rowIndex).Value, 4)

    Dim stamp As Date
    stamp = table.ListColumns(COL_STAMP).DataBodyRange(rowIndex).Value

    ops.AddSimpleOp ticker, stamp, walletOut, OP_WALLET_OUT, 0, amount, 0, "to:" & walletIn, prefixOut & id
    ops.AddSimpleOp ticker, stamp, walletIn, OP_WALLET_IN, amount, 0, 0, "from:" & walletOut, prefixIn & id
End Sub
