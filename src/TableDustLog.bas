Attribute VB_Name = "TableDustLog"
Option Explicit
Option Private Module

Private Const ID_PREFIX As String = "YY-"

Private Const COL_COIN_PREFIX = "Coin "
Private Const COL_AMOUNT_PREFIX = "Amount "
Private Const COL_CHARGE_PREFIX = "Charge "
Private Const COL_VOLUME_PREFIX = "Volume "

Private Const COL_STAMP As String = "Stamp"
Private Const COL_APPLIED As String = "Applied"
Private Const COL_TRANSACTION_ID As String = "Transaction ID"
Private Const COL_TICKER As String = "Ticker"
Private Const COL_COIN_1 As String = COL_COIN_PREFIX & 1
Private Const COL_AMOUNT_1 As String = COL_AMOUNT_PREFIX & 1
Private Const COL_CHARGE_1 As String = COL_CHARGE_PREFIX & 1
Private Const COL_VOLUME_1 As String = COL_VOLUME_PREFIX & 1

Private Const COL_DIFF = COL_TICKER ' (columns.Count - COL_DIFF) / 4 == <number-of-coins>
Private Const COL_LAST = COL_VOLUME_1

Public Function IsDataCleanedUp() As Boolean
    IsDataCleanedUp = (xGetTable().ListRows.count = 0)
End Function

Public Sub CleanUpData()
    Dim table As ListObject
    Set table = xGetTable()
    
    If Not table.DataBodyRange Is Nothing Then
        table.DataBodyRange.Delete
    End If
    
    Dim i As Long
    For i = table.ListColumns.count To table.ListColumns(COL_LAST).index + 1 Step -1
        table.ListColumns(i).Delete
    Next i
End Sub

Public Sub HandleTransaction(ByVal ops As BinanceOps, ByVal item As Dictionary, ByVal stamp As Date)
    Dim table As ListObject
    Set table = xGetTable()
    
    Dim id As String
    id = ID_PREFIX & item("transId")
    
    Dim rowIndex As Long
    rowIndex = xFindTransactionRowIndex(id)
    
    If rowIndex <> 0 Then
        xCollectConversions ops, rowIndex, True
    Else
        xAddNewTransaction id, item, stamp
        ops.unappliedOps = True
    End If
End Sub

Public Sub CollectUnappliedOps(ByVal ops As BinanceOps)
    Dim table As ListObject
    Set table = xGetTable()
    
    Dim rowIndex As Long
    For rowIndex = 1 To table.ListRows.count
        xCollectConversions ops, rowIndex, False
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
            .SortFields.Add key:=col.DataBodyRange, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .header = xlYes
            .Apply
        End With
    End If
End Sub

Private Function xGetTable() As ListObject
    Set xGetTable = ThisWorkbook.Sheets(SHEET_MISC).ListObjects(TABLE_DUST_LOG)
End Function

Private Function xFindTransactionRowIndex(ByVal id As String) As Long
    xFindTransactionRowIndex = 0
    
    Dim table As ListObject
    Set table = xGetTable()
    
    If table.ListRows.count > 0 Then
        Dim cell As Range
        Set cell = table.ListColumns(COL_TRANSACTION_ID).DataBodyRange.Find(what:=id, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    
        If Not cell Is Nothing Then
            xFindTransactionRowIndex = cell.Row - table.HeaderRowRange.Row
        End If
    End If
End Function

Private Sub xAddNewTransaction(ByVal id As String, ByVal item As Dictionary, ByVal stamp As Date)
    Dim table As ListObject
    Set table = xGetTable()
    
    Dim dribblets As Collection
    Set dribblets = item("userAssetDribbletDetails")
    
    Dim i As Long
    i = (table.ListColumns.count - table.ListColumns(COL_DIFF).index) \ 4
    
    Do While i < dribblets.count
        i = i + 1

        table.ListColumns.Add().name = COL_COIN_PREFIX & i
        table.ListColumns.Add().name = COL_AMOUNT_PREFIX & i
        table.ListColumns.Add().name = COL_CHARGE_PREFIX & i
        table.ListColumns.Add().name = COL_VOLUME_PREFIX & i
    Loop
    
    Dim rowIndex As Long
    rowIndex = table.ListRows.Add().index

    table.ListColumns(COL_STAMP).DataBodyRange(rowIndex).Value = stamp
    table.ListColumns(COL_TRANSACTION_ID).DataBodyRange(rowIndex).Value = id
    
    For i = 1 To dribblets.count
        Dim dribblet As Dictionary
        Set dribblet = dribblets(i)
    
        table.ListColumns(COL_COIN_PREFIX & i).DataBodyRange(rowIndex).Value = dribblet("fromAsset")
        table.ListColumns(COL_AMOUNT_PREFIX & i).DataBodyRange(rowIndex).Value = dribblet("amount")
        table.ListColumns(COL_CHARGE_PREFIX & i).DataBodyRange(rowIndex).Value = dribblet("serviceChargeAmount")
        table.ListColumns(COL_VOLUME_PREFIX & i).DataBodyRange(rowIndex).Value = dribblet("transferedAmount")
    Next i
End Sub

Private Sub xCollectConversions(ByVal ops As BinanceOps, ByVal rowIndex As Long, ByVal checkApplience As Boolean)
    Dim table As ListObject
    Set table = xGetTable()
    
    If checkApplience <> table.ListColumns(COL_APPLIED).DataBodyRange(rowIndex).Value Then
        ops.unappliedOps = checkApplience And True
        Exit Sub
    End If

    Dim ticker As String
    ticker = table.ListColumns(COL_TICKER).DataBodyRange(rowIndex).Value
    
    If ticker = "" Then
        ops.unappliedOps = True
        Exit Sub
    End If
    
    Dim id As String
    id = Mid(table.ListColumns(COL_TRANSACTION_ID).DataBodyRange(rowIndex).Value, 4)
    
    Dim stamp As Date
    stamp = table.ListColumns(COL_STAMP).DataBodyRange(rowIndex).Value
    
    Dim coinIndex As Long
    coinIndex = 1
    
    Dim colIndex As Long
    colIndex = table.ListColumns(COL_DIFF).index + 1

    Do While colIndex < table.ListColumns.count
        Dim coin As String
        coin = table.ListColumns(colIndex).DataBodyRange(rowIndex).Value
        
        If coin = "" Then Exit Do
        
        If coin = ticker Then Err.Raise 1001, "TableDustLog.CollectConversions", "Input and output conversion asset are the same for transaction: " & ID_PREFIX & id
        
        Dim amount, charge, volume As Variant
        amount = table.ListColumns(colIndex + 1).DataBodyRange(rowIndex).Value
        charge = table.ListColumns(colIndex + 2).DataBodyRange(rowIndex).Value
        volume = table.ListColumns(colIndex + 3).DataBodyRange(rowIndex).Value
    
        colIndex = colIndex + 4
        
        Dim tid As String
        tid = Hex(159 + coinIndex) & "-" & id
    
        xAddConversions ops, stamp, tid, coin, amount, charge, volume, ticker, "BNB"
        xAddConversions ops, stamp, tid, coin, amount, charge, volume, ticker, "BTC"
        xAddConversions ops, stamp, tid, coin, amount, charge, volume, ticker, "ETH"
        xAddConversions ops, stamp, tid, coin, amount, charge, volume, ticker, "USDT"
        xAddConversions ops, stamp, tid, coin, amount, charge, volume, ticker, "USDC"
    
        coinIndex = coinIndex + 1
    Loop

    If Not checkApplience Then
        table.ListColumns(COL_APPLIED).DataBodyRange(rowIndex).Value = True
    End If
End Sub

Private Sub xAddConversions( _
    ByVal ops As BinanceOps, _
    ByVal stamp As Date, _
    ByVal tid As String, _
    ByVal coin As String, _
    ByVal amount As Variant, _
    ByVal charge As Variant, _
    ByVal volume As Variant, _
    ByVal ticker As String, _
    ByVal asset As String _
    )
    If ticker <> asset Then
        ops.AddSimpleOp asset, stamp, WALLET_SPOT, OP_CONVERT_IN, 0, 0, 0, "", tid
    Else
        Dim priceOut, priceIn As Variant
        volume = volume + charge
        priceOut = Round(Str2Dec(volume) / Str2Dec(amount), 8)
        priceIn = Round(Str2Dec(amount) / Str2Dec(volume), 8)
        ops.AddExchangeOp stamp, coin, ticker, WALLET_SPOT, WALLET_SPOT, OP_CONVERT_OUT, OP_CONVERT_IN, amount, volume, priceOut, priceIn, charge, 0, tid, "", "", "to:dust", "from:dust"
    End If
End Sub

