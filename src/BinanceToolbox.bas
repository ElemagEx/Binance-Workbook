Attribute VB_Name = "BinanceToolbox"
Private Const TEXT_BNB As String = "BNB"

' Prefixes A0..FF are reserved for dust log
Private Const TABLE_DRIBBLETS As String = "Dribblets"

Private Const COL_DRIBBLETS_COIN_PREFIX = "Coin "
Private Const COL_DRIBBLETS_AMOUNT_PREFIX = "Amount "
Private Const COL_DRIBBLETS_CHARGE_PREFIX = "Charge "
Private Const COL_DRIBBLETS_VOLUME_PREFIX = "Volume "

Private Const COL_DRIBBLETS_STAMP As String = "Stamp"
Private Const COL_DRIBBLETS_APPLIED As String = "Applied"
Private Const COL_DRIBBLETS_TX_ID As String = "Dribblet Tx ID"
Private Const COL_DRIBBLETS_TICKER As String = "Ticker"
Private Const COL_DRIBBLETS_COIN_1 As String = COL_DRIBBLETS_COIN_PREFIX & 1
Private Const COL_DRIBBLETS_AMOUNT_1 As String = COL_DRIBBLETS_AMOUNT_PREFIX & 1
Private Const COL_DRIBBLETS_CHARGE_1 As String = COL_DRIBBLETS_CHARGE_PREFIX & 1
Private Const COL_DRIBBLETS_VOLUME_1 As String = COL_DRIBBLETS_VOLUME_PREFIX & 1

Private Const COL_DRIBBLETS_DIFF = COL_DRIBBLETS_TICKER ' (columns.Count - COL_DRIBBLETS_DIFF) / 4 == <number-of-coins>
Private Const COL_DRIBBLETS_LAST = COL_DRIBBLETS_VOLUME_1

'
' Get current selection and try to find it in Assets/Ticker and Currencies/Ticker. If it found opens or creates sheet "<Ticker> Ledger"
'
Public Sub BinanceTool_Ledger_Open()
    Dim ticker As String
    ticker = ""
    
    If TypeName(Selection) = "Range" And Selection.Cells.count = 1 Then
        Dim selectedCell As range
        Set selectedCell = Selection
        
        If ActiveSheet.name = SHEET_ASSETS Then
            Dim assetsTable As ListObject
            Set assetsTable = ThisWorkbook.Sheets(SHEET_ASSETS).ListObjects(TABLE_ASSETS)
            
            Dim assetsTickerCol As ListColumn
            Set assetsTickerCol = assetsTable.ListColumns(COL_ASSETS_TICKER)
            
            If Not Intersect(selectedCell, assetsTickerCol.DataBodyRange) Is Nothing Then
                ticker = selectedCell.value
            End If
        End If
        If ActiveSheet.name = SHEET_CURRENCIES Then
            Dim currenciesTable As ListObject
            Set currenciesTable = ThisWorkbook.Sheets(SHEET_CURRENCIES).ListObjects(TABLE_CURRENCIES)
            
            Dim currenciesTickerCol As ListColumn
            Set currenciesTickerCol = currenciesTable.ListColumns(COL_CURRENCIES_TICKER)
            
            If Not Intersect(selectedCell, currenciesTickerCol.DataBodyRange) Is Nothing Then
                ticker = selectedCell.value
            End If
        End If
    End If
    
    If ticker = "" Then
        MsgBox "Please select single cell in Ticker column of Assets or Currencies Sheet"
        Exit Sub
    End If
    
    Dim ws As Worksheet
    Set ws = findWorksheet(ticker, True)
    
    ws.Activate
End Sub
'
' Reinitialize Ledger
'
Public Sub BinanceTool_Ledger_Reinitialize()
    Dim ticker As String
    Dim ws As Worksheet
    Set ws = getLedgerWorksheet(ticker)
    
    If ws Is Nothing Then
        MsgBox "Open Ledger worksheet first"
        Exit Sub
    End If
    
    Call resetAllUpdateTimes
    
    Call resetLedgerWorksheet(ws, ticker, True)
End Sub
'
' Refresh Ledger
'
Public Sub BinanceTool_Ledger_Refresh()
    Dim ticker As String
    Dim ws As Worksheet
    Set ws = getLedgerWorksheet(ticker)
    
    If ws Is Nothing Then
        MsgBox "Open Ledger worksheet first"
        Exit Sub
    End If
    
    Call resetLedgerWorksheet(ws, ticker, False)
End Sub
'
'
'
Public Sub BinanceTool_Ledger_UpdateTrades()
    Dim ticker As String
    Dim ws As Worksheet
    Set ws = getLedgerWorksheet(ticker)
    
    If ws Is Nothing Then
        MsgBox "Open Ledger worksheet first"
        Exit Sub
    End If
    
    Dim symbols As Dictionary
    Set symbols = collectTradeSymbols(ws, ticker)
    
    If symbols.count = 0 Then
        MsgBox "Not found symbols for " & ticker
        Exit Sub
    End If

    Dim lastUpdateTime, endTime, stamp As Date
    lastUpdateTime = getTradeLastUpdate(ws, ticker)
    
    Dim operation As Dictionary
    Dim Operations As Dictionary
    Set Operations = New Dictionary
    
    Dim quote, asset As String
    
    Dim symbol As Variant
    For Each symbol In symbols.Keys
        quote = symbols(symbol)
        Operations.Add quote, New collection
    Next symbol
    If ticker <> "BNB" And Not Operations.Exists("BNB") Then
        Operations.Add "BNB", New collection
    End If
    Operations.Add ticker, New collection
    
    For Each symbol In symbols.Keys
        quote = symbols(symbol)
        
        endTime = now()
        
        Do
            If endTime < lastUpdateTime Then
                Exit Do
            End If
            
            Do
                Dim list As collection
                Set list = BinanceApi_SpotTrading_GetMyTrades(symbol, 0, endTime, COLLECTION_LIMIT_TRADE)

                If list Is Nothing Then
                    Exit Sub
                End If
    
                Dim item As Dictionary
                For Each item In list
                    stamp = UnixTimestampToDate(item("time"))

                    If endTime > stamp Then
                        endTime = stamp
                    End If
                    If endTime < lastUpdateTime Then
                        Exit For
                    End If
                    
                    Dim isBuyer, isMaker, isFeeBNB As Boolean
                    isBuyer = item("isBuyer")
                    isMaker = item("isMaker")
                    
                    asset = IIf(item.Exists("commissionAsset"), item("commissionAsset"), "")
        
                    If Operations.Exists(asset) Then
                        isFeeBNB = (asset = "BNB")
                    Else
                        MsgBox "Unknown Commission Asset: " & asset
                        Exit For
                    End If
                    
                    Set operation = New Dictionary
                    
                    operation(COL_DETAILS_STAMP) = stamp
                    operation(COL_DETAILS_WALLET) = WALLET_SPOT
                    operation(COL_DETAILS_OPERATION) = IIf(isBuyer, OPERATION_BUY, OPERATION_SELL)
                    operation(COL_DETAILS_ACQUIRED) = IIf(isBuyer, item("qty"), 0)
                    operation(COL_DETAILS_SPENT) = IIf(isBuyer, 0, item("qty"))
                    operation(COL_DETAILS_TICKER) = quote
                    operation(COL_DETAILS_PRICE) = item("price")
                    operation(COL_DETAILS_AMOUNT) = item("quoteQty")
                    operation(COL_DETAILS_CHARGE) = IIf(isFeeBNB, 0, item("commission"))
                    operation(COL_DETAILS_BNB_FEE) = IIf(isFeeBNB, item("commission"), 0)
                    operation(COL_DETAILS_NOTE) = IIf(isMaker, "MAKER", "TAKER")
                    operation(COL_DETAILS_ID) = ID_PREFIX_TRADE & item("id")
                    
                    Operations(ticker).Add operation
                    
                    Set operation = New Dictionary
                    
                    operation(COL_DETAILS_STAMP) = stamp
                    operation(COL_DETAILS_WALLET) = WALLET_SPOT
                    operation(COL_DETAILS_OPERATION) = IIf(isBuyer, OPERATION_EXPENCE, OPERATION_INCOME)
                    operation(COL_DETAILS_ACQUIRED) = IIf(isBuyer, 0, item("quoteQty"))
                    operation(COL_DETAILS_SPENT) = IIf(isBuyer, item("quoteQty"), 0)
                    operation(COL_DETAILS_TICKER) = ticker
                    operation(COL_DETAILS_PRICE) = Round(Str2Dec(item("qty")) / Str2Dec(item("quoteQty")), 8)
                    operation(COL_DETAILS_AMOUNT) = item("qty")
                    operation(COL_DETAILS_CHARGE) = IIf(isFeeBNB, 0, item("commission"))
                    operation(COL_DETAILS_BNB_FEE) = IIf(isFeeBNB, item("commission"), 0)
                    operation(COL_DETAILS_NOTE) = IIf(isMaker, "MAKER", "TAKER")
                    operation(COL_DETAILS_ID) = ID_PREFIX_TRADE & item("id")
                    
                    Operations(quote).Add operation
                    
                    If isFeeBNB Then
                        Set operation = New Dictionary
                        
                        operation(COL_DETAILS_STAMP) = stamp
                        operation(COL_DETAILS_WALLET) = WALLET_SPOT
                        operation(COL_DETAILS_OPERATION) = OPERATION_COMMISSION
                        operation(COL_DETAILS_ACQUIRED) = 0
                        operation(COL_DETAILS_SPENT) = item("commission")
                        operation(COL_DETAILS_TICKER) = ""
                        operation(COL_DETAILS_PRICE) = 0
                        operation(COL_DETAILS_AMOUNT) = 0
                        operation(COL_DETAILS_CHARGE) = 0
                        operation(COL_DETAILS_BNB_FEE) = 0
                        operation(COL_DETAILS_NOTE) = ticker & "/" & quote
                        operation(COL_DETAILS_ID) = ID_PREFIX_TRADE & item("id")
                        
                        Operations("BNB").Add operation
                    End If
                Next item
                
                If list.count < COLLECTION_LIMIT_TRADE Then
                    endTime = DateAdd("s", -1, lastUpdateTime)
                    Exit Do
                Else
                    endTime = DateAdd("s", 1, endTime)
                End If
            Loop
        Loop
    Next
    
    Call BinanceTool_AddOperations(Operations)
    
    Call setLastUpdate(COL_LAST_UPDATE_DISTRIBUTION)
End Sub
'
'
'
Public Sub BinanceTool_Assets_UpdateDribblets()
    Dim hasSkippedOperations As Boolean
    hasSkippedOperations = False
    
    Dim operation As Dictionary
    Dim Operations As Dictionary
    Set Operations = New Dictionary
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim filter, ticker As String
    filter = IIf(isThisLedgerWorksheet(ws, ticker), ticker, "")

    Dim lastUpdateTime, startTime, endTime, stamp As Date

    lastUpdateTime = IIf(filter = "", getLastUpdate(COL_LAST_UPDATE_DRIBBLETS), CDate("2025-06-01"))

    endTime = now()
    
    Do
        If endTime < lastUpdateTime Then
            Exit Do
        End If
        
        startTime = DateAdd("d", -180, endTime)
        
        If startTime < lastUpdateTime Then
            startTime = lastUpdateTime
        End If
        
        Do
            Dim data As Dictionary
            Set data = BinanceApi_Wallet_GetDustLog(startTime, endTime)
        
            If data Is Nothing Then
                Exit Sub
            End If
            If Not data.Exists("userAssetDribblets") Then
                endTime = DateAdd("s", -1, startTime)
                Exit Do
            End If
            
            Dim list As collection
            Set list = data("userAssetDribblets")

            Dim item As Dictionary
            For Each item In list
                stamp = UnixTimestampToDate(item("operateTime"))

                If endTime > stamp Then
                    endTime = stamp
                End If

                If Not tryCollectDribblet(Operations, item, stamp) Then
                    hasSkippedOperations = True
                End If
            Next item
            
            If list.count < 100 Then
                endTime = DateAdd("s", -1, startTime)
                Exit Do
            Else
                endTime = DateAdd("s", -1, endTime)
            End If
        Loop
    Loop
    
    If filter <> "" Then
        If Not Operations.Exists(filter) Then
            Set Operations = New Dictionary
        Else
            Dim ops As collection
            Set ops = Operations(filter)
            Set Operations = New Dictionary
            Operations.Add filter, ops
        End If
    End If
    
    Call BinanceTool_AddOperations(Operations, hasSkippedOperations)
    
    If filter = "" Then
        Call setLastUpdate(COL_LAST_UPDATE_DRIBBLETS)
    End If
End Sub
Private Function Str2Fmt(ByVal str As String) As String
    Dim val As String
    val = Replace(str, ",", Application.ThousandsSeparator)
    val = Replace(str, ".", Application.DecimalSeparator)
    Str2Fmt = val
End Function
' Returns wallet name by type
Private Function getWalletByType(ByVal walletType As Long)

    Select Case walletType
        Case Is = 0
            getWalletByType = WALLET_SPOT
        Case Is = 1
            getWalletByType = WALLET_FUNDING
        Case Else
            getWalletByType = "<UNKNOWN>"
    End Select

End Function
' Returns worksheet by ticker if exists or Nothing otherwise
Private Function findWorksheet(ByVal ticker As String, Optional createIfDoesNotExists As Boolean = False) As Worksheet
    Set findWorksheet = Nothing
    
    Dim name As String
    name = ticker & " " & SHEET_LEDGER_SUFFIX
    
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If name = ws.name Then
            Set findWorksheet = ws
            Exit Function
        End If
    Next ws

    If createIfDoesNotExists Then
        Set findWorksheet = createNewLedgerWorksheet(ticker, name)
    End If
End Function

Private Sub delLedgerDetailsRecord(ByVal ws As Worksheet, ByVal record As Dictionary)
    Dim ticker As String

    If Not isThisLedgerWorksheet(ws, ticker) Then
        Exit Sub
    End If

    Dim table As ListObject
    Set table = ws.ListObjects(ticker + "_" + TABLE_DETAILS)
    
    If table.ListRows.count > 0 Then
        Dim cell As range
        Set cell = table.ListColumns(COL_DETAILS_ID).DataBodyRange.Find(what:=record(COL_DETAILS_ID), LookIn:=xlValues, LookAt:=xlWhole)
    
        If Not cell Is Nothing Then
            table.ListRows(cell.row - table.HeaderRowRange.row).Delete
        End If
    End If
End Sub
Private Function getLedgerWorksheet(ByRef ticker As String, Optional ByVal ws As Worksheet = Nothing) As Worksheet
    Set getLedgerWorksheet = Nothing
    
    If ws Is Nothing Then
        Set ws = ActiveSheet
    End If
    
    Dim parts() As String
    parts = Split(ws.name, " ")
    
    Dim count As Long
    count = UBound(parts) - LBound(parts) + 1

    If count <> 2 Or parts(UBound(parts)) <> SHEET_LEDGER_SUFFIX Then
        Exit Function
    End If

    ticker = parts(LBound(parts))

    If ws.ListObjects(ticker + "_" + TABLE_INFO) Is Nothing Then
        Exit Function
    End If
    If ws.ListObjects(ticker + "_" + TABLE_SUMMARY) Is Nothing Then
        Exit Function
    End If
    If ws.ListObjects(ticker + "_" + TABLE_TRADES) Is Nothing Then
        Exit Function
    End If
    If ws.ListObjects(ticker + "_" + TABLE_DETAILS) Is Nothing Then
        Exit Function
    End If

    Set getLedgerWorksheet = ws
End Function

Private Function isThisLedgerWorksheet(ByVal ws As Worksheet, ByRef ticker As String) As Boolean
    Set ws = getLedgerWorksheet(ticker, ws)

    isThisLedgerWorksheet = Not ws Is Nothing
End Function

Private Function collectTradeSymbols(ByVal ws As Worksheet, ByVal ticker As String) As Dictionary
    Set collectTradeSymbols = New Dictionary
    
    Dim table As ListObject
    Set table = ws.ListObjects(ticker + "_" + TABLE_TRADES)
    
    Dim base, quote, symbol As String
    base = table.ListColumns(COL_TRADES_TRADES).DataBodyRange(1).value
    
    Dim i As Long
    For i = 2 To table.ListColumns.count - 1
        If table.ListColumns(i).DataBodyRange(1).value = True Then
            quote = table.ListColumns(i).name
            
            symbol = base & quote
            
            collectTradeSymbols.Add symbol, quote
        End If
    Next i
End Function

Private Function getTradeLastUpdate(ByVal ws As Worksheet, ByVal ticker As String) As Date

    Dim table As ListObject
    Set table = ws.ListObjects(ticker + "_" + TABLE_INFO)
    
    getTradeLastUpdate = CDate(table.ListColumns(COL_INFO_VALUE).DataBodyRange(ROW_INFO_ACTIVITY_START_DATE))
    
End Function

Private Function getLastUpdate(ByVal colName As String) As Date
    Dim table As ListObject
    Set table = ThisWorkbook.Sheets(SHEET_ASSETS).ListObjects(TABLE_LAST_UPDATE)
   
    Dim cell As range
    Set cell = table.ListColumns(colName).DataBodyRange(1)
    
    If IsDate(cell.value) Then
        getLastUpdate = cell.value
    
        If updateToNow Then
            cell.value = now()
        End If
    Else
        If updateToNow Then
            cell.value = now()
        End If
        
        Set table = ThisWorkbook.Sheets(SHEET_ASSETS).ListObjects(TABLE_ACCOUNT)
        
        Set cell = table.ListColumns(COL_ACCOUNT_ACTIVITY_START_DATE).DataBodyRange(1)
    
        getLastUpdate = cell.value
    End If
End Function

Private Function setLastUpdate(ByVal colName As String, Optional ByVal newTime As Date = 0) As Date
    Dim table As ListObject
    Set table = ThisWorkbook.Sheets(SHEET_ASSETS).ListObjects(TABLE_LAST_UPDATE)
   
    Dim cell As range
    Set cell = table.ListColumns(colName).DataBodyRange(1)
    
    setLastUpdate = IIf(newTime <> 0, newTime, now())
    
    cell.value = setLastUpdate
End Function

Private Function BinanceFunc_Currencies_GetTickerQuotes(ByVal ticker As String) As collection
    Dim quotes As collection
    Set quotes = New collection
    
    Set BinanceFunc_Currencies_GetTickerQuotes = quotes
    
    Dim exchangeInfo As Dictionary
    Set exchangeInfo = BinanceApi_SpotTrading_exchangeInfo(False)

    Dim symbols As collection
    Set symbols = exchangeInfo("symbols")

    Dim symbol As Dictionary
    For Each symbol In symbols
        If ticker = symbol("baseAsset") Then
            quotes.Add symbol("quoteAsset")
        End If
    Next symbol
End Function

Private Function collectDribbletOperations(ByVal Operations As Dictionary, Optional ByVal transId As String = "") As Boolean

    collectDribbletOperations = False

    Dim table As ListObject
    Set table = ThisWorkbook.Sheets(SHEET_MISC).ListObjects(TABLE_DRIBBLETS)
    
    Dim col As ListColumn
    Set col = table.ListColumns(COL_DRIBBLETS_TX_ID)
    
    If col.DataBodyRange Is Nothing Then
        Exit Function
    End If
    
    Dim ids As collection
    Set ids = New collection
    
    Dim cell As range
    
    If transId <> "" Then
        ids.Add ID_PREFIX_DRIBBLET & transId
    Else
        For Each cell In col.DataBodyRange
            ids.Add cell.value
        Next cell
    End If

    Dim id As Variant
    For Each id In ids
        Set cell = col.DataBodyRange.Find(what:=id, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    
        If cell Is Nothing Then
            Exit Function
        End If
        
        Dim rowNum As Long
        rowNum = cell.row - table.HeaderRowRange.row
    
        If Not IsEmpty(table.ListColumns(COL_DRIBBLETS_TICKER).DataBodyRange(rowNum)) Then
            If transId = "" And table.ListColumns(COL_DRIBBLETS_APPLIED).DataBodyRange(rowNum).value = True Then
                Exit Function
            End If
            
            collectDribbletOperations = True
            
            table.ListColumns(COL_DRIBBLETS_APPLIED).DataBodyRange(rowNum).value = True
            
            Dim tid As String
            tid = Mid(table.ListColumns(COL_DRIBBLETS_TX_ID).DataBodyRange(rowNum).value, 4)
            
            Dim stamp As Date
            stamp = table.ListColumns(COL_DRIBBLETS_STAMP).DataBodyRange(rowNum)
            
            Dim ticker As String
            ticker = table.ListColumns(COL_DRIBBLETS_TICKER).DataBodyRange(rowNum)
            
            If Not Operations.Exists(ticker) Then
                Operations.Add ticker, New collection
            End If
            
            Dim num_coins As Long
            num_coins = (table.ListColumns.count - table.ListColumns(COL_DRIBBLETS_DIFF).Index) \ 4
            
            Dim i As Long
            For i = 1 To num_coins
                Dim asset, amount, charge, volume As Variant
                asset = table.ListColumns(COL_DRIBBLETS_COIN_PREFIX & i).DataBodyRange(rowNum).value
                
                If IsEmpty(asset) Then
                    Exit For
                End If
                
                If Not Operations.Exists(asset) Then
                    Operations.Add asset, New collection
                End If

                amount = CDec(table.ListColumns(COL_DRIBBLETS_AMOUNT_PREFIX & i).DataBodyRange(rowNum).value)
                charge = CDec(table.ListColumns(COL_DRIBBLETS_CHARGE_PREFIX & i).DataBodyRange(rowNum).value)
                volume = CDec(table.ListColumns(COL_DRIBBLETS_VOLUME_PREFIX & i).DataBodyRange(rowNum).value)
                
                Dim operation As Dictionary
                Set operation = New Dictionary
                
                operation(COL_DETAILS_STAMP) = stamp
                operation(COL_DETAILS_WALLET) = WALLET_SPOT
                operation(COL_DETAILS_OPERATION) = OPERATION_DUST_OUT
                operation(COL_DETAILS_ACQUIRED) = 0
                operation(COL_DETAILS_SPENT) = amount
                operation(COL_DETAILS_TICKER) = ticker
                operation(COL_DETAILS_PRICE) = Round(volume / amount, 8)
                operation(COL_DETAILS_AMOUNT) = volume
                operation(COL_DETAILS_CHARGE) = charge
                operation(COL_DETAILS_BNB_FEE) = 0
                operation(COL_DETAILS_NOTE) = ""
                operation(COL_DETAILS_ID) = Hex(159 + i) & "-" & tid
                
                Operations(asset).Add operation
                
                Set operation = New Dictionary
                
                operation(COL_DETAILS_STAMP) = stamp
                operation(COL_DETAILS_WALLET) = WALLET_SPOT
                operation(COL_DETAILS_OPERATION) = OPERATION_DUST_IN
                operation(COL_DETAILS_ACQUIRED) = volume
                operation(COL_DETAILS_SPENT) = 0
                operation(COL_DETAILS_TICKER) = asset
                operation(COL_DETAILS_PRICE) = Round(amount / volume, 8)
                operation(COL_DETAILS_AMOUNT) = amount
                operation(COL_DETAILS_CHARGE) = charge
                operation(COL_DETAILS_BNB_FEE) = 0
                operation(COL_DETAILS_NOTE) = ""
                operation(COL_DETAILS_ID) = Hex(159 + i) & "-" & tid
                
                Operations(ticker).Add operation
            Next i
        End If
    Next id
End Function

Private Function tryCollectDribblet(ByVal Operations As Dictionary, ByVal item As Dictionary, ByVal stamp As Date) As Boolean

    tryCollectDribblet = collectDribbletOperations(Operations, item("transId"))

    If tryCollectDribblet Then
        Exit Function
    End If

    Dim entry As Dictionary
    Set entry = New Dictionary

    entry(COL_DRIBBLETS_STAMP) = stamp
    entry(COL_DRIBBLETS_APPLIED) = False
    entry(COL_DRIBBLETS_TX_ID) = ID_PREFIX_DRIBBLET & item("transId")

    Dim assets As collection
    Set assets = item("userAssetDribbletDetails")
    
    Dim i As Long
    For i = 1 To assets.count
        Dim asset As Dictionary
        Set asset = assets(i)
        
        entry(COL_DRIBBLETS_COIN_PREFIX & i) = asset("fromAsset")
        entry(COL_DRIBBLETS_AMOUNT_PREFIX & i) = asset("amount")
        entry(COL_DRIBBLETS_CHARGE_PREFIX & i) = asset("serviceChargeAmount")
        entry(COL_DRIBBLETS_VOLUME_PREFIX & i) = asset("transferedAmount")
    Next i

    Dim table As ListObject
    Set table = ThisWorkbook.Sheets(SHEET_MISC).ListObjects(TABLE_DRIBBLETS)
    
    Dim num_coins As Long
    num_coins = (table.ListColumns.count - table.ListColumns(COL_DRIBBLETS_DIFF).Index) \ 4

    Do While num_coins < assets.count
        num_coins = num_coins + 1

        table.ListColumns.Add().name = COL_DRIBBLETS_COIN_PREFIX & num_coins
        table.ListColumns.Add().name = COL_DRIBBLETS_AMOUNT_PREFIX & num_coins
        table.ListColumns.Add().name = COL_DRIBBLETS_CHARGE_PREFIX & num_coins
        table.ListColumns.Add().name = COL_DRIBBLETS_VOLUME_PREFIX & num_coins
    Loop

    table.ListRows.Add

    Dim rowNum As Long
    rowNum = table.ListRows.count

    Dim key As Variant
    For Each key In entry.Keys
        table.ListColumns(key).DataBodyRange(rowNum).value = entry(key)
    Next key
End Function
