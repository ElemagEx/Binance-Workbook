Private Const TEXT_BNB As String = "BNB"
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
