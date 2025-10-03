VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BinanceOps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_ops As New Dictionary
Private m_endDate As Date
Public unappliedOps

Public Property Get endDate() As Date
    endDate = IIf(m_endDate = 0, Now(), m_endDate)
End Property
Public Property Let endDate(ByVal newEndDate As Date)
    m_endDate = newEndDate
End Property

Public Property Get Tickers() As Variant
    Tickers = m_ops.Keys
End Property

Public Property Get item(ByVal ticker As String) As Collection
    Set item = IIf(m_ops.Exists(ticker), m_ops(ticker), Nothing)
End Property

Public Function Contains(ByVal ticker As String) As Boolean
    Contains = m_coins.Exists(ticker)
End Function

Public Function HasOnlyDelOps(ByVal ticker As String) As Boolean
    HasOnlyDelOps = False
    
    If Not m_ops.Exists(ticker) Then
        Exit Function
    End If
    Dim op As BinanceOp
    For Each op In m_ops(ticker)
        If op.acquired > 0 Or op.spent > 0 Then Exit Function
    Next op

    HasOnlyDelOps = True
End Function

Public Sub Collect_Transfers(ByVal lastUpdateTime As Date)
    Dim transferTypes As New Collection
    transferTypes.Add "MAIN_FUNDING"
    transferTypes.Add "FUNDING_MAIN"
    
    Dim sourceWallet As String
    Dim targetWallet As String
    
    Dim transferType As Variant
    For Each transferType In transferTypes
        
        If Left(transferType, 5) = "MAIN_" Then
            sourceWallet = WALLET_SPOT
        ElseIf Left(transferType, 8) = "FUNDING_" Then
            sourceWallet = WALLET_FUNDING
        Else
            Assert_Fail
        End If
        If Right(transferType, 5) = "_MAIN" Then
            targetWallet = WALLET_SPOT
        ElseIf Right(transferType, 8) = "_FUNDING" Then
            targetWallet = WALLET_FUNDING
        Else
            Assert_Fail
        End If
        
        Dim startTime, endTime, stamp As Date
        
        endTime = Me.endDate
        
        Do While endTime >= lastUpdateTime
            
            startTime = DateAdd("d", -MAX_PERIOD.WALLET_TRANSFER_HISTORY, endTime)

            If startTime < lastUpdateTime Then startTime = lastUpdateTime

            Do
                Dim data As Dictionary
                Set data = BinanceAPI.Wallet_GetTransferHistory(transferType, startTime, endTime, MAX_LIMIT.WALLET_TRANSFER_HISTORY)
                
                If Not data.Exists("rows") Then
                    endTime = DateAdd("s", -1, startTime)
                    Exit Do
                End If
                
                Dim list As Collection
                Set list = data("rows")
    
                Dim item As Dictionary
                For Each item In list
                    stamp = xUnixTimestampToDate(item("timestamp"))
                    
                    If endTime > stamp Then endTime = stamp
                
                    If item("status") = "CONFIRMED" Then
                        Dim amount As Variant
                        Dim ticker, tranId As String
                        ticker = item("asset")
                        amount = item("amount")
                        tranId = item("tranId")
                        
                        AddSimpleOp ticker, stamp, sourceWallet, OP_WALLET_OUT, 0, amount, 0, "to:" & targetWallet, ID_PREFIX_WALLET_OUT & tranId
                        AddSimpleOp ticker, stamp, targetWallet, OP_WALLET_IN, amount, 0, 0, "from:" & sourceWallet, ID_PREFIX_WALLET_IN & tranId
                    End If
                Next item
                
                If list.count < MAX_LIMIT.WALLET_TRANSFER_HISTORY Then
                    endTime = DateAdd("s", -1, startTime)
                    Exit Do
                Else
                    endTime = DateAdd("s", -1, endTime)
                End If
            Loop
        Loop
    Next transferType
End Sub

Public Sub Collect_Conversions(ByVal lastUpdateTime As Date)
    Dim startTime, endTime, stamp As Date
    
    endTime = Me.endDate
    
    Do While endTime >= lastUpdateTime
        
        startTime = DateAdd("d", -MAX_PERIOD.CONVERT_TRADE_HISTORY, endTime)
        
        If startTime < lastUpdateTime Then startTime = lastUpdateTime
        
        Do
            Dim data As Dictionary
            Set data = BinanceAPI.Convert_GetConvertHistory(startTime, endTime, MAX_LIMIT.CONVERT_TRADE_HISTORY)
            
            If Not data.Exists("list") Then
                endTime = DateAdd("s", -1, startTime)
                Exit Do
            End If
            
            Dim list As Collection
            Set list = data("list")
    
            Dim item As Dictionary
            For Each item In list
                stamp = xUnixTimestampToDate(item("createTime"))
                
                If endTime > stamp Then endTime = stamp
                
                If item("orderStatus") = "SUCCESS" Then
                    Select Case item("walletType")
                        Case "SPOT"
                            xHandleSimpleConversion item, stamp, WALLET_SPOT
                        Case "FUNDING"
                            xHandleSimpleConversion item, stamp, WALLET_FUNDING
                        Case "EARN"
                            xHandleSimpleConversion item, stamp, WALLET_FLEXIBLE_EARN
                        Case "SPOT_FUNDING"
                            xHandleCombosConversion item, stamp, True, True, False
                            xHandleSimpleConversion item, stamp, WALLET_SPOT
                        Case "SPOT_EARN"
                            xHandleCombosConversion item, stamp, True, False, True
                            xHandleSimpleConversion item, stamp, WALLET_SPOT
                        Case "FUNDING_EARN"
                            xHandleCombosConversion item, stamp, False, True, True
                            xHandleSimpleConversion item, stamp, WALLET_FUNDING
                        Case "SPOT_FUNDING_EARN"
                            xHandleCombosConversion item, stamp, True, True, True
                            xHandleSimpleConversion item, stamp, WALLET_SPOT
                    End Select
                End If
            Next item
            
            If list.count < MAX_LIMIT.CONVERT_TRADE_HISTORY Then
                endTime = DateAdd("s", -1, startTime)
                Exit Do
            Else
                endTime = DateAdd("s", -1, endTime)
            End If
        Loop
    Loop
End Sub

Public Sub Collect_DustLog(ByVal lastUpdateTime As Date)
    Dim startTime, endTime, stamp As Date
    
    endTime = Me.endDate
    
    Do While endTime >= lastUpdateTime
        
        startTime = DateAdd("d", -MAX_PERIOD.WALLET_DRIBLETS_HISTORY, endTime)
        
        If startTime < lastUpdateTime Then startTime = lastUpdateTime
        
        Do
            Dim data As Dictionary
            Set data = BinanceAPI.Wallet_GetDustLog(startTime, endTime)
            
            If Not data.Exists("userAssetDribblets") Then
                endTime = DateAdd("s", -1, startTime)
                Exit Do
            End If
            
            Dim list As Collection
            Set list = data("userAssetDribblets")
    
            Dim item As Dictionary
            For Each item In list
                stamp = xUnixTimestampToDate(item("operateTime"))
                
                If endTime > stamp Then endTime = stamp
                
                TableDustLog.HandleTransaction Me, item, stamp
            Next item
            
            If list.count < MAX_LIMIT.WALLET_DRIBLETS_HISTORY Then
                endTime = DateAdd("s", -1, startTime)
                Exit Do
            Else
                endTime = DateAdd("s", -1, endTime)
            End If
        Loop
    Loop
End Sub

Public Sub Collect_FiatBuys(ByVal lastUpdateTime As Date)
    Dim startTime, endTime, stamp As Date
    
    endTime = Me.endDate
    
    Do While endTime >= lastUpdateTime
        
        startTime = DateAdd("d", -MAX_PERIOD.FIAT_OPERATIONS, endTime)
        
        If startTime < lastUpdateTime Then startTime = lastUpdateTime
        
        Do
            Dim data As Dictionary
            Set data = BinanceAPI.Fiat_GetBuys(startTime, endTime, MAX_LIMIT.FIAT_OPERATIONS)
            
            If Not data.Exists("data") Then
                endTime = DateAdd("s", -1, startTime)
                Exit Do
            End If
            
            Dim list As Collection
            Set list = data("data")
    
            Dim item As Dictionary
            For Each item In list
                stamp = xUnixTimestampToDate(item("updateTime"))
                
                If endTime > stamp Then endTime = stamp
                
                If item("status") = "Completed" Then
                    Dim amount, charge As Variant
                    Dim ticker, tranId As String
                    ticker = item("fiatCurrency")
                    tranId = item("orderNo")
                    amount = Str2Dec(item("sourceAmount"))
                    charge = Str2Dec(item("totalFee"))
                    
                    AddSimpleOp ticker, stamp, WALLET_SPOT, OP_FIAT_BUY, amount, 0, charge, STR_NA, ID_PREFIX_FIAT_BUY & tranId

                    Dim op As BinanceOp
                    Set op = New BinanceOp
                    
                    op.stamp = stamp
                    op.wallet = WALLET_SPOT
                    op.operation = OP_CONVERT_OUT
                    op.acquired = 0
                    op.spent = Str2Dec(item("sourceAmount")) - Str2Dec(item("totalFee"))
                    op.ticker = item("cryptoCurrency")
                    op.price = Round(Str2Dec(item("obtainAmount")) / (Str2Dec(item("sourceAmount")) - Str2Dec(item("totalFee"))), 8)
                    op.amount = item("obtainAmount")
                    op.charge = 0
                    op.bnb_fee = 0
                    op.note = "FIAT:BUY"
                    op.id = ID_PREFIX_CONVERT_OUT & tranId
                    
                    xAddOp ticker, op
                
                    ticker = item("cryptoCurrency")
                    
                    Set op = New BinanceOp

                    op.stamp = stamp
                    op.wallet = WALLET_SPOT
                    op.operation = OP_CONVERT_IN
                    op.acquired = item("obtainAmount")
                    op.spent = 0
                    op.ticker = item("fiatCurrency")
                    op.price = item("price")
                    op.amount = Str2Dec(item("sourceAmount")) - Str2Dec(item("totalFee"))
                    op.charge = 0
                    op.bnb_fee = 0
                    op.note = "FIAT:BUY"
                    op.id = ID_PREFIX_CONVERT_IN & tranId
                
                    xAddOp ticker, op
                End If
            Next item
            
            If list.count < MAX_LIMIT.FIAT_OPERATIONS Then
                endTime = DateAdd("s", -1, startTime)
                Exit Do
            Else
                endTime = DateAdd("s", -1, endTime)
            End If
        Loop
    Loop
End Sub

Public Sub Collect_FiatSells(ByVal lastUpdateTime As Date)
    Dim startTime, endTime, stamp As Date
    
    endTime = Me.endDate
    
    Do While endTime >= lastUpdateTime
        
        startTime = DateAdd("d", -MAX_PERIOD.FIAT_OPERATIONS, endTime)
        
        If startTime < lastUpdateTime Then startTime = lastUpdateTime
        
        Do
            Dim data As Dictionary
            Set data = BinanceAPI.Fiat_GetSells(startTime, endTime, MAX_LIMIT.FIAT_OPERATIONS)
            
            If Not data.Exists("data") Then
                endTime = DateAdd("s", -1, startTime)
                Exit Do
            End If
            
            Dim list As Collection
            Set list = data("data")
    
            Dim item As Dictionary
            For Each item In list
                stamp = xUnixTimestampToDate(item("updateTime"))
                
                If endTime > stamp Then endTime = stamp
                
                If item("status") = "Completed" Then
                    Dim amount, charge As Variant
                    Dim ticker, tranId As String
                    
                    ticker = item("cryptoCurrency")
                    
                    Dim op As BinanceOp
                    Set op = New BinanceOp

                    op.stamp = DateAdd("s", -1, stamp)
                    op.wallet = WALLET_SPOT
                    op.operation = OP_CONVERT_OUT
                    op.acquired = 0
                    op.spent = item("obtainAmount")
                    op.ticker = item("fiatCurrency")
                    op.price = item("price")
                    op.amount = Str2Dec(item("sourceAmount")) + Str2Dec(item("totalFee"))
                    op.charge = 0
                    op.bnb_fee = 0
                    op.note = "FIAT:SELL"
                    op.id = ID_PREFIX_CONVERT_OUT & item("orderNo")
                
                    xAddOp ticker, op
                    
                    ticker = item("fiatCurrency")
                    
                    Set op = New BinanceOp
                    
                    op.stamp = DateAdd("s", -1, stamp)
                    op.wallet = WALLET_SPOT
                    op.operation = OP_CONVERT_IN
                    op.acquired = Str2Dec(item("sourceAmount")) + Str2Dec(item("totalFee"))
                    op.spent = 0
                    op.ticker = item("cryptoCurrency")
                    op.price = Round(Str2Dec(item("obtainAmount")) / (Str2Dec(item("sourceAmount")) + Str2Dec(item("totalFee"))), 8)
                    op.amount = item("obtainAmount")
                    op.charge = 0
                    op.bnb_fee = 0
                    op.note = "FIAT:SELL"
                    op.id = ID_PREFIX_CONVERT_IN & item("orderNo")
                    
                    xAddOp ticker, op
                    
                    ticker = item("fiatCurrency")
                    tranId = item("orderNo")
                    amount = Str2Dec(item("sourceAmount"))
                    charge = Str2Dec(item("totalFee"))
                    
                    AddSimpleOp ticker, stamp, WALLET_SPOT, OP_FIAT_SELL, 0, amount, charge, STR_NA, ID_PREFIX_FIAT_SELL & tranId
                End If
            Next item
            
            If list.count < MAX_LIMIT.FIAT_OPERATIONS Then
                endTime = DateAdd("s", -1, startTime)
                Exit Do
            Else
                endTime = DateAdd("s", -1, endTime)
            End If
        Loop
    Loop
End Sub

Public Sub Collect_FiatDeposits(ByVal lastUpdateTime As Date)
    Dim startTime, endTime, stamp As Date
    
    endTime = Me.endDate
    
    Do While endTime >= lastUpdateTime
        
        startTime = DateAdd("d", -MAX_PERIOD.FIAT_OPERATIONS, endTime)
        
        If startTime < lastUpdateTime Then startTime = lastUpdateTime
        
        Do
            Dim data As Dictionary
            Set data = BinanceAPI.Fiat_GetDeposits(startTime, endTime, MAX_LIMIT.FIAT_OPERATIONS)
            
            If Not data.Exists("data") Then
                endTime = DateAdd("s", -1, startTime)
                Exit Do
            End If
            
            Dim list As Collection
            Set list = data("data")
    
            Dim item As Dictionary
            For Each item In list
                stamp = xUnixTimestampToDate(item("updateTime"))
                
                If endTime > stamp Then endTime = stamp
                
                If item("status") = "Successful" Then
                    Dim amount, charge As Variant
                    Dim ticker, tranId As String
                    ticker = item("fiatCurrency")
                    tranId = item("orderNo")
                    amount = Str2Dec(item("indicatedAmount"))
                    charge = Str2Dec(item("totalFee"))
                    
                    AddSimpleOp ticker, stamp, WALLET_SPOT, OP_DEPOSIT, amount, 0, charge, item("method"), ID_PREFIX_FIAT_DEPOSIT & tranId
                End If
            Next item
            
            If list.count < MAX_LIMIT.FIAT_OPERATIONS Then
                endTime = DateAdd("s", -1, startTime)
                Exit Do
            Else
                endTime = DateAdd("s", -1, endTime)
            End If
        Loop
    Loop
End Sub

Public Sub Collect_FiatWithdraws(ByVal lastUpdateTime As Date)
    Dim startTime, endTime, stamp As Date
    
    endTime = Me.endDate
    
    Do While endTime >= lastUpdateTime
        
        startTime = DateAdd("d", -MAX_PERIOD.FIAT_OPERATIONS, endTime)
        
        If startTime < lastUpdateTime Then startTime = lastUpdateTime
        
        Do
            Dim data As Dictionary
            Set data = BinanceAPI.Fiat_GetWithdraws(startTime, endTime, MAX_LIMIT.FIAT_OPERATIONS)
            
            If Not data.Exists("data") Then
                endTime = DateAdd("s", -1, startTime)
                Exit Do
            End If
            
            Dim list As Collection
            Set list = data("data")
    
            Dim item As Dictionary
            For Each item In list
                stamp = xUnixTimestampToDate(item("updateTime"))
                
                If endTime > stamp Then endTime = stamp
                
                If item("status") = "Successful" Then
                    Dim amount, charge As Variant
                    Dim ticker, tranId As String
                    ticker = item("fiatCurrency")
                    tranId = item("orderNo")
                    amount = Str2Dec(item("amount"))
                    charge = Str2Dec(item("totalFee"))
                    
                    AddSimpleOp ticker, stamp, WALLET_SPOT, OP_WITHDRAW, 0, amount, charge, item("method"), ID_PREFIX_FIAT_WITHDRAW & tranId
                End If
            Next item
            
            If list.count < MAX_LIMIT.FIAT_OPERATIONS Then
                endTime = DateAdd("s", -1, startTime)
                Exit Do
            Else
                endTime = DateAdd("s", -1, endTime)
            End If
        Loop
    Loop
End Sub

Public Sub Collect_CryptoDeposits(ByVal lastUpdateTime As Date, Optional ByVal filter As String = "")
    Dim startTime, endTime, stamp As Date
    
    endTime = Me.endDate
    
    Do While endTime >= lastUpdateTime
        
        startTime = DateAdd("d", -MAX_PERIOD.WALLET_DEPOSIT_HISTORY, endTime)
        
        If startTime < lastUpdateTime Then startTime = lastUpdateTime
        
        Do
            Dim list As Collection
            Set list = BinanceAPI.Wallet_GetDepositHistory(filter, startTime, endTime, MAX_LIMIT.WALLET_DEPOSIT_HISTORY, includeSource:=True)
            
            Dim item As Dictionary
            For Each item In list
                stamp = xUnixTimestampToDate(item("completeTime"))
                
                If endTime > stamp Then endTime = stamp
                
                If item("status") = 1 Then
                    Dim amount, charge As Variant
                    Dim ticker, tranId, wallet As String
                    ticker = item("coin")
                    tranId = item("id")
                    wallet = IIf(item("walletType") = 0, WALLET_SPOT, WALLET_FUNDING)
                    amount = Str2Dec(item("amount"))
                    charge = 0
                    
                    AddSimpleOp ticker, stamp, wallet, OP_DEPOSIT, amount, 0, charge, item("network") & ":" & item("sourceAddress"), ID_PREFIX_CRYPTO_DEPOSIT & tranId
                End If
            Next item
            
            If list.count < MAX_LIMIT.WALLET_DEPOSIT_HISTORY Then
                endTime = DateAdd("s", -1, startTime)
                Exit Do
            Else
                endTime = DateAdd("s", -1, endTime)
            End If
        Loop
    Loop
End Sub

Public Sub Collect_CryptoWithdraws(ByVal lastUpdateTime As Date, Optional ByVal filter As String = "")
    Dim startTime, endTime, stamp As Date
    
    endTime = Me.endDate
    
    Do While endTime >= lastUpdateTime
        
        startTime = DateAdd("d", -MAX_PERIOD.WALLET_WITHDRAW_HISTORY, endTime)
        
        If startTime < lastUpdateTime Then startTime = lastUpdateTime
        
        Do
            Dim list As Collection
            Set list = BinanceAPI.Wallet_GetWithdrawHistory(filter, startTime, endTime, MAX_LIMIT.WALLET_WITHDRAW_HISTORY)
            
            Dim item As Dictionary
            For Each item In list
                stamp = item("completeTime")
                
                If endTime > stamp Then endTime = stamp
                
                If item("status") = 6 Then
                    Dim amount, charge As Variant
                    Dim ticker, tranId, wallet As String
                    ticker = item("coin")
                    tranId = item("id")
                    wallet = IIf(item("walletType") = 0, WALLET_SPOT, WALLET_FUNDING)
                    amount = Str2Dec(item("amount"))
                    charge = Str2Dec(item("transactionFee"))
                    
                    AddSimpleOp ticker, stamp, wallet, OP_WITHDRAW, 0, amount, charge, item("network") & ":" & item("address"), ID_PREFIX_CRYPTO_WITHDRAW & tranId
                End If
            Next item
            
            If list.count < MAX_LIMIT.WALLET_WITHDRAW_HISTORY Then
                endTime = DateAdd("s", -1, startTime)
                Exit Do
            Else
                endTime = DateAdd("s", -1, endTime)
            End If
        Loop
    Loop
End Sub

Public Sub Collect_Distributions(ByVal lastUpdateTime As Date, Optional ByVal filter As String = "")
    Dim startTime, endTime, stamp As Date
    
    endTime = Me.endDate
    
    Do While endTime >= lastUpdateTime
        
        startTime = DateAdd("d", -MAX_PERIOD.WALLET_DIVIDEND_HISTORY, endTime)
        
        If startTime < lastUpdateTime Then startTime = lastUpdateTime
        
        Do
            Dim data As Dictionary
            Set data = BinanceAPI.Wallet_GetDividendHistory(filter, startTime, endTime, MAX_LIMIT.WALLET_DIVIDEND_HISTORY)
            
            If Not data.Exists("rows") Then
                endTime = DateAdd("s", -1, startTime)
                Exit Do
            End If
            
            Dim list As Collection
            Set list = data("rows")
    
            Dim item As Dictionary
            For Each item In list
                stamp = xUnixTimestampToDate(item("divTime"))
                
                If endTime > stamp Then endTime = stamp
                
                Dim ticker As String
                ticker = item("asset")
                
                AddSimpleOp ticker, stamp, WALLET_SPOT, OP_DISTRIBUTION, item("amount"), 0, 0, item("enInfo"), ID_PREFIX_DISTRIBUTION & item("id")
            Next item
            
            If list.count < MAX_LIMIT.WALLET_DIVIDEND_HISTORY Then
                endTime = DateAdd("s", -1, startTime)
                Exit Do
            Else
                endTime = DateAdd("s", -1, endTime)
            End If
        Loop
    Loop
End Sub

Public Sub Collect_FlexibleEarnSubscriptions(ByVal lastUpdateTime As Date, Optional ByVal filter As String = "")
    Dim startTime, endTime, stamp As Date
    
    endTime = Me.endDate

    Do While endTime >= lastUpdateTime
    
        startTime = DateAdd("d", -MAX_PERIOD.SIMPLE_EARN_OPERATIONS, endTime)
        
        If startTime < lastUpdateTime Then startTime = lastUpdateTime
        
        Do
            Dim data As Dictionary
            Set data = BinanceAPI.SimpleEarn_GetFlexibleSubscriptions(filter, startTime, endTime, MAX_LIMIT.SIMPLE_EARN_OPERATIONS)
            
            If Not data.Exists("rows") Then
                endTime = DateAdd("s", -1, startTime)
                Exit Do
            End If
            
            Dim list As Collection
            Set list = data("rows")
        
            Dim item As Dictionary
            For Each item In list
                stamp = xUnixTimestampToDate(item("time"))
                
                If endTime > stamp Then endTime = stamp
                
                If item("status") = "SUCCESS" Then
                    Dim amount As Variant
                    Dim ticker, tranId As String
                    
                    ticker = item("asset")
                    amount = item("amount")
                    tranId = item("purchaseId")
                    
                    Select Case item("type")
                        Case "CONVERT"
                            AddSimpleOp ticker, stamp, WALLET_LOCKED_EARN, OP_WALLET_OUT, 0, 0, amount, "cvt:" & WALLET_FLEXIBLE_EARN, ID_PREFIX_WALLET_CVT & tranId
                        
                        Case "AUTO", "NORMAL"
                            If item("sourceAccount") = "SPOT" Or (item("sourceAccount") = "SPOTANDFUNDING" And item("amtFromSpot") > 0) Then
                                amount = IIf(item("sourceAccount") = "SPOT", amount, item("amtFromSpot"))

                                AddSimpleOp ticker, stamp, WALLET_SPOT, OP_WALLET_OUT, 0, amount, 0, "to:" & WALLET_FLEXIBLE_EARN, ID_PREFIX_WALLET_SEO & tranId
                            End If
                            If item("sourceAccount") = "FUNDING" Or (item("sourceAccount") = "SPOTANDFUNDING" And item("amtFromFunding") > 0) Then
                                amount = IIf(item("sourceAccount") = "FUNDING", amount, item("amtFromFunding"))

                                AddSimpleOp ticker, stamp, WALLET_FUNDING, OP_WALLET_OUT, 0, amount, 0, "to:" & WALLET_FLEXIBLE_EARN, ID_PREFIX_WALLET_FEO & tranId
                            End If
                    End Select
                End If
            Next item

            If list.count < MAX_LIMIT.SIMPLE_EARN_OPERATIONS Then
                endTime = DateAdd("s", -1, startTime)
                Exit Do
            Else
                endTime = DateAdd("s", -1, endTime)
            End If
        Loop
    Loop
End Sub

Public Sub Collect_FlexibleEarnRedemptions(ByVal lastUpdateTime As Date, Optional ByVal filter As String = "")
    Dim startTime, endTime, stamp As Date
    
    endTime = Me.endDate

    Do While endTime >= lastUpdateTime
    
        startTime = DateAdd("d", -MAX_PERIOD.SIMPLE_EARN_OPERATIONS, endTime)
        
        If startTime < lastUpdateTime Then startTime = lastUpdateTime
        
        Do
            Dim data As Dictionary
            Set data = BinanceAPI.SimpleEarn_GetFlexibleRedemptions(filter, startTime, endTime, MAX_LIMIT.SIMPLE_EARN_OPERATIONS)
            
            If Not data.Exists("rows") Then
                endTime = DateAdd("s", -1, startTime)
                Exit Do
            End If
            
            Dim list As Collection
            Set list = data("rows")
        
            Dim item As Dictionary
            For Each item In list
                stamp = xUnixTimestampToDate(item("time"))
                
                If endTime > stamp Then endTime = stamp
                
                If item("status") = "PAID" Then
                    Dim amount As Variant
                    Dim ticker, tranId As String
                    
                    ticker = item("asset")
                    amount = item("amount")
                    tranId = item("redeemId")
                    
                    Select Case item("destAccount")
                        Case "SPOT"
                            AddSimpleOp ticker, stamp, WALLET_SPOT, OP_WALLET_IN, amount, 0, 0, "from:" & WALLET_FLEXIBLE_EARN, ID_PREFIX_WALLET_SEI & tranId
                        Case "FUNDING"
                            AddSimpleOp ticker, stamp, WALLET_FUNDING, OP_WALLET_IN, amount, 0, 0, "from:" & WALLET_FLEXIBLE_EARN, ID_PREFIX_WALLET_FEI & tranId
                    End Select
                End If
            Next item
        
            If list.count < MAX_LIMIT.SIMPLE_EARN_OPERATIONS Then
                endTime = DateAdd("s", -1, startTime)
                Exit Do
            Else
                endTime = DateAdd("s", -1, endTime)
            End If
        Loop
    Loop
End Sub

Public Sub Collect_LockedEarnSubscriptions(ByVal lastUpdateTime As Date, Optional ByVal filter As String = "")

    Dim startTime, endTime, stamp As Date
    
    endTime = Me.endDate

    Do While endTime >= lastUpdateTime
    
        startTime = DateAdd("d", -MAX_PERIOD.SIMPLE_EARN_OPERATIONS, endTime)
        
        If startTime < lastUpdateTime Then startTime = lastUpdateTime
        
        Do
            Dim data As Dictionary
            Set data = BinanceAPI.SimpleEarn_GetLockedSubscriptions(filter, startTime, endTime, MAX_LIMIT.SIMPLE_EARN_OPERATIONS)
            
            If Not data.Exists("rows") Then
                endTime = DateAdd("s", -1, startTime)
                Exit Do
            End If
            
            Dim list As Collection
            Set list = data("rows")
        
            Dim item As Dictionary
            For Each item In list
                stamp = xUnixTimestampToDate(item("time"))
                
                If endTime > stamp Then endTime = stamp
                
                If item("status") = "SUCCESS" Then
                    Dim amount As Variant
                    Dim ticker, tranId As String
                    
                    ticker = item("asset")
                    amount = item("amount")
                    tranId = item("purchaseId")
                    
                    Select Case item("type")
                        Case "AUTO", "NORMAL"
                            If item("sourceAccount") = "SPOT" Or (item("sourceAccount") = "SPOTANDFUNDING" And item("amtFromSpot") > 0) Then
                                amount = IIf(item("sourceAccount") = "SPOT", amount, item("amtFromSpot"))

                                AddSimpleOp ticker, stamp, WALLET_SPOT, OP_WALLET_OUT, 0, amount, 0, "to:" & WALLET_LOCKED_EARN, ID_PREFIX_WALLET_SEO & tranId
                                AddSimpleOp ticker, stamp, WALLET_LOCKED_EARN, OP_WALLET_IN, amount, 0, 0, "from:" & WALLET_SPOT, ID_PREFIX_WALLET_SEI & tranId
                            End If
                            If item("sourceAccount") = "FUNDING" Or (item("sourceAccount") = "SPOTANDFUNDING" And item("amtFromFunding") > 0) Then
                                amount = IIf(item("sourceAccount") = "FUNDING", amount, item("amtFromFunding"))

                                AddSimpleOp ticker, stamp, WALLET_FUNDING, OP_WALLET_OUT, 0, amount, 0, "to:" & WALLET_LOCKED_EARN, ID_PREFIX_WALLET_FEO & tranId
                                AddSimpleOp ticker, stamp, WALLET_LOCKED_EARN, OP_WALLET_IN, amount, 0, 0, "from:" & WALLET_FUNDING, ID_PREFIX_WALLET_FEI & tranId
                            End If
                    End Select
                End If
            Next item
        
            If list.count < MAX_LIMIT.SIMPLE_EARN_OPERATIONS Then
                endTime = DateAdd("s", -1, startTime)
                Exit Do
            Else
                endTime = DateAdd("s", -1, endTime)
            End If
        Loop
    Loop
End Sub

Public Sub Collect_LockedEarnRedemptions(ByVal lastUpdateTime As Date, Optional ByVal filter As String = "")

    Dim startTime, endTime, stamp As Date
    
    endTime = Me.endDate

    Do While endTime >= lastUpdateTime
    
        startTime = DateAdd("d", -MAX_PERIOD.SIMPLE_EARN_OPERATIONS, endTime)
        
        If startTime < lastUpdateTime Then startTime = lastUpdateTime
        
        Do
            Dim data As Dictionary
            Set data = BinanceAPI.SimpleEarn_GetLockedRedemptions(filter, startTime, endTime, MAX_LIMIT.SIMPLE_EARN_OPERATIONS)
            
            If Not data.Exists("rows") Then
                endTime = DateAdd("s", -1, startTime)
                Exit Do
            End If
            
            Dim list As Collection
            Set list = data("rows")
        
            Dim item As Dictionary
            For Each item In list
                stamp = xUnixTimestampToDate(item("time"))
                
                If endTime > stamp Then endTime = stamp
                
                If item("status") = "PAID" Then
                    Dim amount As Variant
                    Dim ticker, tranId As String
                    
                    ticker = item("asset")
                    amount = item("amount")
                    tranId = item("redeemId")
                    
                    Select Case item("type")
                        Case "MATURE", "AHEAD"
                            AddSimpleOp ticker, stamp, WALLET_LOCKED_EARN, OP_WALLET_OUT, 0, amount, 0, "to:" & WALLET_SPOT, ID_PREFIX_WALLET_SEO & tranId
                            AddSimpleOp ticker, stamp, WALLET_SPOT, OP_WALLET_IN, amount, 0, 0, "from:" & WALLET_LOCKED_EARN, ID_PREFIX_WALLET_SEI & tranId

                        Case "NEW_TRANSFERRED"
                    
                    End Select
                End If
            Next item
        
            If list.count < MAX_LIMIT.SIMPLE_EARN_OPERATIONS Then
                endTime = DateAdd("s", -1, startTime)
                Exit Do
            Else
                endTime = DateAdd("s", -1, endTime)
            End If
        Loop
    Loop
End Sub

Public Sub AddSimpleOp( _
    ByVal ticker As String, _
    ByVal stamp As Date, _
    ByVal wallet As String, _
    ByVal operation As String, _
    ByVal acquired As Variant, _
    ByVal spent As Variant, _
    ByVal charge As Variant, _
    ByVal note As String, _
    ByVal id As String _
    )
    Dim op As BinanceOp
    Set op = New BinanceOp
    
    op.stamp = stamp
    op.wallet = wallet
    op.operation = operation
    op.acquired = acquired
    op.spent = spent
    op.ticker = IIf(charge = 0, "", ticker)
    op.price = 0
    op.amount = 0
    op.charge = charge
    op.bnb_fee = 0
    op.note = note
    op.id = id
    
    xAddOp ticker, op
End Sub

Public Sub AddExchangeOp( _
    ByVal stamp As Date, _
    ByVal tickerOut As String, _
    ByVal tickerIn As String, _
    ByVal walletOut As String, _
    ByVal walletIn As String, _
    ByVal operationOut As String, _
    ByVal operationIn As String, _
    ByVal amountOut As Variant, _
    ByVal amountIn As Variant, _
    ByVal priceOut As Variant, _
    ByVal priceIn As Variant, _
    ByVal charge As Variant, _
    ByVal bnbFee As Variant, _
    ByVal id As String, _
    ByVal prefixOut As String, _
    ByVal prefixIn As String, _
    ByVal noteOut As String, _
    Optional ByVal noteIn As String = "" _
    )
    If noteIn = "" Then noteIn = noteOut

    Dim op As BinanceOp
    Set op = New BinanceOp
    
    op.stamp = stamp
    op.wallet = walletOut
    op.operation = operationOut
    op.acquired = 0
    op.spent = amountOut
    op.ticker = tickerIn
    op.price = priceOut
    op.amount = amountIn
    op.charge = charge
    op.bnb_fee = bnbFee
    op.note = noteOut
    op.id = prefixOut & id
    
    xAddOp tickerOut, op

    Set op = New BinanceOp
    
    op.stamp = stamp
    op.wallet = walletIn
    op.operation = operationIn
    op.acquired = amountIn
    op.spent = 0
    op.ticker = tickerOut
    op.price = priceIn
    op.amount = amountOut
    op.charge = charge
    op.bnb_fee = bnbFee
    op.note = noteIn
    op.id = prefixIn & id
    
    xAddOp tickerIn, op
    
    If bnbFee <> 0 Then
        Assert_Fail 'must be implemented
    End If
    
End Sub

Private Sub xAddOp(ByVal ticker As String, ByVal op As BinanceOp)
    If Not m_ops.Exists(ticker) Then m_ops.Add ticker, New Collection

    m_ops(ticker).Add op
End Sub

Private Function xUnixTimestampToDate(ByVal unixTime As LongLong) As Date
    xUnixTimestampToDate = DateAdd("s", unixTime / 1000, "1/1/1970")
End Function

Private Sub xHandleSimpleConversion(ByVal item As Dictionary, ByVal stamp As Date, ByVal wallet As String)
    AddExchangeOp _
        stamp, _
        item("fromAsset"), _
        item("toAsset"), _
        wallet, _
        WALLET_SPOT, _
        OP_CONVERT_OUT, _
        OP_CONVERT_IN, _
        item("fromAmount"), _
        item("toAmount"), _
        item("ratio"), _
        item("inverseRatio"), _
        0, 0, _
        item("orderId"), _
        ID_PREFIX_CONVERT_OUT, _
        ID_PREFIX_CONVERT_IN, _
        item("orderType") & ":" & item("side")
End Sub

Private Sub xHandleCombosConversion(ByVal item As Dictionary, ByVal stamp As Date, ByVal spotWallet As Boolean, ByVal fundingWallet As Boolean, ByVal earnWallet As Boolean)
    TableConversions.HandleComboConversion Me, item, stamp, spotWallet, fundingWallet, earnWallet
End Sub

