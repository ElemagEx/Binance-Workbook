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

'Public Const ID_PREFIX_DRIBBLET As String = "FF-"
Private Const ID_PREFIX_WALLET_IN As String = "WI-"  ' Universal Transfer In
Private Const ID_PREFIX_WALLET_OUT As String = "WO-" ' Universal Transfer Out
Private Const ID_PREFIX_WALLET_CVT As String = "WC-" ' Locked Earn to Flexible Earn Transfer
Private Const ID_PREFIX_WALLET_STE As String = "WS-" ' Spot to Flexible/Locked Earn
Private Const ID_PREFIX_WALLET_ETS As String = "WE-" ' Flexible/Locked Earn to Spot
Private Const ID_PREFIX_WALLET_FTE As String = "WF-" ' Funding to Flexible/Locked Earn
Private Const ID_PREFIX_WALLET_ETF As String = "WA-" ' Flexible/Locked Earn to Funding

Private Const ID_PREFIX_CONVERT_IN As String = "VI-"
Private Const ID_PREFIX_CONVERT_OUT As String = "VO-"
Private Const ID_PREFIX_CONVERT_FSI As String = "VF-" ' transfer from funding to spot before convert to spot asset
Private Const ID_PREFIX_CONVERT_FSO As String = "VG-" ' transfer from funding to spot before convert to spot asset
Private Const ID_PREFIX_CONVERT_ESI As String = "VE-" ' transfer from earn to spot before convert to spot asset
Private Const ID_PREFIX_CONVERT_ESO As String = "VD-" ' transfer from earn to spot before convert to spot asset
Private Const ID_PREFIX_CONVERT_EFI As String = "VX-" ' input transfer from earn to funding before convert funding asset
Private Const ID_PREFIX_CONVERT_EFO As String = "VY-" ' output transfer from earn to funding before convert funding asset
Private Const ID_PREFIX_DISTRIBUTION As String = "RI-"
Private Const ID_PREFIX_TRADE As String = "XC-"
Private Const ID_PREFIX_FIAT_BUY As String = "TB-"
Private Const ID_PREFIX_FIAT_SELL As String = "TS-"
Private Const ID_PREFIX_FIAT_DEPOSIT As String = "TD-"
Private Const ID_PREFIX_FIAT_WITHDRAW As String = "TW-"
Private Const ID_PREFIX_CRYPTO_DEPOSIT As String = "ND-"
Private Const ID_PREFIX_CRYPTO_WITHDRAW As String = "NW-"

Private m_ops As New Dictionary
Private m_endDate As Date

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
                        
                        xAddSimpleOp ticker, stamp, sourceWallet, OP_WALLET_OUT, 0, amount, 0, "to:" & targetWallet, ID_PREFIX_WALLET_OUT & tranId
                        xAddSimpleOp ticker, stamp, targetWallet, OP_WALLET_IN, amount, 0, 0, "from:" & sourceWallet, ID_PREFIX_WALLET_IN & tranId
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
                    Dim ticker As String
                    ticker = item("fiatCurrency")
                    
                    xAddTicker ticker
                    
                    Dim op As BinanceOp
                    Set op = New BinanceOp
                    
                    op.stamp = DateAdd("s", -1, stamp)
                    op.wallet = WALLET_SPOT
                    op.operation = OP_FIAT_BUY
                    op.acquired = item("sourceAmount")
                    op.spent = 0
                    op.ticker = ticker
                    op.price = 0
                    op.amount = 0
                    op.charge = item("totalFee")
                    op.bnb_fee = 0
                    op.note = STR_NA
                    op.id = ID_PREFIX_FIAT_BUY & item("orderNo")
                    
                    xAddOp ticker, op
                    
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
                    op.id = ID_PREFIX_CONVERT_OUT & item("orderNo")
                    
                    xAddOp ticker, op
                
                    ticker = item("cryptoCurrency")
                    
                    xAddTicker ticker
                    
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
                    op.id = ID_PREFIX_CONVERT_IN & item("orderNo")
                
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
                    Dim ticker As String
                    ticker = item("cryptoCurrency")
                    
                    xAddTicker ticker
                    
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
                    
                    xAddTicker ticker
                    
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
                    
                    Set op = New BinanceOp
                    
                    op.stamp = stamp
                    op.wallet = WALLET_SPOT
                    op.operation = OP_FIAT_SELL
                    op.acquired = 0
                    op.spent = item("sourceAmount")
                    op.ticker = ticker
                    op.price = 0
                    op.amount = 0
                    op.charge = item("totalFee")
                    op.bnb_fee = 0
                    op.note = STR_NA
                    op.id = ID_PREFIX_FIAT_SELL & item("orderNo")
                    
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
                    Dim ticker As String
                    ticker = item("fiatCurrency")
                    
                    xAddTicker ticker
                    
                    Dim op As BinanceOp
                    Set op = New BinanceOp
                    
                    op.stamp = stamp
                    op.wallet = WALLET_SPOT
                    op.operation = OP_DEPOSIT
                    op.acquired = item("indicatedAmount")
                    op.spent = 0
                    op.ticker = ticker
                    op.price = 0
                    op.amount = 0
                    op.charge = item("totalFee")
                    op.bnb_fee = 0
                    op.note = item("method")
                    op.id = ID_PREFIX_FIAT_DEPOSIT & item("orderNo")
                    
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
                    Dim ticker As String
                    ticker = item("fiatCurrency")
                    
                    xAddTicker ticker
                    
                    Dim op As BinanceOp
                    Set op = New BinanceOp
                    
                    op.stamp = stamp
                    op.wallet = WALLET_SPOT
                    op.operation = OP_WITHDRAW
                    op.acquired = 0
                    op.spent = item("amount")
                    op.ticker = ticker
                    op.price = 0
                    op.amount = 0
                    op.charge = item("totalFee")
                    op.bnb_fee = 0
                    op.note = item("method")
                    op.id = ID_PREFIX_FIAT_WITHDRAW & item("orderNo")
                    
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
                    Dim ticker As String
                    ticker = item("coin")
                    
                    xAddTicker ticker
                    
                    Dim op As BinanceOp
                    Set op = New BinanceOp
                    
                    op.stamp = stamp
                    op.wallet = IIf(item("walletType") = 0, WALLET_SPOT, WALLET_FUNDING)
                    op.operation = OP_DEPOSIT
                    op.acquired = item("amount")
                    op.spent = 0
                    op.ticker = ""
                    op.price = 0
                    op.amount = 0
                    op.charge = 0
                    op.bnb_fee = 0
                    op.note = item("network") & ":" & item("sourceAddress")
                    op.id = ID_PREFIX_CRYPTO_DEPOSIT & item("id")
                    
                    xAddOp ticker, op
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
                    Dim ticker As String
                    ticker = item("coin")
                    
                    xAddTicker ticker
                    
                    Dim op As BinanceOp
                    Set op = New BinanceOp
                    
                    op.stamp = stamp
                    op.wallet = IIf(item("walletType") = 0, WALLET_SPOT, WALLET_FUNDING)
                    op.operation = OP_WITHDRAW
                    op.acquired = 0
                    op.spent = item("amount")
                    op.ticker = ticker
                    op.price = 0
                    op.amount = 0
                    op.charge = item("transactionFee")
                    op.bnb_fee = 0
                    op.note = item("network") & ":" & item("address")
                    op.id = ID_PREFIX_CRYPTO_WITHDRAW & item("id")
                    
                    xAddOp ticker, op
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
                
                xAddSimpleOp ticker, stamp, WALLET_SPOT, OP_DISTRIBUTION, item("amount"), 0, 0, item("enInfo"), ID_PREFIX_DISTRIBUTION & item("id")
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
                            xAddSimpleOp ticker, stamp, WALLET_LOCKED_EARN, OP_WALLET_OUT, 0, 0, amount, "cvt:" & WALLET_FLEXIBLE_EARN, ID_PREFIX_WALLET_CVT & tranId
                        
                        Case "AUTO", "NORMAL"
                            If item("sourceAccount") = "SPOT" Or (item("sourceAccount") = "SPOTANDFUNDING" And item("amtFromSpot") > 0) Then
                                amount = IIf(item("sourceAccount") = "SPOT", amount, item("amtFromSpot"))

                                xAddSimpleOp ticker, stamp, WALLET_SPOT, OP_WALLET_OUT, 0, amount, 0, "to:" & WALLET_FLEXIBLE_EARN, ID_PREFIX_WALLET_STE & tranId
                            End If
                            If item("sourceAccount") = "FUNDING" Or (item("sourceAccount") = "SPOTANDFUNDING" And item("amtFromFunding") > 0) Then
                                amount = IIf(item("sourceAccount") = "FUNDING", amount, item("amtFromFunding"))

                                xAddSimpleOp ticker, stamp, WALLET_FUNDING, OP_WALLET_OUT, 0, amount, 0, "to:" & WALLET_FLEXIBLE_EARN, ID_PREFIX_WALLET_FTE & tranId
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
                            xAddSimpleOp ticker, stamp, WALLET_SPOT, OP_WALLET_IN, amount, 0, 0, "from:" & WALLET_FLEXIBLE_EARN, ID_PREFIX_WALLET_ETS & tranId
                        Case "FUNDING"
                            xAddSimpleOp ticker, stamp, WALLET_FUNDING, OP_WALLET_IN, amount, 0, 0, "from:" & WALLET_FLEXIBLE_EARN, ID_PREFIX_WALLET_ETF & tranId
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

                                xAddSimpleOp ticker, stamp, WALLET_SPOT, OP_WALLET_OUT, 0, amount, 0, "to:" & WALLET_LOCKED_EARN, ID_PREFIX_WALLET_STE & tranId
                                xAddSimpleOp ticker, stamp, WALLET_LOCKED_EARN, OP_WALLET_IN, amount, 0, 0, "from:" & WALLET_SPOT, ID_PREFIX_WALLET_ETS & tranId
                            End If
                            If item("sourceAccount") = "FUNDING" Or (item("sourceAccount") = "SPOTANDFUNDING" And item("amtFromFunding") > 0) Then
                                amount = IIf(item("sourceAccount") = "FUNDING", amount, item("amtFromFunding"))

                                xAddSimpleOp ticker, stamp, WALLET_FUNDING, OP_WALLET_OUT, 0, amount, 0, "to:" & WALLET_LOCKED_EARN, ID_PREFIX_WALLET_FTE & tranId
                                xAddSimpleOp ticker, stamp, WALLET_LOCKED_EARN, OP_WALLET_IN, amount, 0, 0, "from:" & WALLET_FUNDING, ID_PREFIX_WALLET_ETF & tranId
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
                            xAddSimpleOp ticker, stamp, WALLET_LOCKED_EARN, OP_WALLET_OUT, 0, amount, 0, "to:" & WALLET_SPOT, ID_PREFIX_WALLET_ETS & tranId
                            xAddSimpleOp ticker, stamp, WALLET_SPOT, OP_WALLET_IN, amount, 0, 0, "from:" & WALLET_LOCKED_EARN, ID_PREFIX_WALLET_STE & tranId

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
Private Sub xAddTicker(ByVal ticker As String)
    If Not m_ops.Exists(ticker) Then
        m_ops.Add ticker, New Collection
    End If
End Sub

Private Sub xAddOp(ByVal ticker As String, ByVal op As BinanceOp)
    m_ops(ticker).Add op
End Sub

Public Function xUnixTimestampToDate(ByVal unixTime As LongLong) As Date
    xUnixTimestampToDate = DateAdd("s", unixTime / 1000, "1/1/1970")
End Function

Private Sub xAddSimpleOp( _
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
    xAddTicker ticker
    
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

