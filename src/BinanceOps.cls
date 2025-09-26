VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BinanceOps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_operations As Dictionary
'
'
'
Public Sub Class_Initialize()
    Set m_operations = New Dictionary
End Sub
'
'
'
Public Property Get Operations() As Dictionary
    Set Operations = m_operations
End Property
'
'
'
Public Sub Collect_Wallet_TransferHistory(ByVal lastUpdateTime As Date, Optional ByVal filter As String = "")
    Dim types As New collection
    types.Add "MAIN_FUNDING"
    types.Add "FUNDING_MAIN"
    
    Dim walletSource As String
    Dim walletTarget As String
    
    Dim transferType As Variant
    For Each transferType In types
        
        If Left(transferType, 5) = "MAIN_" Then
            walletSource = WALLET_SPOT
        ElseIf Left(transferType, 8) = "FUNDING_" Then
            walletSource = WALLET_FUNDING
        Else
            MsgBox "Error: Unknown source wallet"
            Exit Sub
        End If
        If Right(transferType, 5) = "_MAIN" Then
            walletTarget = WALLET_SPOT
        ElseIf Right(transferType, 8) = "_FUNDING" Then
            walletTarget = WALLET_FUNDING
        Else
            MsgBox "Error: Unknown target wallet"
            Exit Sub
        End If
        
        Dim startTime, endTime, opTime As Date
        
        endTime = now()
        
        Do While endTime >= lastUpdateTime
            
            startTime = DateAdd("d", -MAX_PERIOD_WALLET_TRANSFERS, endTime)
            
            If startTime < lastUpdateTime Then
                startTime = lastUpdateTime
            End If
        
            Do
                Dim data As Dictionary
                Set data = BinanceApi_Wallet_GetTransferHistory(transferType, startTime, endTime, MAX_LIMIT_WALLET_TRANSFERS)
                
                If data Is Nothing Then
                    Exit Sub
                End If
                If Not data.Exists("rows") Then
                    endTime = DateAdd("s", -1, startTime)
                    Exit Do
                End If
                
                Dim list As collection
                Set list = data("rows")
    
                Dim item As Dictionary
                For Each item In list
                    opTime = UnixTimestamp2Date(item("timestamp"))
                    
                    If endTime > opTime Then
                        endTime = opTime
                    End If
                
                    If item("status") = "CONFIRMED" Then
                        Dim amount As Variant
                        Dim ticker, opId As String
                        ticker = item("asset")
                        
                        If filter = "" Or filter = ticker Then
                            amount = item("amount")
                            opId = item("tranId")
                        
                            xAddTicker ticker
                            
                            xAddWalletTransferOps ticker, opTime, amount, opId, walletSource, walletTarget, walletTarget, walletSource, ID_PREFIX_WALLET_OUT, ID_PREFIX_WALLET_IN
                        End If
                    End If
                Next item
                
                If list.count < MAX_LIMIT_WALLET_TRANSFERS Then
                    endTime = DateAdd("s", -1, startTime)
                    Exit Do
                Else
                    endTime = DateAdd("s", -1, endTime)
                End If
            Loop
        Loop
    Next transferType
End Sub
'
'
'
Public Sub Collect_SimpleEarn_FlexibleSubscriptions(ByVal lastUpdateTime As Date, Optional ByVal filter As String = "")

    Dim startTime, endTime, opTime As Date
    
    endTime = now()

    Do While endTime >= lastUpdateTime
    
        startTime = DateAdd("d", -MAX_PERIOD_SIMPLEEARN_FLEXIBLE_SUBSCRIPTIONS, endTime)
        
        If startTime < lastUpdateTime Then
            startTime = lastUpdateTime
        End If
        
        Do
            Dim data As Dictionary
            Set data = BinanceApi_SimpleEarn_GetFlexibleSubscriptions(filter, startTime, endTime, MAX_LIMIT_SIMPLEEARN_FLEXIBLE_SUBSCRIPTIONS)
        
            If data Is Nothing Then
                Exit Sub
            End If
            If Not data.Exists("rows") Then
                endTime = DateAdd("s", -1, startTime)
                Exit Do
            End If
            
            Dim list As collection
            Set list = data("rows")
        
            Dim item As Dictionary
            For Each item In list
                opTime = UnixTimestamp2Date(item("time"))
                
                If endTime > opTime Then
                    endTime = opTime
                End If
                
                If item("status") = "SUCCESS" Then
                    Dim amount As Variant
                    Dim ticker, opId As String
                    
                    ticker = item("asset")
                    amount = item("amount")
                    opId = item("purchaseId")
                    
                    xAddTicker ticker
                
                    Select Case item("type")
                        Case "CONVERT"
                            xAddWalletTransferOps ticker, opTime, amount, opId, WALLET_EARN, WALLET_NONE, WALLET_DYNAMIC, "", ID_PREFIX_WALLET_SEO, ""
                        
                        Case "AUTO", "NORMAL"
                            If item("sourceAccount") = "SPOT" Or (item("sourceAccount") = "SPOTANDFUNDING" And item("amtFromSpot") > 0) Then
                                amount = IIf(item("sourceAccount") = "SPOT", amount, item("amtFromSpot"))
                            
                                xAddWalletTransferOps ticker, opTime, amount, opId, WALLET_SPOT, WALLET_NONE, WALLET_DYNAMIC, "", ID_PREFIX_WALLET_OUT, ""
                            End If
                            If item("sourceAccount") = "FUNDING" Or (item("sourceAccount") = "SPOTANDFUNDING" And item("amtFromFunding") > 0) Then
                                amount = IIf(item("sourceAccount") = "FUNDING", amount, item("amtFromFunding"))

                                xAddWalletTransferOps ticker, opTime, amount, opId, WALLET_FUNDING, WALLET_NONE, WALLET_DYNAMIC, "", ID_PREFIX_WALLET_OUT, ""
                            End If

                        Case Else
                            MsgBox "Unsupported Flexible Simple Earn Subscription Operation Type: " & item("type")
                    End Select
                End If
            Next item
        
            If list.count < MAX_LIMIT_SIMPLEEARN_FLEXIBLE_SUBSCRIPTIONS Then
                endTime = DateAdd("s", -1, startTime)
                Exit Do
            Else
                endTime = DateAdd("s", -1, endTime)
            End If
        Loop
    Loop
End Sub
'
'
'
Public Sub Collect_SimpleEarn_FlexibleRedemptions(ByVal lastUpdateTime As Date, Optional ByVal filter As String = "")

    Dim startTime, endTime, opTime As Date
    
    endTime = now()

    Do While endTime >= lastUpdateTime
    
        startTime = DateAdd("d", -MAX_PERIOD_SIMPLEEARN_FLEXIBLE_REDEMPTIONS, endTime)
        
        If startTime < lastUpdateTime Then
            startTime = lastUpdateTime
        End If
        
        Do
            Dim data As Dictionary
            Set data = BinanceApi_SimpleEarn_GetFlexibleRedemptions(filter, startTime, endTime, MAX_LIMIT_SIMPLEEARN_FLEXIBLE_REDEMPTIONS)
        
            If data Is Nothing Then
                Exit Sub
            End If
            If Not data.Exists("rows") Then
                endTime = DateAdd("s", -1, startTime)
                Exit Do
            End If
            
            Dim list As collection
            Set list = data("rows")
        
            Dim item As Dictionary
            For Each item In list
                opTime = UnixTimestamp2Date(item("time"))
                
                If endTime > opTime Then
                    endTime = opTime
                End If
                
                If item("status") = "PAID" Then
                    Dim amount As Variant
                    Dim ticker, opId, toWallet As String
                    
                    ticker = item("asset")
                    amount = item("amount")
                    opId = item("redeemId")
                    
                    xAddTicker ticker
                    
                    Select Case item("destAccount")
                        Case "SPOT"
                            toWallet = WALLET_SPOT
                        Case "FUNDING"
                            toWallet = WALLET_FUNDING
                        Case Else
                            toWallet = WALLET_NONE
                            MsgBox "Unsupported Flexible Simple Earn Redemption destination wallet: " & item("destAcount")
                    End Select
                    
                    xAddWalletTransferOps ticker, opTime, amount, opId, WALLET_NONE, toWallet, "", WALLET_DYNAMIC, "", ID_PREFIX_WALLET_IN
                End If
            Next item
        
            If list.count < MAX_LIMIT_SIMPLEEARN_FLEXIBLE_REDEMPTIONS Then
                endTime = DateAdd("s", -1, startTime)
                Exit Do
            Else
                endTime = DateAdd("s", -1, endTime)
            End If
        Loop
    Loop
End Sub
'
'
'
Public Sub Collect_SimpleEarn_LockedSubscriptions(ByVal lastUpdateTime As Date, Optional ByVal filter As String = "")

    Dim startTime, endTime, opTime As Date
    
    endTime = now()

    Do While endTime >= lastUpdateTime
    
        startTime = DateAdd("d", -MAX_PERIOD_LOCKED_FLEXIBLE_SUBSCRIPTIONS, endTime)
        
        If startTime < lastUpdateTime Then
            startTime = lastUpdateTime
        End If
        
        Do
            Dim data As Dictionary
            Set data = BinanceApi_SimpleEarn_GetLockedSubscriptions(filter, startTime, endTime, MAX_LIMIT_SIMPLEEARN_LOCKED_SUBSCRIPTIONS)
        
            If data Is Nothing Then
                Exit Sub
            End If
            If Not data.Exists("rows") Then
                endTime = DateAdd("s", -1, startTime)
                Exit Do
            End If
            
            Dim list As collection
            Set list = data("rows")
        
            Dim item As Dictionary
            For Each item In list
                opTime = UnixTimestamp2Date(item("time"))
                
                If endTime > opTime Then
                    endTime = opTime
                End If
                
                If item("status") = "SUCCESS" Then
                    Dim amount As Variant
                    Dim ticker, opId As String
                    
                    ticker = item("asset")
                    amount = item("amount")
                    opId = item("purchaseId")
                    
                    xAddTicker ticker
                    
                    Select Case item("type")
                        Case "AUTO", "NORMAL"
                            If item("sourceAccount") = "SPOT" Or (item("sourceAccount") = "SPOTANDFUNDING" And item("amtFromSpot") > 0) Then
                                amount = IIf(item("sourceAccount") = "SPOT", amount, item("amtFromSpot"))
                            
                                xAddWalletTransferOps ticker, opTime, amount, opId, WALLET_SPOT, WALLET_EARN, WALLET_EARN, WALLET_SPOT, ID_PREFIX_WALLET_OUT, ID_PREFIX_WALLET_SEI
                            End If
                            If item("sourceAccount") = "FUNDING" Or (item("sourceAccount") = "SPOTANDFUNDING" And item("amtFromFunding") > 0) Then
                                amount = IIf(item("sourceAccount") = "FUNDING", amount, item("amtFromFunding"))

                                xAddWalletTransferOps ticker, opTime, amount, opId, WALLET_FUNDING, WALLET_EARN, WALLET_EARN, WALLET_FUNDING, ID_PREFIX_WALLET_OUT, ID_PREFIX_WALLET_SEI
                            End If

                        Case Else
                            MsgBox "Unsupported Flexible Simple Earn Subscription Operation Type: " & item("type")
                    End Select
                End If
            Next item
        
            If list.count < MAX_LIMIT_SIMPLEEARN_LOCKED_SUBSCRIPTIONS Then
                endTime = DateAdd("s", -1, startTime)
                Exit Do
            Else
                endTime = DateAdd("s", -1, endTime)
            End If
        Loop
    Loop
End Sub
'
'
'
Public Sub Collect_SimpleEarn_LockedRedemptions(ByVal lastUpdateTime As Date, Optional ByVal filter As String = "")

    Dim startTime, endTime, opTime As Date
    
    endTime = now()

    Do While endTime >= lastUpdateTime
    
        startTime = DateAdd("d", -MAX_PERIOD_LOCKED_FLEXIBLE_REDEMPTIONS, endTime)
        
        If startTime < lastUpdateTime Then
            startTime = lastUpdateTime
        End If
        
        Do
            Dim data As Dictionary
            Set data = BinanceApi_SimpleEarn_GetLockedRedemptions(filter, startTime, endTime, MAX_LIMIT_SIMPLEEARN_LOCKED_REDEMPTIONS)
        
            If data Is Nothing Then
                Exit Sub
            End If
            If Not data.Exists("rows") Then
                endTime = DateAdd("s", -1, startTime)
                Exit Do
            End If
            
            Dim list As collection
            Set list = data("rows")
        
            Dim item As Dictionary
            For Each item In list
                opTime = UnixTimestamp2Date(item("time"))
                
                If endTime > opTime Then
                    endTime = opTime
                End If
                
                If item("status") = "PAID" Then
                    Dim amount As Variant
                    Dim ticker, opId As String
                    
                    ticker = item("asset")
                    amount = item("amount")
                    opId = item("redeemId")
                    
                    xAddTicker ticker
                    
                    Select Case item("type")
                        Case "NEW_TRANSFERRED"
                            
                        Case "MATURE", "AHEAD"
                            xAddWalletTransferOps ticker, opTime, amount, opId, WALLET_EARN, WALLET_SPOT, WALLET_SPOT, WALLET_EARN, ID_PREFIX_WALLET_SEO, ID_PREFIX_WALLET_IN
                        Case Else
                            MsgBox "Unsupported Flexible Simple Earn Redemption operation type: " & item("type")
                    End Select
                End If
            Next item
        
            If list.count < MAX_LIMIT_SIMPLEEARN_FLEXIBLE_REDEMPTIONS Then
                endTime = DateAdd("s", -1, startTime)
                Exit Do
            Else
                endTime = DateAdd("s", -1, endTime)
            End If
        Loop
    Loop
End Sub


Private Sub xAddTicker(ByVal ticker As String)
    If Not m_operations.Exists(ticker) Then
        m_operations.Add ticker, New collection
    End If
End Sub

Private Sub xAddWalletTransferOps( _
    ByVal ticker As String, _
    ByVal opTime As Date, _
    ByVal amount As Variant, _
    ByVal opId As String, _
    ByVal walletOut As String, _
    ByVal walletIn As String, _
    ByVal noteSuffixOut As String, _
    ByVal noteSuffixIn As String, _
    ByVal idPrefixOut As String, _
    ByVal idPrefixIn As String _
    )
    Dim operation As Dictionary
    
    If walletOut <> WALLET_NONE Then
        Set operation = New Dictionary
        
        operation(COL_DETAILS_STAMP) = DateAdd("s", -1, opTime)
        operation(COL_DETAILS_WALLET) = walletOut
        operation(COL_DETAILS_OPERATION) = OPERATION_WALLET_OUT
        operation(COL_DETAILS_ACQUIRED) = 0
        operation(COL_DETAILS_SPENT) = amount
        operation(COL_DETAILS_TICKER) = ""
        operation(COL_DETAILS_PRICE) = 0
        operation(COL_DETAILS_AMOUNT) = 0
        operation(COL_DETAILS_CHARGE) = 0
        operation(COL_DETAILS_BNB_FEE) = 0
        operation(COL_DETAILS_NOTE) = "to:" & noteSuffixOut
        operation(COL_DETAILS_ID) = idPrefixOut & opId
    
        m_operations(ticker).Add operation
    End If

    If walletIn <> WALLET_NONE Then
        Set operation = New Dictionary
        
        operation(COL_DETAILS_STAMP) = opTime
        operation(COL_DETAILS_WALLET) = walletIn
        operation(COL_DETAILS_OPERATION) = OPERATION_WALLET_IN
        operation(COL_DETAILS_ACQUIRED) = amount
        operation(COL_DETAILS_SPENT) = 0
        operation(COL_DETAILS_TICKER) = ""
        operation(COL_DETAILS_PRICE) = 0
        operation(COL_DETAILS_AMOUNT) = 0
        operation(COL_DETAILS_CHARGE) = 0
        operation(COL_DETAILS_BNB_FEE) = 0
        operation(COL_DETAILS_NOTE) = "from:" & noteSuffixIn
        operation(COL_DETAILS_ID) = idPrefixIn & opId
    
        m_operations(ticker).Add operation
    End If
End Sub

