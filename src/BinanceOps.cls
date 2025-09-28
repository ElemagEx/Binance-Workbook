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
'Public Const ID_PREFIX_WALLET_IN As String = "WI-"
'Public Const ID_PREFIX_WALLET_OUT As String = "WO-"
'Public Const ID_PREFIX_WALLET_SEI As String = "WS-"
'Public Const ID_PREFIX_WALLET_SEO As String = "WR-"

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
    endDate = IIf(m_endDate = 0, now(), m_endDate)
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

Public Sub AddTicker(ByVal ticker As String)
    If Not m_coins.Exists(ticker) Then
        m_coins.Add ticker, New Collection
    End If
End Sub

Public Sub Collect_FiatBuys(ByVal lastUpdateTime As Date)
    Dim startTime, endTime, opTime As Date
    
    endTime = Me.endDate
    
    Do
        If endTime < lastUpdateTime Then
            Exit Do
        End If
        
        startTime = DateAdd("d", -MAX_PERIOD.FIAT_OPERATIONS, endTime)
        
        If startTime < lastUpdateTime Then
            startTime = lastUpdateTime
        End If
        
        Do
            Dim data As Dictionary
            Set data = BinanceAPI.Fiat_GetBuys(startTime, endTime, MAX_LIMIT.FIAT_OPERATIONS)
            
            If data Is Nothing Then
                Exit Sub
            End If
            If Not data.Exists("data") Then
                endTime = DateAdd("s", -1, startTime)
                Exit Do
            End If
            
            Dim list As Collection
            Set list = data("data")
    
            Dim item As Dictionary
            For Each item In list
                opTime = UnixTimestamp2Date(item("updateTime"))
                
                If endTime > opTime Then
                    endTime = opTime
                End If
                
                If item("status") = "Completed" Then
                    Dim ticker As String
                    ticker = item("fiatCurrency")
                    
                    xAddTicker ticker
                    
                    Dim op As BinanceOp
                    Set op = New BinanceOp
                    
                    op.stamp = DateAdd("s", -1, opTime)
                    op.wallet = WALLET_SPOT
                    op.operation = OPERATION_FIAT_BUY
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
                    
                    op.stamp = opTime
                    op.wallet = WALLET_SPOT
                    op.operation = OPERATION_CONVERT_OUT
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

                    op.stamp = opTime
                    op.wallet = WALLET_SPOT
                    op.operation = OPERATION_CONVERT_IN
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
    Dim startTime, endTime, opTime As Date
    
    endTime = Me.endDate
    
    Do
        If endTime < lastUpdateTime Then
            Exit Do
        End If
        
        startTime = DateAdd("d", -MAX_PERIOD.FIAT_OPERATIONS, endTime)
        
        If startTime < lastUpdateTime Then
            startTime = lastUpdateTime
        End If
        
        Do
            Dim data As Dictionary
            Set data = BinanceAPI.Fiat_GetSells(startTime, endTime, MAX_LIMIT.FIAT_OPERATIONS)
            
            If data Is Nothing Then
                Exit Sub
            End If
            If Not data.Exists("data") Then
                endTime = DateAdd("s", -1, startTime)
                Exit Do
            End If
            
            Dim list As Collection
            Set list = data("data")
    
            Dim item As Dictionary
            For Each item In list
                opTime = UnixTimestamp2Date(item("updateTime"))
                
                If endTime > opTime Then
                    endTime = opTime
                End If
                
                If item("status") = "Completed" Then
                    Dim ticker As String
                    ticker = item("cryptoCurrency")
                    
                    xAddTicker ticker
                    
                    Dim op As BinanceOp
                    Set op = New BinanceOp

                    op.stamp = DateAdd("s", -1, opTime)
                    op.wallet = WALLET_SPOT
                    op.operation = OPERATION_CONVERT_OUT
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
                    
                    op.stamp = DateAdd("s", -1, opTime)
                    op.wallet = WALLET_SPOT
                    op.operation = OPERATION_CONVERT_IN
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
                    
                    op.stamp = opTime
                    op.wallet = WALLET_SPOT
                    op.operation = OPERATION_FIAT_SELL
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
    Dim startTime, endTime, opTime As Date
    
    endTime = Me.endDate
    
    Do
        If endTime < lastUpdateTime Then
            Exit Do
        End If
        
        startTime = DateAdd("d", -MAX_PERIOD.FIAT_OPERATIONS, endTime)
        
        If startTime < lastUpdateTime Then
            startTime = lastUpdateTime
        End If
        
        Do
            Dim data As Dictionary
            Set data = BinanceAPI.Fiat_GetDeposits(startTime, endTime, MAX_LIMIT.FIAT_OPERATIONS)
            
            If data Is Nothing Then
                Exit Sub
            End If
            If Not data.Exists("data") Then
                endTime = DateAdd("s", -1, startTime)
                Exit Do
            End If
            
            Dim list As Collection
            Set list = data("data")
    
            Dim item As Dictionary
            For Each item In list
                opTime = UnixTimestamp2Date(item("updateTime"))
                
                If endTime > opTime Then
                    endTime = opTime
                End If
                
                If item("status") = "Successful" Then
                    Dim ticker As String
                    ticker = item("fiatCurrency")
                    
                    xAddTicker ticker
                    
                    Dim op As BinanceOp
                    Set op = New BinanceOp
                    
                    op.stamp = opTime
                    op.wallet = WALLET_SPOT
                    op.operation = OPERATION_DEPOSIT
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
    Dim startTime, endTime, opTime As Date
    
    endTime = Me.endDate
    
    Do
        If endTime < lastUpdateTime Then
            Exit Do
        End If
        
        startTime = DateAdd("d", -MAX_PERIOD.FIAT_OPERATIONS, endTime)
        
        If startTime < lastUpdateTime Then
            startTime = lastUpdateTime
        End If
        
        Do
            Dim data As Dictionary
            Set data = BinanceAPI.Fiat_GetWithdraws(startTime, endTime, MAX_LIMIT.FIAT_OPERATIONS)
            
            If data Is Nothing Then
                Exit Sub
            End If
            If Not data.Exists("data") Then
                endTime = DateAdd("s", -1, startTime)
                Exit Do
            End If
            
            Dim list As Collection
            Set list = data("data")
    
            Dim item As Dictionary
            For Each item In list
                opTime = UnixTimestamp2Date(item("updateTime"))
                
                If endTime > opTime Then
                    endTime = opTime
                End If
                
                If item("status") = "Successful" Then
                    Dim ticker As String
                    ticker = item("fiatCurrency")
                    
                    xAddTicker ticker
                    
                    Dim op As BinanceOp
                    Set op = New BinanceOp
                    
                    op.stamp = opTime
                    op.wallet = WALLET_SPOT
                    op.operation = OPERATION_WITHDRAW
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

Public Sub Collect_Distributions(ByVal lastUpdateTime As Date, Optional ByVal filter As String = "")
    Dim startTime, endTime, opTime As Date
    
    endTime = Me.endDate
    
    Do
        If endTime < lastUpdateTime Then
            Exit Do
        End If
        
        startTime = DateAdd("d", -MAX_PERIOD.WALLET_DIVIDEND_HISTORY, endTime)
        
        If startTime < lastUpdateTime Then
            startTime = lastUpdateTime
        End If
        
        Do
            Dim data As Dictionary
            Set data = BinanceAPI.Wallet_GetDividendHistory(filter, startTime, endTime, MAX_LIMIT.WALLET_DIVIDEND_HISTORY)
            
            If data Is Nothing Then
                Exit Sub
            End If
            If Not data.Exists("rows") Then
                endTime = DateAdd("s", -1, startTime)
                Exit Do
            End If
            
            Dim list As Collection
            Set list = data("rows")
    
            Dim item As Dictionary
            For Each item In list
                opTime = UnixTimestamp2Date(item("divTime"))
                
                If endTime > opTime Then
                    endTime = opTime
                End If
                
                Dim ticker As String
                ticker = item("asset")
                
                xAddTicker ticker
                
                Dim op As BinanceOp
                Set op = New BinanceOp
                
                op.stamp = opTime
                op.wallet = WALLET_SPOT
                op.operation = OPERATION_DISTRIBUTION
                op.acquired = item("amount")
                op.spent = 0
                op.ticker = ""
                op.price = 0
                op.amount = 0
                op.charge = 0
                op.bnb_fee = 0
                op.note = item("enInfo")
                op.id = ID_PREFIX_DISTRIBUTION & item("id")
                
                xAddOp ticker, op
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

Private Sub xAddTicker(ByVal ticker As String)
    If Not m_ops.Exists(ticker) Then
        m_ops.Add ticker, New Collection
    End If
End Sub

Private Sub xAddOp(ByVal ticker As String, ByVal op As BinanceOp)
    m_ops(ticker).Add op
End Sub

