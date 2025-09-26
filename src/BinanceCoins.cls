VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BinanceCoins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_tz As String
Private m_coins As New Dictionary

Public Property Get ExchangeTimezone() As String
    ExchangeTimezone = m_tz
End Property

Public Sub collectBasicInfo(ByVal updateOnlyExisting As Boolean)
    Dim list As collection
    Set list = BinanceAPI.Wallet_GetAllCoinsInfo()

    If list Is Nothing Then
        Exit Sub
    End If
    
    Dim item As Dictionary
    For Each item In list
        Dim ticker As String
        ticker = item("coin")
        
        Dim coin As BinanceCoin
        
        If m_coins.Exists(ticker) Then
            Set coin = m_coins(ticker)
        ElseIf Not updateOnlyExisting Then
            Set coin = New BinanceCoin
            m_coins.Add ticker, coin
        End If
        
        If Not coin Is Nothing Then
            coin.ticker = ticker
            coin.name = item("name")
            coin.isFiat = item("isLegalMoney")
            coin.bases = 0
            coin.quotes = 0
            coin.trades.RemoveAll
        End If
    Next item
End Sub

Public Sub collectExchangeInfo()
    Dim info As Dictionary
    Set info = BinanceAPI.SpotTrading_GetExchangeInfo(False)
    
    If info Is Nothing Then
        Exit Sub
    End If
    
    m_tz = info("timezone")
    
    Dim symbols As collection
    Set symbols = info("symbols")
    
    Dim symbol As Dictionary
    For Each symbol In symbols
        Dim ticker As String
        ticker = symbol("baseAsset")
        
        If m_coins.Exists(ticker) Then
            Dim coin As BinanceCoin
            Set coin = m_coins(ticker)
            coin.bases = coin.bases + 1
            
            If coin.precision > symbol("baseAssetPrecision") Then
                coin.precision = symbol("baseAssetPrecision")
            End If
        
            ticker = symbol("quoteAsset")
            If m_coins.Exists(ticker) Then
                coin.trades.Add ticker, False
                
                Set coin = m_coins(ticker)
                coin.quotes = coin.quotes + 1
            
                If coin.precision > symbol("quoteAssetPrecision") Then
                    coin.precision = symbol("quoteAssetPrecision")
                End If
            End If
        End If
    Next symbol
End Sub
