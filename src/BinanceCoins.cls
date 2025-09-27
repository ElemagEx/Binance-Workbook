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

Public Property Get Tickers() As Variant
    Tickers = m_coins.Keys
End Property

Public Property Get item(ByVal ticker As String) As BinanceCoin
    Set item = IIf(m_coins.Exists(ticker), m_coins(ticker), Nothing)
End Property

Public Function Contains(ByVal ticker As String) As Boolean
    Contains = m_coins.Exists(ticker)
End Function

Public Sub AddTicker(ByVal ticker As String)
    If Not m_coins.Exists(ticker) Then
        m_coins.Add ticker, New BinanceCoin
    End If
End Sub

Public Sub collectBasicInfo(ByVal updateOnlyExisting As Boolean)
    Dim list As Collection
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
            coin.precision = IIf(coin.isFiat, 2, 8)
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
    
    Dim symbols As Collection
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
                If Not coin.markets.Exists(ticker) Then
                    coin.markets.Add ticker, False
                End If
                
                Set coin = m_coins(ticker)
                coin.quotes = coin.quotes + 1
            
                If coin.precision > symbol("quoteAssetPrecision") Then
                    coin.precision = symbol("quoteAssetPrecision")
                End If
            End If
        End If
    Next symbol
End Sub
