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

Public Property Get tickers() As Variant
    tickers = m_coins.Keys
End Property

Public Property Get item(ByVal ticker As String) As BinanceCoin
    If m_coins.Exists(ticker) Then
        Set item = m_coins(ticker)
    Else
        Set item = Nothing
    End If
End Property

Public Function Contains(ByVal ticker As String) As Boolean
    Contains = m_coins.Exists(ticker)
End Function

Public Sub AddTicker(ByVal ticker As String)
    If Not m_coins.Exists(ticker) Then
        m_coins.Add ticker, Nothing
    End If
End Sub

Public Sub RemoveUncollected()
    Dim keysToRemove As New Collection
    Dim ticker As Variant
    For Each ticker In m_coins.Keys
        If m_coins(ticker) Is Nothing Then
            keysToRemove.Add ticker
        End If
    Next ticker
    For Each ticker In keysToRemove
        m_coins.remove ticker
    Next ticker
End Sub

Public Sub collectBasicInfo(ByVal updateOnlyExisting As Boolean)
    Dim list As Collection
    Set list = BinanceAPI.Wallet_GetAllCoinsInfo()

    Dim item As Dictionary
    For Each item In list
        Dim ticker As String
        ticker = item("coin")

        Dim coin As BinanceCoin
        
        If m_coins.Exists(ticker) Then
            Set coin = m_coins(ticker)
            If coin Is Nothing Then
                Set m_coins(ticker) = New BinanceCoin
                Set coin = m_coins(ticker)
            End If
        ElseIf Not updateOnlyExisting Then
            Set coin = New BinanceCoin
            m_coins.Add ticker, coin
        Else
            Set coin = Nothing
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
            
            If Not coin Is Nothing Then
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
                    
                    If Not coin Is Nothing Then
                        coin.quotes = coin.quotes + 1
                    
                        If coin.precision > symbol("quoteAssetPrecision") Then
                            coin.precision = symbol("quoteAssetPrecision")
                        End If
                    End If
                End If
            End If
        End If
    Next symbol
End Sub
