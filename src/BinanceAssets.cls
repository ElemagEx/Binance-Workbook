VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BinanceAssets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_wallet As New Dictionary
'
'
'
Public Property Get wallet() As Dictionary
    Set wallet = m_wallet
End Property
'
'
'
Public Sub CollectSpotAssets(ByVal name As String, Optional ByVal filter As String = "")

    Dim assets As Collection
    Set assets = BinanceAPI.Wallet_GetUserAssets(filter, False)
    
    Dim ticker As String
    Dim amount As Variant

    Dim asset As Dictionary
    For Each asset In assets
        ticker = asset("asset")
        amount = Str2Dec(asset("free")) + Str2Dec(asset("locked")) + Str2Dec(asset("freeze")) + Str2Dec(asset("withdrawing"))
        
        xAddTicker ticker
        xAddAmount ticker, name, amount
    Next asset

End Sub
'
'
'
Public Sub CollectFundingAssets(ByVal name As String, Optional ByVal filter As String = "")

    Dim assets As Collection
    Set assets = BinanceAPI.Wallet_GetFundingAssets(filter, False)

    Dim ticker As String
    Dim amount As Variant
    
    Dim asset As Dictionary
    For Each asset In assets
        ticker = asset("asset")
        amount = Str2Dec(asset("free")) + Str2Dec(asset("locked")) + Str2Dec(asset("freeze")) + Str2Dec(asset("withdrawing"))
        
        xAddTicker ticker
        xAddAmount ticker, name, amount
    Next asset

End Sub
'
'
'
Public Sub CollectSimpleEarnFlexibleAssets(ByVal name As String, Optional ByVal filter As String = "")

    Dim position As Dictionary
    Dim positions As Dictionary
    
    Dim ticker As String
    Dim amount As Variant
    
    Dim page As Long
    For page = 1 To 1000
        Set positions = BinanceAPI.SimpleEarn_GetFlexiblePositions(filter, MAX_LIMIT.SIMPLE_EARN_OPERATIONS, page)
        
        If Not positions.Exists("rows") Then
            Exit For
        End If
        
        For Each position In positions("rows")
            ticker = position("asset")
            amount = position("totalAmount")
            
            xAddTicker ticker
            xAddAmount ticker, name, amount
        Next position
        
        If positions("rows").count < MAX_LIMIT.SIMPLE_EARN_OPERATIONS Then
            Exit For
        End If
    Next page
End Sub
'
'
'
Public Sub CollectSimpleEarnLockedAssets(ByVal name As String, Optional ByVal filter As String = "")
    Dim position As Dictionary
    Dim positions As Dictionary
    
    Dim ticker As String
    Dim amount As Variant
    
    Dim page As Long
    For page = 1 To 1000
        Set positions = BinanceAPI.SimpleEarn_GetLockedPositions(filter, MAX_LIMIT.SIMPLE_EARN_OPERATIONS, page)
        
        If Not positions.Exists("rows") Then
            Exit For
        End If
        
        For Each position In positions("rows")
            ticker = position("asset")
            amount = position("amount")
            
            xAddTicker ticker
            xAddAmount ticker, name, amount
        Next position
        
        If positions("rows").count < MAX_LIMIT.SIMPLE_EARN_OPERATIONS Then
            Exit For
        End If
    Next page
End Sub

Private Sub xAddTicker(ByVal ticker As String)
    If Not m_wallet.Exists(ticker) Then
        m_wallet.Add ticker, New Dictionary
    End If
End Sub

Private Sub xAddAmount(ByVal ticker As String, ByVal name As String, ByVal amount As Variant)
    If Not m_wallet(ticker).Exists(name) Then
        m_wallet(ticker).Add name, amount
    Else
        m_wallet(ticker)(name) = m_wallet(ticker)(name) + amount
    End If
End Sub

