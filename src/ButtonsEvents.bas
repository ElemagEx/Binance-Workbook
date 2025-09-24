Attribute VB_Name = "ButtonsEvents"
Option Explicit

Public Sub test1()
    
    Dim assets As collection
    Set assets = BinanceApi_Wallet_GetUserAssets()
    
    Dim asset As Dictionary
    For Each asset In assets
        Debug.Print asset("asset") & " : " & asset("free")
    Next asset
    
End Sub
