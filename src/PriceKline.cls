VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PriceKline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public openTime As Date
Public closeTime As Date
Public openPrice As Variant
Public highPrice As Variant
Public lowPrice As Variant
Public closePrice As Variant

Public Property Get avaregePriceStd() As Variant
    avaregePriceStd = (openPrice + highPrice + lowPrice + closePrice) / CDec(4)
End Property

Public Property Get avaregePriceTech() As Variant
    avaregePriceTech = (highPrice + lowPrice + closePrice) / CDec(3)
End Property

Public Sub InitFromBinance(ByVal item As Collection)
    openTime = UnixTimestampToDate(item(1))
    openPrice = Str2Dec(item(2))
    highPrice = Str2Dec(item(3))
    lowPrice = Str2Dec(item(4))
    closePrice = Str2Dec(item(5))
    closeTime = UnixTimestampToDate(item(7))
End Sub

