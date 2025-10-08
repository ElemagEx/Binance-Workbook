VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PriceQuote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ticker As String
Public method As String
Public refers As String
Public format As String

Public Property Get isFiat() As Boolean
    isFiat = method <> STR_EVAL_METHOD_MARKET
End Property

Public Property Let isFiat(ByVal val As Boolean)
    method = IIf(val, STR_EVAL_METHOD_EXCHANGE, STR_EVAL_METHOD_MARKET)
    refers = ""
    format = IIf(val, NUMBER_FORMAT_MAIN_FIAT, NUMBER_FORMAT_MAIN_CRYPTO)
End Property
