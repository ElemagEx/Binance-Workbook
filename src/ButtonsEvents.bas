Attribute VB_Name = "ButtonsEvents"
Option Explicit

Public Const SHEET_ASSETS As String = "Assets"

Public Function Str2Dec(ByVal str As String) As Variant
    Dim val As String
    val = Replace(str, ".", Application.DecimalSeparator)
    Str2Dec = CDec(val)
End Function

Public Sub test1()
        
    TableUserInfo.CleanUpData
        
End Sub
