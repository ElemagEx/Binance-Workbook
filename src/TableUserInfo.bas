Attribute VB_Name = "TableUserInfo"
Option Explicit

Private Const TABLE_NAME_USER_INFO As String = "UserInfo"

Private Const COL_USER_INFO_UID As String = "UID"
Private Const COL_USER_INFO_MAKER_FEE As String = "Maker Fee"
Private Const COL_USER_INFO_TAKER_FEE As String = "Taker Fee"
Private Const COL_USER_INFO_BUYER_FEE As String = "Buyer Fee"
Private Const COL_USER_INFO_SELLER_FEE As String = "Seller Fee"
Private Const COL_USER_INFO_ACTIVITY_START_DATE As String = "Activity Start Date"

Public Sub Hello1()

End Sub


Public Static Sub ClearData()
    Dim table As ListObject

    Set table = ThisWorkbook.Sheets(SHEET_ASSETS).ListObjects(TABLE_NAME_USER_INFO)
    
    table.ListColumns(COL_USER_INFO_UID).DataBodyRange(1).ClearContents
    table.ListColumns(COL_USER_INFO_MAKER_FEE).DataBodyRange(1).ClearContents
    table.ListColumns(COL_USER_INFO_TAKER_FEE).DataBodyRange(1).ClearContents
    table.ListColumns(COL_USER_INFO_BUYER_FEE).DataBodyRange(1).ClearContents
    table.ListColumns(COL_USER_INFO_SELLER_FEE).DataBodyRange(1).ClearContents
End Sub

Public Static Sub RefreshData()
    Dim account As Dictionary
    Set account = BinanceApi_SpotTrading_GetAccountInfo()
    
    If account Is Nothing Then
        Exit Sub
    End If

    Dim rates As Dictionary
    Set rates = account("commissionRates")
    
    Dim table As ListObject
    Set table = ThisWorkbook.Sheets(SHEET_ASSETS).ListObjects(TABLE_ACCOUNT)
    
    table.ListColumns(COL_ACCOUNT_UID).DataBodyRange(1).Value = account("uid")
    table.ListColumns(COL_ACCOUNT_MAKER_FEE).DataBodyRange(1).Value = Str2Dec(rates("maker"))
    table.ListColumns(COL_ACCOUNT_TAKER_FEE).DataBodyRange(1).Value = Str2Dec(rates("taker"))
    table.ListColumns(COL_ACCOUNT_BUYER_FEE).DataBodyRange(1).Value = Str2Dec(rates("buyer"))
    table.ListColumns(COL_ACCOUNT_SELLER_FEE).DataBodyRange(1).Value = Str2Dec(rates("seller"))
End Sub

