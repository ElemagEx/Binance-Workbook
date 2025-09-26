Attribute VB_Name = "TableUserInfo"
Option Explicit

Private Const SHEET_NAME As String = SHEET_ASSETS
Private Const TABLE_NAME As String = "UserInfo"

Private Const COL_UID As String = "UID"
Private Const COL_MAKER_FEE As String = "Maker Fee"
Private Const COL_TAKER_FEE As String = "Taker Fee"
Private Const COL_BUYER_FEE As String = "Buyer Fee"
Private Const COL_SELLER_FEE As String = "Seller Fee"
Private Const COL_ACTIVITY_START_DATE As String = "Activity Start Date"

Public Function IsDataCleanedUp()
    Dim table As ListObject
    Set table = ThisWorkbook.Sheets(SHEET_NAME).ListObjects(TABLE_NAME)

    IsCleanedUp = IsEmpty(table.ListColumns(COL_ACCOUNT_UID).DataBodyRange(1).Value) _
        And IsEmpty(table.ListColumns(COL_ACCOUNT_MAKER_FEE).DataBodyRange(1).Value) _
        And IsEmpty(table.ListColumns(COL_ACCOUNT_TAKER_FEE).DataBodyRange(1).Value) _
        And IsEmpty(table.ListColumns(COL_ACCOUNT_BUYER_FEE).DataBodyRange(1).Value) _
        And IsEmpty(table.ListColumns(COL_ACCOUNT_SELLER_FEE).DataBodyRange(1).Value)
End Function

Public Sub CleanUpData()
    Dim table As ListObject
    Set table = ThisWorkbook.Sheets(SHEET_NAME).ListObjects(TABLE_NAME)
    
    table.ListColumns(COL_UID).DataBodyRange.ClearContents
    table.ListColumns(COL_MAKER_FEE).DataBodyRange.ClearContents
    table.ListColumns(COL_TAKER_FEE).DataBodyRange.ClearContents
    table.ListColumns(COL_BUYER_FEE).DataBodyRange.ClearContents
    table.ListColumns(COL_SELLER_FEE).DataBodyRange.ClearContents
End Sub

Public Sub UpdateData()
    Dim account As Dictionary
    Set account = BinanceAPI.SpotTrading_GetAccountInfo()
    
    If account Is Nothing Then
        Exit Sub
    End If

    Dim rates As Dictionary
    Set rates = account("commissionRates")
    
    Dim table As ListObject
    Set table = ThisWorkbook.Sheets(SHEET_NAME).ListObjects(TABLE_NAME)
    
    table.ListColumns(COL_UID).DataBodyRange(1).Value = account("uid")
    table.ListColumns(COL_MAKER_FEE).DataBodyRange(1).Value = Str2Dec(rates("maker"))
    table.ListColumns(COL_TAKER_FEE).DataBodyRange(1).Value = Str2Dec(rates("taker"))
    table.ListColumns(COL_BUYER_FEE).DataBodyRange(1).Value = Str2Dec(rates("buyer"))
    table.ListColumns(COL_SELLER_FEE).DataBodyRange(1).Value = Str2Dec(rates("seller"))
End Sub

Public Property Get ActivityStartDate() As Date
    Dim table As ListObject
    Set table = ThisWorkbook.Sheets(SHEET_NAME).ListObjects(TABLE_NAME)
    
    ActivityStartDate = CDate(table.ListColumns(COL_ACTIVITY_START_DATE).DataBodyRange(1).Value)
End Property
