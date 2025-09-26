Attribute VB_Name = "TableLastUpdate"
Option Explicit
Option Private Module

Private Const SHEET_NAME As String = SHEET_ASSETS
Private Const TABLE_NAME As String = "LastUpdate"

Private Const COL_WALLET As String = "Wallet"
Private Const COL_FIAT_BUY As String = "Fiat Buy"
Private Const COL_FIAT_SELL As String = "Fiat Sell"
Private Const COL_FIAT_DEPOSIT As String = "Fiat Deposit"
Private Const COL_FIAT_WITHDRAW As String = "Fiat Withdraw"
Private Const COL_CRYPTO_DEPOSIT As String = "Crypto Deposit"
Private Const COL_CRYPTO_WITHDRAW As String = "Crypto Withdraw"
Private Const COL_CONVERT As String = "Convert"
Private Const COL_DISTRIBUTION As String = "Distribution"
Private Const COL_TRANSFER As String = "Transfer"
Private Const COL_DRIBBLETS As String = "Dribblets"

Public Function IsDataCleanedUp()
    Dim table As ListObject
    Set table = ThisWorkbook.Sheets(SHEET_NAME).ListObjects(TABLE_NAME)

    IsCleanedUp = IsEmpty(table.ListColumns(COL_WALLET).DataBodyRange(1).Value) _
        And IsEmpty(table.ListColumns(COL_FIAT_BUY).DataBodyRange(1).Value) _
        And IsEmpty(table.ListColumns(COL_FIAT_SELL).DataBodyRange(1).Value) _
        And IsEmpty(table.ListColumns(COL_FIAT_DEPOSIT).DataBodyRange(1).Value) _
        And IsEmpty(table.ListColumns(COL_FIAT_WITHDRAW).DataBodyRange(1).Value) _
        And IsEmpty(table.ListColumns(COL_CRYPTO_DEPOSIT).DataBodyRange(1).Value) _
        And IsEmpty(table.ListColumns(COL_CRYPTO_WITHDRAW).DataBodyRange(1).Value) _
        And IsEmpty(table.ListColumns(COL_CONVERT).DataBodyRange(1).Value) _
        And IsEmpty(table.ListColumns(COL_DISTRIBUTION).DataBodyRange(1).Value) _
        And IsEmpty(table.ListColumns(COL_TRANSFER).DataBodyRange(1).Value) _
        And IsEmpty(table.ListColumns(COL_DRIBBLETS).DataBodyRange(1).Value)
End Function

Public Sub CleanUpData()
    Dim table As ListObject
    Set table = ThisWorkbook.Sheets(SHEET_NAME).ListObjects(TABLE_NAME)
    
    table.ListColumns(COL_WALLET).DataBodyRange.ClearContents
    table.ListColumns(COL_FIAT_BUY).DataBodyRange.ClearContents
    table.ListColumns(COL_FIAT_SELL).DataBodyRange.ClearContents
    table.ListColumns(COL_FIAT_DEPOSIT).DataBodyRange.ClearContents
    table.ListColumns(COL_FIAT_WITHDRAW).DataBodyRange.ClearContents
    table.ListColumns(COL_CRYPTO_DEPOSIT).DataBodyRange.ClearContents
    table.ListColumns(COL_CRYPTO_WITHDRAW).DataBodyRange.ClearContents
    table.ListColumns(COL_CONVERT).DataBodyRange.ClearContents
    table.ListColumns(COL_DISTRIBUTION).DataBodyRange.ClearContents
    table.ListColumns(COL_TRANSFER).DataBodyRange.ClearContents
    table.ListColumns(COL_DRIBBLETS).DataBodyRange.ClearContents
End Sub

Public Property Get Wallet() As Date
    Wallet = xGetLastUpdate(COL_WALLET)
End Property

Public Property Let Wallet(ByVal val As Date)
    xSetLastUpdate COL_WALLET, val
End Property

Public Property Get FiatBuy() As Date
    Wallet = xGetLastUpdate(COL_FIAT_BUY)
End Property

Public Property Let FiatBuy(ByVal val As Date)
    xSetLastUpdate COL_FIAT_BUY, val
End Property

Public Property Get FiatSell() As Date
    Wallet = xGetLastUpdate(COL_FIAT_SELL)
End Property

Public Property Let FiatSell(ByVal val As Date)
    xSetLastUpdate COL_FIAT_SELL, val
End Property
Public Function xGetLastUpdate(ByVal name As String) As Date
    Dim cell As range
    Set cell = ThisWorkbook.Sheets(SHEET_NAME).ListObjects(TABLE_NAME).ListColumns(name).DataBodyRange(1)
    xGetLastUpdate = IIf(IsEmpty(cell.Value), TableUserInfo.ActivityStartDate, CDate(cell.Value))
End Function

Public Sub xSetLastUpdate(ByVal name As String, ByVal val As Date)
    Dim cell As range
    Set cell = ThisWorkbook.Sheets(SHEET_NAME).ListObjects(TABLE_NAME).ListColumns(name).DataBodyRange(1)
    If val = 0 Then
        cell.ClearContents
    Else
        cell.Value = val
    End If
End Sub
