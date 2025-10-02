Attribute VB_Name = "TableLastUpdate"
Option Explicit
Option Private Module

Private Const FIRST_DATA_ROW = 2

Private Const COL_WALLET As String = "Wallet"
Private Const COL_TRANSFERS As String = "Transfers"
Private Const COL_CONVERTIONS As String = "Conversions"
Private Const COL_DUST_LOG As String = "Dust Log"
Private Const COL_FIAT_BUYS As String = "Fiat Buys"
Private Const COL_FIAT_SELLS As String = "Fiat Sells"
Private Const COL_FIAT_DEPOSITS As String = "Fiat Deposits"
Private Const COL_FIAT_WITHDRAWS As String = "Fiat Withdraws"
Private Const COL_CRYPTO_DEPOSITS As String = "Crypto Deposits"
Private Const COL_CRYPTO_WITHDRAWS As String = "Crypto Withdraws"
Private Const COL_DISTRIBUTIONS As String = "Distributions"
Private Const COL_FLEXIBLE_EARNS As String = "Flexible Earns"
Private Const COL_LOCKED_EARNS As String = "Locked Earns"

Public Function IsDataCleanedUp()
    IsDataCleanedUp = False
    
    Dim table As ListObject
    Set table = xGetTable()

    Dim i As Long
    For i = FIRST_DATA_ROW To table.ListColumns.count
        If Not IsEmpty(table.ListColumns(i).DataBodyRange(1).Value) Then
            Exit Function
        End If
    Next i
    
    IsDataCleanedUp = True
End Function

Public Sub CleanUpData()
    ClearData True
End Sub

Public Sub ClearData(ByVal clearWalletData As Boolean)
    Dim table As ListObject
    Set table = xGetTable()

    Dim i As Long
    For i = FIRST_DATA_ROW To table.ListColumns.count
        If table.ListColumns(i).name <> COL_WALLET Or clearWalletData Then
            table.ListColumns(i).DataBodyRange(1).ClearContents
        End If
    Next i
End Sub

Public Property Get wallet() As Date
    wallet = xGetLastUpdate(COL_WALLET)
End Property

Public Property Let wallet(ByVal val As Date)
    xSetLastUpdate COL_WALLET, val
End Property

Public Property Get Transfers() As Date
    Transfers = xGetLastUpdate(COL_TRANSFERS)
End Property

Public Property Let Transfers(ByVal val As Date)
    xSetLastUpdate COL_TRANSFERS, val
End Property

Public Property Get Convertions() As Date
    Convertions = xGetLastUpdate(COL_CONVERTIONS)
End Property

Public Property Let Convertions(ByVal val As Date)
    xSetLastUpdate COL_CONVERTIONS, val
End Property

Public Property Get DustLog() As Date
    DustLog = xGetLastUpdate(COL_DUST_LOG)
End Property

Public Property Let DustLog(ByVal val As Date)
    xSetLastUpdate COL_DUST_LOG, val
End Property

Public Property Get FiatBuys() As Date
    FiatBuys = xGetLastUpdate(COL_FIAT_BUYS)
End Property

Public Property Let FiatBuys(ByVal val As Date)
    xSetLastUpdate COL_FIAT_BUYS, val
End Property

Public Property Get FiatSells() As Date
    FiatSells = xGetLastUpdate(COL_FIAT_SELLS)
End Property

Public Property Let FiatSells(ByVal val As Date)
    xSetLastUpdate COL_FIAT_SELLS, val
End Property

Public Property Get FiatDeposits() As Date
    FiatDeposits = xGetLastUpdate(COL_FIAT_DEPOSITS)
End Property

Public Property Let FiatDeposits(ByVal val As Date)
    xSetLastUpdate COL_FIAT_DEPOSITS, val
End Property

Public Property Get FiatWithdraws() As Date
    FiatWithdraws = xGetLastUpdate(COL_FIAT_WITHDRAWS)
End Property

Public Property Let FiatWithdraws(ByVal val As Date)
    xSetLastUpdate COL_FIAT_WITHDRAWS, val
End Property

Public Property Get CryptoDeposits() As Date
    CryptoDeposits = xGetLastUpdate(COL_CRYPTO_DEPOSITS)
End Property

Public Property Let CryptoDeposits(ByVal val As Date)
    xSetLastUpdate COL_CRYPTO_DEPOSITS, val
End Property

Public Property Get CryptoWithdraws() As Date
    CryptoWithdraws = xGetLastUpdate(COL_CRYPTO_WITHDRAWS)
End Property

Public Property Let CryptoWithdraws(ByVal val As Date)
    xSetLastUpdate COL_CRYPTO_WITHDRAWS, val
End Property

Public Property Get Distributions() As Date
    Distributions = xGetLastUpdate(COL_DISTRIBUTIONS)
End Property

Public Property Let Distributions(ByVal val As Date)
    xSetLastUpdate COL_DISTRIBUTIONS, val
End Property

Public Property Get FlexibleEarns() As Date
    FlexibleEarns = xGetLastUpdate(COL_FLEXIBLE_EARNS)
End Property

Public Property Let FlexibleEarns(ByVal val As Date)
    xSetLastUpdate COL_FLEXIBLE_EARNS, val
End Property

Public Property Get LockedEarns() As Date
    LockedEarns = xGetLastUpdate(COL_LOCKED_EARNS)
End Property

Public Property Let LockedEarns(ByVal val As Date)
    xSetLastUpdate COL_LOCKED_EARNS, val
End Property

Public Function xGetTable() As ListObject
    Set xGetTable = ThisWorkbook.Sheets(SHEET_ASSETS).ListObjects(TABLE_LAST_UPDATE)
End Function

Private Function xGetLastUpdate(ByVal name As String) As Date
    Dim cell As Range
    Set cell = xGetTable().ListColumns(name).DataBodyRange(1)
    xGetLastUpdate = IIf(IsEmpty(cell.Value), TableUserInfo.ActivityStartDate, CDate(cell.Value))
End Function

Private Sub xSetLastUpdate(ByVal name As String, ByVal val As Date)
    Dim cell As Range
    Set cell = xGetTable().ListColumns(name).DataBodyRange(1)
    If val = 0 Then
        cell.ClearContents
    Else
        cell.Value = val
    End If
End Sub
