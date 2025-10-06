VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SheetAssets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.count > 1 Then Exit Sub
    
    Dim cell As Range
    Set cell = Me.ListObjects(TABLE_USER_INFO).ListColumns("Quote").DataBodyRange(1)
    
    If Target.Row <> cell.Row Or Target.Column <> cell.Column Then Exit Sub
    
    Application.EnableEvents = False
    TableWallet.ClearPrices
    Application.EnableEvents = True
End Sub

