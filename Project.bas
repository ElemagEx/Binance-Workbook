Attribute VB_Name = "Project"
Option Explicit
Option Private Module

Public Function GetProjectFiles(ByVal thirdParty As Boolean) As Collection
    Dim files As New Collection

    If thirdParty Then
        files.Add "VBA-Dictionary/Dictionary.cls"
        files.Add "VBA-JSON/JsonConverter.bas"
        files.Add "VBA-Web/src/WebHelpers.bas"
        files.Add "VBA-Web/src/WebClient.cls"
        files.Add "VBA-Web/src/WebRequest.cls"
        files.Add "VBA-Web/src/WebResponse.cls"
        files.Add "VBA-Web/src/IWebAuthenticator.cls"
    Else
        files.Add "src/KeysAndSecrets.bas"
        files.Add "src/Main.bas"
        files.Add "src/Actions.bas"
        files.Add "src/BinanceAPI.bas"
        files.Add "src/BinanceOp.cls"
        files.Add "src/BinanceOps.cls"
        files.Add "src/BinanceCoin.cls"
        files.Add "src/BinanceCoins.cls"
        files.Add "src/BinanceAssets.cls"
        files.Add "src/BinanceLedger.cls"
        files.Add "src/BinanceLedgers.bas"
        files.Add "src/BinanceWebQuery.cls"
        files.Add "src/PriceEvaluator.cls"
        files.Add "src/PriceQuote.cls"
        files.Add "src/TableWallet.bas"
        files.Add "src/TableDustLog.bas"
        files.Add "src/TableUserInfo.bas"
        files.Add "src/TableCurrencies.bas"
        files.Add "src/TableEvaluation.bas"
        files.Add "src/TableLastUpdate.bas"
        files.Add "src/TableConversions.bas"
    End If

    Set GetProjectFiles = files
End Function
