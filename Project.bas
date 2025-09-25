Attribute VB_Name = "Project"
Public Function GetProjectFiles(ByVal thirdParty As Boolean) As collection
    Dim files As New collection

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
        files.Add "src/BinanceApi1.bas"
        files.Add "src/ButtonsEvents.bas"
        files.Add "src/TableUserInfo.bas"
    End If

    Set GetProjectFiles = files
End Function
