Attribute VB_Name = "Project"
Public Function GetProjectFiles(ByVal thirdParty As Boolean) As collection
    Dim files As New collection

    If thirdParty Then
        files.Add "VBA-Dictionary/Dictionary.cls"
        files.Add "VBA-JSON/JsonConverter.bas"
        files.Add "VBA-Web/WebHelpers.bas"
        files.Add "VBA-Web/WebClient.cls"
        files.Add "VBA-Web/WebRequest.cls"
        files.Add "VBA-Web/WebResponse.cls"
        files.Add "VBA-Web/IWebAuthenticator.cls"
    Else
        files.Add "src/KeysAndSecrets.bas"
        files.Add "src/BinanceApi.bas"
        files.Add "src/ButtonsEvents.bas"
    End If

    Set GetProjectFiles = files
End Function
