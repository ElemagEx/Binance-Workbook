Attribute VB_Name = "Project"
Public Function GetProjectFiles(ByVal thirdParty As Boolean) As collection
    Dim files As New collection

    If thirdParty Then
        files.Add "VBA-JSON/JsonConverter.bas"
        files.Add "VBA-Dictionary/Dictionary.cls"
    Else
        files.Add "src/KeysAndSecrets.bas"
    End If

    Set GetProjectFiles = files
End Function
