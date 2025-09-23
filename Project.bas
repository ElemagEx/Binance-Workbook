Attribute VB_Name = "Project"

Public Function GetProjectFiles() As collection
    Dim files As New collection

    files.Add "VBA-JSON/JsonConverter.bas"
    files.Add "src/test1.bas"

    Set GetProjectFiles = files
End Function
