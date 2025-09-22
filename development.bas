Attribute VB_Name = "Development"
Private Const PROJECT_FILE = "project.bas"

Private Function GetNameOfFile(ByVal file As String) As String
    Dim name As String
    name = Mid(file, InStrRev(file, "/") + 1)
    name = Left(name, InStr(1, name, ".") - 1)
    GetNameOfFile = name
End Function

Private Function GetComponent(ByVal name As String) As VBComponent
    Set GetComponent = Nothing
    
    Dim component As VBComponent
    For Each component In ThisWorkbook.VBProject.VBComponents
        If StrComp(component.name, name, vbTextCompare) = 0 Then
            Set GetComponent = component
            Exit For
        End If
    Next component
End Function

Private Sub RemoveComponent(ByVal name As String, Optional ByVal path As String = "")
    Dim component As VBComponent
    Set component = GetComponent(name)
    If Not component Is Nothing Then
        If path <> "" Then
            component.export path
        End If
        ThisWorkbook.VBProject.VBComponents.Remove component
    End If
End Sub

Private Sub HandleFile(ByVal dir As String, ByVal file As String, ByVal export As Boolean, ByVal import As Boolean)
    Dim name As String
    name = GetNameOfFile(file)

    Dim path As String
    path = dir & "/" & file

    If InStr(1, Application.OperatingSystem, "Windows", vbTextCompare) > 0 Then
        path = Replace(path, "/", "\")
    End If

    RemoveComponent name, IIf(export, path, "")

    If import Then
        ThisWorkbook.VBProject.VBComponents.import path
    End If
End Sub

Private Sub HandleProjectFiles(ByVal dir As String, ByVal export As Boolean, ByVal import As Boolean)
    Dim files As collection
    Set files = GetProjectFiles()
    
    Dim file As Variant
    For Each file In files
        HandleFile dir, file, export, import
    Next file
End Sub

Public Sub ImportProjectFiles()
    Dim dir As String
    dir = ActiveWorkbook.path

    HandleFile dir, PROJECT_FILE, False, True

    HandleProjectFiles dir, False, True
End Sub

Public Sub ExportProjectFiles()
    If GetComponent(GetNameOfFile(PROJECT_FILE)) Is Nothing Then
        Exit Sub
    End If

    Dim dir As String
    dir = ActiveWorkbook.path

    HandleProjectFiles dir, True, False

    HandleFile dir, PROJECT_FILE, True, False
End Sub

Public Sub CleanUpProjectFiles()
    If GetComponent(GetNameOfFile(PROJECT_FILE)) Is Nothing Then
        Exit Sub
    End If

    Dim dir As String
    dir = ActiveWorkbook.path

    HandleProjectFiles dir, False, False
    HandleFile dir, PROJECT_FILE, False, False
End Sub
