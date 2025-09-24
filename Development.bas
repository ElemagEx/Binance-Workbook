Private Const PROJECT_FILE = "project.bas"

Private Sub CleanUpAllData(ByVal removeLedgerSheets As Boolean)

End Sub

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

Private Sub HandleFile(ByVal dir As String, ByVal file As String, ByVal export As Boolean, ByVal import As Boolean, ByVal override As Boolean)
    Dim name As String
    name = GetNameOfFile(file)

    Dim path As String
    path = dir & "/" & file

    If InStr(1, Application.OperatingSystem, "Windows", vbTextCompare) > 0 Then
        path = Replace(path, "/", "\")
    End If

    Dim component As VBComponent
    Set component = GetComponent(name)
    If component Is Nothing Then
        If import Then
            ThisWorkbook.VBProject.VBComponents.import path
        End If
    Else
        If export Then
            component.export path
        End If
        If override Then
            ThisWorkbook.VBProject.VBComponents.Remove component
        End If
        If import And override Then
            ThisWorkbook.VBProject.VBComponents.import path
        End If
    End If
End Sub

Private Sub HandleProjectFiles(ByVal dir As String, ByVal export As Boolean, ByVal import As Boolean, ByVal override As Boolean)
    Dim file As Variant
    Dim files As collection
    
    Set files = GetProjectFiles(True)
    
    For Each file In files
        HandleFile dir, file, False, import, override
    Next file

    Set files = GetProjectFiles(False)
   
    For Each file In files
        HandleFile dir, file, export, import, override
    Next file
End Sub

Public Sub ImportProjectFiles()
    Dim dir As String
    dir = ActiveWorkbook.path

    HandleFile dir, PROJECT_FILE, False, True, True

    HandleProjectFiles dir, False, True, True
End Sub

Public Sub ExportProjectFiles()
    If GetComponent(GetNameOfFile(PROJECT_FILE)) Is Nothing Then
        Exit Sub
    End If

    Dim dir As String
    dir = ActiveWorkbook.path

    HandleProjectFiles dir, True, False, False

    HandleFile dir, PROJECT_FILE, True, False, False
End Sub

Public Sub CleanUpProject()
    If GetComponent(GetNameOfFile(PROJECT_FILE)) Is Nothing Then
        Exit Sub
    End If

    CleanUpAllData True

    Dim dir As String
    dir = ActiveWorkbook.path

    HandleProjectFiles dir, False, False, True

    HandleFile dir, PROJECT_FILE, False, False, True
End Sub

Public Function IsProjectCleanedUp() As Boolean
    IsProjectCleanedUp = (GetComponent(GetNameOfFile(PROJECT_FILE)) Is Nothing)
End Function
