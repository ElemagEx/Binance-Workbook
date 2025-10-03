Attribute VB_Name = "Development"
Option Explicit

Private Const PROJECT_FILE = "project.bas"
Private Const THIS_FILE = "development.bas"

Public Sub Project_ImportFiles()
    Dim dir As String
    dir = ActiveWorkbook.path

    xHandleFile dir, PROJECT_FILE, False, True, True

    xHandleProjectFiles dir, False, True, True
End Sub

Public Sub Project_ExportFiles()
    If xIsProjectInactive() Then
        Exit Sub
    End If

    Dim dir As String
    dir = ActiveWorkbook.path

    xHandleProjectFiles dir, True, False, False

    xHandleFile dir, PROJECT_FILE, True, False, False
    
    xHandleFile dir, THIS_FILE, True, False, False
End Sub

Public Sub Project_CleanUpFiles()
    If xIsProjectInactive() Then
        Exit Sub
    End If

    Dim dir As String
    dir = ActiveWorkbook.path

    xHandleProjectFiles dir, False, True, False

    xHandleFile dir, PROJECT_FILE, False, True, False
End Sub

Public Sub Project_CleanUpData()
    If xIsProjectInactive() Then
        Exit Sub
    End If
    
    xCleanUpDate
End Sub

Public Sub Project_CloseAllLedgers()
    If xIsProjectInactive() Then
        Exit Sub
    End If
    
    xCloseAllLedgers
End Sub

Public Function Project_TryCleanUpData() As Boolean
    Project_TryCleanUpData = False
    
    If xIsProjectInactive() Then
        Exit Function
    End If
    
    xCleanUpDate

    Project_TryCleanUpData = True
End Function

Public Function Project_IsDataCleanedUp()
    Project_IsDataCleanedUp = False
    
    If xIsProjectInactive() Then
        Exit Function
    End If
    
    Project_IsDataCleanedUp = xIsDateCleanup()
End Function

Private Sub xCloseAllLedgers()
    BinanceLedgers.RemoveAll
    TableLastUpdate.ClearData False
End Sub

Private Sub xCleanUpDate()
    TableConversions.CleanUpData
    TableDustLog.CleanUpData

    TableCurrencies.CleanUpData
    
    TableLastUpdate.CleanUpData
    TableUserInfo.CleanUpData
    TableWallet.CleanUpData
    
    BinanceLedgers.RemoveAll
End Sub

Private Function xIsDateCleanup() As Boolean
    xIsDateCleanup = _
        TableConversions.IsDataCleanedUp And _
        TableDustLog.IsDataCleanedUp And _
        TableCurrencies.IsDataCleanedUp And _
        TableLastUpdate.IsDataCleanedUp And _
        TableUserInfo.IsDataCleanedUp And _
        TableWallet.IsDataCleanedUp And _
        Not BinanceLedgers.HasExistingLedgers
End Function

Public Function xIsProjectInactive() As Boolean
    xIsProjectInactive = (xGetComponent(xGetNameOfFile(PROJECT_FILE)) Is Nothing)
End Function

Private Function xGetNameOfFile(ByVal file As String) As String
    Dim name As String
    name = Mid(file, InStrRev(file, "/") + 1)
    name = Left(name, InStr(1, name, ".") - 1)
    xGetNameOfFile = name
End Function

Private Function xGetComponent(ByVal name As String) As VBComponent
    Set xGetComponent = Nothing
    
    Dim component As VBComponent
    For Each component In ThisWorkbook.VBProject.VBComponents
        If StrComp(component.name, name, vbTextCompare) = 0 Then
            Set xGetComponent = component
            Exit For
        End If
    Next component
End Function

Private Sub xHandleFile(ByVal dir As String, ByVal file As String, ByVal export As Boolean, ByVal remove As Boolean, ByVal import As Boolean)
    Dim name As String
    name = xGetNameOfFile(file)

    Dim path As String
    path = dir & "/" & file

    If InStr(1, Application.OperatingSystem, "Windows", vbTextCompare) > 0 Then
        path = Replace(path, "/", "\")
    End If

    Dim component As VBComponent
    Set component = xGetComponent(name)
    If component Is Nothing Then
        If import Then
            ThisWorkbook.VBProject.VBComponents.import path
        End If
    Else
        If export Then
            component.export path
        End If
        If remove Then
            ThisWorkbook.VBProject.VBComponents.remove component
        End If
        If import Then
            ThisWorkbook.VBProject.VBComponents.import path
        End If
    End If
End Sub

Private Sub xHandleProjectFiles(ByVal dir As String, ByVal export As Boolean, ByVal remove As Boolean, ByVal import As Boolean)
    Dim file As Variant
    Dim files As Collection
    
    Set files = GetProjectFiles(True)
    
    For Each file In files
        xHandleFile dir, file, False, remove, import
    Next file

    Set files = GetProjectFiles(False)

    For Each file In files
        xHandleFile dir, file, export, remove, import
    Next file
End Sub
