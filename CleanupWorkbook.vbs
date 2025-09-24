If WScript.Arguments.Count < 1 Then
    WScript.Echo "Usage: cscript RunMacro.vbs ""<path-to-xlsm>"""
    WScript.Quit(0)
End If

' Get arguments from the command line
excelFilePath = WScript.Arguments(0)

' Create a FileSystemObject to help with path manipulation
Set fso = CreateObject("Scripting.FileSystemObject")

' Get the directory where the script is located
scriptDirectory = fso.GetParentFolderName(WScript.ScriptFullName)

' BuildPath intelligently adds the backslash separator if needed
excelFullPath = fso.BuildPath(scriptDirectory, excelFilePath)

' Create an Excel application object
Set objExcel = CreateObject("Excel.Application")

' Make Excel invisible to the user
objExcel.Visible = False

' Disable alerts (like "Save changes?" prompts)
objExcel.DisplayAlerts = False

' Open the specified Excel workbook
Set objWorkbook = objExcel.Workbooks.Open(excelFullPath)

' Check for clean-up
isProjectCleanedUp = objWorkbook.Application.Run("IsProjectCleanedUp")

if isProjectCleanedUp = False Then
    objWorkbook.Application.Run "CleanUpProject"
    objWorkbook.Save
end if

objWorkbook.Close

objExcel.Quit

Set objWorkbook = Nothing
Set objExcel = Nothing

WScript.Echo "Macro '" & macroName & "' executed successfully in '" & excelFilePath & "'"

WScript.Quit(1)
