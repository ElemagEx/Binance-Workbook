If WScript.Arguments.Count < 1 Then
    WScript.Echo "Usage: cscript RunMacro.vbs ""<path-to-xlsm>"""
    WScript.Quit
End If

' Get arguments from the command line
excelFilePath = WScript.Arguments(0)
macroName = "CleanUpProject"

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

' Run the specified macro
objWorkbook.Application.Run macroName

' Save the workbook
objWorkbook.Save

' Close the workbook
objWorkbook.Close

' Quit the Excel application
objExcel.Quit

' Clean up the objects
Set objWorkbook = Nothing
Set objExcel = Nothing

WScript.Echo "Macro '" & macroName & "' executed successfully in '" & excelFilePath & "'"
