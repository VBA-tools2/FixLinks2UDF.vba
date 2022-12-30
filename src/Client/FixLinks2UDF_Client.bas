Attribute VB_Name = "FixLinks2UDF_Client"

'@Folder("FixLinks2UDF")

Option Explicit
Option Private Module

Public Sub RegisterAddInToFixLinks2UDF()
'    '---
'    'early bound call
'    '(add reference to the "FixLinks2UDF" AddIn)
'    FixLinks2UDF.AddAddInToCollection ThisWorkbook
    '---
    'late bound call
    On Error GoTo errFixLinks2UdfNotPresent
    Application.Run "FixLinks2UDF.xlam!AddAddInToCollection", ThisWorkbook
    On Error GoTo 0
    '---
TidyUp:
    Exit Sub
    
errFixLinks2UdfNotPresent:
    Debug.Print "!!! " & ThisWorkbook.Name & ": The AddIn 'FixLinks2UDF' is blocked or isn't found."
    GoTo TidyUp
End Sub
