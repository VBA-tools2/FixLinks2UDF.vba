VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IWorksheetProtection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("FixLinks2UDF.SheetProtection")
'@ModuleDescription("Interface for Worksheet protection.")
'@Interface

Option Explicit

Public Enum eWorksheetProtectionError
    [_First] = vbObjectError + 1
    ErrIsNothing = [_First]
    [_Last] = ErrIsNothing
End Enum

'@Description("Returns, if the Worksheet is (password) protected")
Public Function IsWorksheetProtected() As Boolean
End Function

'@Description("Removes the Worksheets (password) protection")
Public Sub Unprotect()
End Sub

'@Description("Restores the Worksheets (password) protection")
Public Sub RestoreProtection()
End Sub
