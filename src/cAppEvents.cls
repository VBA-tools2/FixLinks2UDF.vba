VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cAppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("FixLinks2UDF")

'------------------------------------------------------------------------------
' Module    : cAppEvents
' Company   : JKP Application Development Services (c) 2008
' Author    : Jan Karel Pieterse
' Created   : 2-6-2008
' Purpose   : Handles Excel Application events
' URL       : <https://jkp-ads.com/Articles/FixLinks2UDF.asp>
'------------------------------------------------------------------------------
Option Explicit

'This object variable will hold the object who's events we want to respond to
'Note the "WithEvents" keyword, which is what we need to tell VBA it is an object
'with events.
'@MemberAttribute VB_VarHelpID, -1
Public WithEvents App As Application
Attribute App.VB_VarHelpID = -1

Private Sub App_WorkbookAddinUninstall(ByVal wkb As Workbook)
    RemoveAddInFromCollection wkb
End Sub

Private Sub App_WorkbookOpen(ByVal wkb As Workbook)
    '---
    'NOTE: use this when AddIns should "register" to be checked for old/wrong links.
    ProcessNewBookOpened wkb
    '---
'    'NOTE: use this if *all* loaded/active AddIns should be checked for old/wrong links
'    If wkb.IsAddin Then
'        AddAddInToCollection wkb
'    Else
'        ProcessNewBookOpened wkb
'    End If
'    '---
End Sub

Private Sub Class_Terminate()
    Set App = Nothing
End Sub
