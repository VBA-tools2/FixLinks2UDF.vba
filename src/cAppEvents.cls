VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cAppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'------------------------------------------------------------------------------
' Module    : cAppEvents
' Company   : JKP Application Development Services (c) 2008
' Author    : Jan Karel Pieterse
' Created   : 2-6-2008
' Purpose   : Handles Excel Application events
' URL       : <http://www.jkp-ads.com/Articles/FixLinks2UDF.asp>
'------------------------------------------------------------------------------
Option Explicit

'This object variable will hold the object who's events we want to respond to
'Note the "WithEvents" keyword, which is what we need to tell VBA it is an object
'with events.
Public WithEvents App As Application
Attribute App.VB_VarHelpID = -1

Private Sub App_WorkbookOpen(ByVal wb As Workbook)
'------------------------------------------------------------------------------
' Procedure : App_WorkbookOpen Created by Jan Karel Pieterse
' Company   : JKP Application Development Services (c) 2008
' Author    : Jan Karel Pieterse
' Created   : 2-6-2008
' Purpose   : Runs code to redirect UDF's to this workbook
'------------------------------------------------------------------------------
    ProcessNewBookOpened wb
End Sub

Private Sub Class_Terminate()
    Set App = Nothing
End Sub
