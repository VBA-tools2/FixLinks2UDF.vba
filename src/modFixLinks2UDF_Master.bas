Attribute VB_Name = "modFixLinks2UDF_Master"

'@Folder("FixLinks2UDF")

'------------------------------------------------------------------------------
' Module    : modFixLinks2UDF_Master
' Company   : JKP Application Development Services (c) 2005
' Author    : Jan Karel Pieterse
' Created   : 02-06-2008
' Purpose   : Workbook event code
' URL       : <https://jkp-ads.com/Articles/FixLinks2UDF.asp>
'------------------------------------------------------------------------------

Option Explicit
Option Private Module

Private mdNextTime As Double
Private mbUninstall As Boolean

Public Sub AddinUninstallHandler()
    mbUninstall = True
End Sub

Public Sub BeforeCloseHandler()
    If Not mbUninstall Then
        mdNextTime = Now
        '(to avoid a runtime error 1004 if a file is in "protected view",
        ' e.g. if a file from the internet is opened the first time)
        On Error Resume Next
        Application.OnTime mdNextTime, "InitApp"
        On Error GoTo 0
    End If
End Sub

Public Sub DeactivateHandler()
    On Error Resume Next
    Application.OnTime mdNextTime, "InitApp", , False
    On Error GoTo 0
End Sub

Public Sub OpenHandler()
    
    If Not StandAlone Then
        EnsureBuiltinDocumentPropertiesTitleIsNotEmpty
    End If
    
    'Initialize the application
    modInit.InitApp
    
    modProcessWBOpen.TimesLooped = 0
    
    'Schedule macro to run after initialization of Excel has fully been done.
    'Sometimes, the AddIn hasn't fully been initialized and the
    'workbook we want checked is opened BEFORE we have fully initialized the
    'AddIn.
    'This may happen when one double clicks a file in explorer
    Application.OnTime VBA.Now() + VBA.TimeValue("00:00:02"), "CheckIfBookOpened"
    
    If StandAlone Then
        AddAddInToCollection ThisWorkbook
    End If
    
End Sub
