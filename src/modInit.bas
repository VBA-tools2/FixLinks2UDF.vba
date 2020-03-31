Attribute VB_Name = "modInit"

'@Folder("FixLinks2UDF")

Option Explicit

'Create a module level object variable that will keep the instance of the
'event listener in memory (and hence alive)
Private moAppEventHandler As cAppEvents


Public Sub InitApp()
    'Create a new instance of cAppEvents class
    Set moAppEventHandler = New cAppEvents
    With moAppEventHandler
        'Tell it to listen to Excel's events
        Set .App = Application
    End With
End Sub
