Attribute VB_Name = "modProcessWBOpen"

Option Explicit

'Counter to keep score of how many workbooks are open
Dim mlBookCount As Long

'Counter to check how many times we've looped
Private mlTimesLooped As Long


Sub ProcessNewBookOpened(oBk As Workbook)
'------------------------------------------------------------------------------
' Procedure : ProcessNewBookOpened
' Company   : JKP Application Development Services (c) 2005
' Author    : Jan Karel Pieterse
' Created   : 2-6-2008
' Purpose   : When a new workbook is opened, this sub will be run.
' Called from: clsAppEvents.App_Workbook_Open and ThisWorkbook.Workbook_Open
'------------------------------------------------------------------------------
'Sometimes OBk is nothing?
    If oBk Is Nothing Then Exit Sub
    If oBk Is ThisWorkbook Then Exit Sub
    If oBk.IsInplace Then Exit Sub
    CheckAndFixLinks oBk
    ReplaceMyFunctions oBk
    CountBooks
End Sub

Sub CountBooks()
    mlBookCount = Workbooks.Count
End Sub

Function BookAdded() As Boolean
    If mlBookCount <> Workbooks.Count Then
        BookAdded = True
        CountBooks
    End If
End Function

Sub CheckIfBookOpened()
'------------------------------------------------------------------------------
' Procedure : CheckIfBookOpened
' Company   : JKP Application Development Services (c) 2005
' Author    : Jan Karel Pieterse
' Created   : 6-6-2008
' Purpose   : Checks if a new workbook has been opened (repeatedly until activeworkbook is not nothing)
'------------------------------------------------------------------------------
    'First, we check if the number of workbooks has changed
    If BookAdded Then
        If ActiveWorkbook Is Nothing Then
            mlBookCount = 0
            'Increment the loop counter
            TimesLooped = TimesLooped + 1
            'May be needed if Excel is opened from Internet explorer
            Application.Visible = True
            If TimesLooped < 20 Then
                'We've not yet done this 20 times, schedule another in 1 sec
                Application.OnTime Now + TimeValue("00:00:01"), "CheckIfBookOpened"
            Else
                'We've done this 20 times, do not schedule another
                'and reset the counter
                TimesLooped = 0
            End If
        Else
            ProcessNewBookOpened ActiveWorkbook
        End If
    End If
End Sub

Public Property Get TimesLooped() As Long
    TimesLooped = mlTimesLooped
End Property

Public Property Let TimesLooped(ByVal lTimesLooped As Long)
    mlTimesLooped = lTimesLooped
End Property

Sub CheckAndFixLinks(oBook As Workbook)
'------------------------------------------------------------------------------
' Procedure : CheckAndFixLinks Created by Jan Karel Pieterse
' Company   : JKP Application Development Services (c) 2008
' Author    : Jan Karel Pieterse
' Created   : 2-6-2008
' Purpose   : Checks for links to addin and fixes them
'             if they are not pointing to proper location
'------------------------------------------------------------------------------
    Dim vLink As Variant
    Dim vLinks As Variant
    'Get all links
    vLinks = oBook.LinkSources(xlExcelLinks)
    'Check if we have any links, if not, exit
    If IsEmpty(vLinks) Then Exit Sub
    For Each vLink In vLinks
        If vLink Like "*" & ThisWorkbook.Name Then
            'We've found a link to our add-in, redirect it to
            'its current location. Avoid prompts
            Application.DisplayAlerts = False
            oBook.ChangeLink vLink, ThisWorkbook.FullName, xlLinkTypeExcelLinks
            Application.DisplayAlerts = True
        End If
    Next
    On Error GoTo 0
End Sub

Private Sub ReplaceMyFunctions(oBk As Workbook)
'------------------------------------------------------------------------------
' Procedure : ReplaceMyFunctions Created by Jan Karel Pieterse
' Company   : JKP Application Development Services (c) 2008
' Author    : Jan Karel Pieterse
' Created   : 2-6-2008
' Purpose   : Ensures My functions point to this addin
'------------------------------------------------------------------------------
    Dim oSh As Worksheet
    Dim oFirstFound As Range
    Dim oFound As Range
    
    On Error Resume Next
    'Search through all sheets looking for the UDF "UDFDemo("
    For Each oSh In oBk.Worksheets
        Set oFirstFound = _
            oSh.UsedRange.Cells.Find(What:="UDFDemo(", After:=oSh.UsedRange.Cells(1, 1), _
                    LookIn:=xlFormulas, LookAt:=xlPart, _
                    SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
        If Not oFirstFound Is Nothing Then
            'Found one, change the formula (prepend with path to me)
            'We assume the function is on its own, NOT nested inside another!!!
            oFirstFound.Formula = "='" & ThisWorkbook.FullName & "'!" & _
                    Right(oFirstFound.Formula, _
                            Len(oFirstFound.Formula) - _
                            InStr(oFirstFound.Formula, "My(") + 1)
            Set oFound = oFirstFound
            Do
                Set oFound = _
                        oSh.UsedRange.Cells.Find(What:="UDFDemo(", After:=oFound, LookIn:=xlFormulas, _
                        LookAt:=xlPart, SearchOrder:=xlByRows, _
                        SearchDirection:=xlNext, MatchCase:=False)
                If Not oFound Is Nothing Then
                    'We assume the function is on its own, NOT nested inside another!!!
                    oFound.Formula = "='" & ThisWorkbook.FullName & "'!" & _
                            Right(oFound.Formula, Len(oFound.Formula) - _
                                    InStr(oFound.Formula, "My(") + 1)
                End If
            Loop Until oFound Is Nothing Or oFound.Address = oFirstFound.Address
        End If
    Next
End Sub
