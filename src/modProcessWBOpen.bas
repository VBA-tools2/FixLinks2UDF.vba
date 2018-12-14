Attribute VB_Name = "modProcessWBOpen"

Option Explicit

'Counter to keep score of how many workbooks are open
Private mlBookCount As Long
'Counter to check how many times we've looped
Private mlTimesLooped As Long


Public Sub ProcessNewBookOpened(oBk As Workbook)
'------------------------------------------------------------------------------
' Procedure : ProcessNewBookOpened
' Company   : JKP Application Development Services (c) 2005
' Author    : Jan Karel Pieterse
' Created   : 2-6-2008
' Purpose   : When a new workbook is opened, this sub will be run.
' Called from: clsAppEvents.App_Workbook_Open and ThisWorkbook.Workbook_Open
'------------------------------------------------------------------------------
'Sometimes oBk is nothing?
    If oBk Is Nothing Then Exit Sub
    If oBk Is ThisWorkbook Then Exit Sub
    If oBk.IsInplace Then Exit Sub
    CheckAndFixLinks oBk
    ReplaceMyFunctions oBk
    CountBooks
End Sub

Private Sub CountBooks()
    mlBookCount = Workbooks.Count
End Sub

Private Function BookAdded() As Boolean
    If mlBookCount <> Workbooks.Count Then
        BookAdded = True
        CountBooks
    End If
End Function

Public Sub CheckIfBookOpened()
'------------------------------------------------------------------------------
' Procedure : CheckIfBookOpened
' Company   : JKP Application Development Services (c) 2005
' Author    : Jan Karel Pieterse
' Created   : 6-6-2008
' Purpose   : Checks if a new workbook has been opened (repeatedly until ActiveWorkbook is not Nothing)
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

Private Property Get TimesLooped() As Long
    TimesLooped = mlTimesLooped
End Property

Public Property Let TimesLooped(ByVal lTimesLooped As Long)
    mlTimesLooped = lTimesLooped
End Property

Private Sub CheckAndFixLinks(oBook As Workbook)
'------------------------------------------------------------------------------
' Procedure : CheckAndFixLinks Created by Jan Karel Pieterse
' Company   : JKP Application Development Services (c) 2008
' Author    : Jan Karel Pieterse
' Created   : 2-6-2008
' Purpose   : Checks for links to AddIn and fixes them
'             if they are not pointing to proper location
'------------------------------------------------------------------------------
    Dim vLink As Variant
    Dim vLinks As Variant
    'Get all links
    vLinks = oBook.LinkSources(xlExcelLinks)
    'Check if we have any links, if not, exit
    If IsEmpty(vLinks) Then Exit Sub
    For Each vLink In vLinks
        If vLink Like "*" & ThisWorkbook.Name And vLink <> ThisWorkbook.FullName Then
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
'             and Improved by Jorge Belenguer Faguas
' Company   : JKP Application Development Services (c) 2008
' Author    : Jan Karel Pieterse
' Created   : 2-6-2008
' Modified  : 1-21-2009 by Jorge Belenguer Faguas
' Purpose   : Ensures My functions point to this AddIn
'------------------------------------------------------------------------------
    Dim oSh As Worksheet
    For Each oSh In oBk.Worksheets
        Dim oFirstFound As Range
        Dim lWorkbookName As String
        Dim lWorkBookNameLength As Long
        Dim oFound As Range
        Dim lCondition As Boolean
        
        lWorkbookName = ThisWorkbook.Name & "'!"
        lWorkBookNameLength = Len(lWorkbookName)
        On Error Resume Next
        Set oFirstFound = oSh.Cells.Find(What:=lWorkbookName, LookIn:=xlFormulas, LookAt:=xlPart, _
                SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
        On Error GoTo 0
        If Not oFirstFound Is Nothing Then
            Set oFound = oFirstFound
            lCondition = True
            Debug.Assert False
            'Find all the cells containing references to the UDF
            Do
                Dim vFormula As Variant
                Dim lPos1 As Long
                Dim lPos2 As Long
                'Replace all references to the UDF from the formula
                vFormula = oFound.Formula
                lPos2 = InStr(vFormula, lWorkbookName)
                Do While lPos2 > 0
                    lPos1 = InStrRev(vFormula, "'", InStr(lPos2, vFormula, ThisWorkbook.Name))
                    lPos2 = lPos2 + lWorkBookNameLength
                    vFormula = Left(vFormula, lPos1 - 1) & Right(vFormula, Len(vFormula) - lPos2 + 1)
                    lPos2 = InStr(vFormula, lWorkbookName)
                Loop
                If oFound.HasArray Then 'check if the formula is part of a matrix
                    oFound.FormulaArray = vFormula
                Else
                    oFound.Formula = vFormula
                End If
                Set oFound = oSh.UsedRange.Cells.FindNext(After:=oFound)
                If (oFound Is Nothing) Then
                    lCondition = False
                ElseIf (oFound.Address = oFirstFound.Address) Then
                    lCondition = False
                End If
            Loop While lCondition
        End If
    Next oSh
End Sub
