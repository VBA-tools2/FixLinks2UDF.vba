Attribute VB_Name = "modProcessWBOpen"

'@Folder("FixLinks2UDF")

Option Explicit

'Counter to keep score of how many workbooks are open
Private mlBookCount As Long
'Counter to check how many times we've looped
Private mlTimesLooped As Long


'When a new workbook is opened, this sub will be run.
Public Sub ProcessNewBookOpened(oBk As Workbook)
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

'Checks if a new workbook has been opened (repeatedly until ActiveWorkbook is not Nothing)
Public Sub CheckIfBookOpened()
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

'Check for links to AddIn and fix them if they are not pointing to proper location
Private Sub CheckAndFixLinks(oBook As Workbook)
    Dim wkb As Workbook
    Set wkb = ThisWorkbook
    
    Dim vLink As Variant
    Dim vLinks As Variant
    'Get all links
    vLinks = oBook.LinkSources(xlExcelLinks)
    'Check if we have any links, if not, exit
    If IsEmpty(vLinks) Then Exit Sub
    For Each vLink In vLinks
        If MeetsCriteriaToChangeLink(vLink, wkb) Then
            'We've found a link to our add-in, redirect it to
            'its current location. Avoid prompts
            Application.DisplayAlerts = False
            oBook.ChangeLink vLink, wkb.FullName, xlLinkTypeExcelLinks
            Application.DisplayAlerts = True
        End If
    Next
    On Error GoTo 0
End Sub

'in case to compare "only" base names
'(eventually add a reference to the "Microsoft Scripting Runtime" library)
Private Function MeetsCriteriaToChangeLink( _
    ByVal vLink As Variant, _
    ByVal wkb As Workbook _
        ) As Boolean
    
    MeetsCriteriaToChangeLink = False
    
    'the link is already correct
    If vLink = wkb.FullName Then Exit Function
    
    '---
    'if the AddIn (file) name should be identical
    If Not vLink Like "*" & wkb.Name Then Exit Function
'    '---
'    'if the AddIn (file) name could have another (AddIn) extension
'    '(add reference to "Microsoft Scripting Runtime" library)
'    Dim fso As New Scripting.FileSystemObject
'    Dim WkbBaseName As String
'    WkbBaseName = fso.GetBaseName(wkb.Name)
'
'    If vLink Like "*" & WkbBaseName & ".xlam" Then
'        'fine
'    ElseIf vLink Like "*" & WkbBaseName & ".xla" Then
'        'fine
'    Else
'        Exit Function
'    End If
'    '---
    
    MeetsCriteriaToChangeLink = True
    
End Function

'Ensure (relevant) functions point to this AddIn
Private Sub ReplaceMyFunctions(oBk As Workbook)
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
                    vFormula = Left$(vFormula, lPos1 - 1) & Right$(vFormula, Len(vFormula) - lPos2 + 1)
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
