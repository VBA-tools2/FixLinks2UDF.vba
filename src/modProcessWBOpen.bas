Attribute VB_Name = "modProcessWBOpen"

'@Folder("FixLinks2UDF")

Option Explicit

'Counter to keep score of how many workbooks are open
Private mlBookCount As Long
'Counter to check how many times we've looped
Private mlTimesLooped As Long


Private Property Get TimesLooped() As Long
    TimesLooped = mlTimesLooped
End Property

Public Property Let TimesLooped(ByVal lTimesLooped As Long)
    mlTimesLooped = lTimesLooped
End Property


'Checks if a new workbook has been opened (repeatedly until ActiveWorkbook is not Nothing)
Public Sub CheckIfBookOpened()
    'First, we check if the number of workbooks has changed
    If BookAdded Then
        If ActiveWorkbook Is Nothing Then
            ManageTimesLooped
        Else
            ProcessNewBookOpened ActiveWorkbook
        End If
    End If
End Sub

Private Function BookAdded() As Boolean
    If mlBookCount <> Workbooks.Count Then
        BookAdded = True
        CountBooks
    End If
End Function

Private Sub CountBooks()
    mlBookCount = Workbooks.Count
End Sub

Private Sub ManageTimesLooped()
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
End Sub

'When a new workbook is opened, this sub will be run.
Public Sub ProcessNewBookOpened(wkb As Workbook)
    If wkb Is Nothing Then Exit Sub
    If wkb Is ThisWorkbook Then Exit Sub
    If wkb.IsInplace Then Exit Sub
    CheckAndFixLinks wkb
    ReplaceMyFunctions wkb
    CountBooks
End Sub

'Check for links to AddIn and fix them if they are not pointing to proper location
Private Sub CheckAndFixLinks(wkb As Workbook)
    Dim thisWkb As Workbook
    Set thisWkb = ThisWorkbook
    
    Dim vLink As Variant
    Dim vLinks As Variant
    'Get all links
    vLinks = wkb.LinkSources(xlExcelLinks)
    'Check if we have any links, if not, exit
    If IsEmpty(vLinks) Then Exit Sub
    For Each vLink In vLinks
        If MeetsCriteriaToChangeLink(vLink, thisWkb) Then
            'We've found a link to our add-in, redirect it to
            'its current location. Avoid prompts
            Application.DisplayAlerts = False
            wkb.ChangeLink vLink, thisWkb.FullName, xlLinkTypeExcelLinks
            Application.DisplayAlerts = True
        End If
    Next
    On Error GoTo 0
End Sub

'in case to compare "only" base names
'(eventually add a reference to the "Microsoft Scripting Runtime" library)
Private Function MeetsCriteriaToChangeLink( _
    ByVal vLink As Variant, _
    ByVal thisWkb As Workbook _
        ) As Boolean
    
    MeetsCriteriaToChangeLink = False
    
    'the link is already correct
    If vLink = thisWkb.FullName Then Exit Function
    
    '---
    'if the AddIn (file) name should be identical
    If Not vLink Like "*" & thisWkb.Name Then Exit Function
'    '---
'    'if the AddIn (file) name could have another (AddIn) extension
'    '(add reference to "Microsoft Scripting Runtime" library)
'    Dim fso As New Scripting.FileSystemObject
'    Dim WkbBaseName As String
'    WkbBaseName = fso.GetBaseName(thisWkb.Name)
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
Private Sub ReplaceMyFunctions(wkb As Workbook)
    Dim ws As Worksheet
    For Each ws In wkb.Worksheets
        Dim rngFirstFound As Range
        Dim sWorkbookName As String
        Dim lWorkBookNameLength As Long
        Dim rngFound As Range
        Dim bCondition As Boolean
        
        sWorkbookName = ThisWorkbook.Name & "'!"
        lWorkBookNameLength = Len(sWorkbookName)
        On Error Resume Next
        Set rngFirstFound = ws.Cells.Find(What:=sWorkbookName, LookIn:=xlFormulas, LookAt:=xlPart, _
                SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
        On Error GoTo 0
        If Not rngFirstFound Is Nothing Then
            Set rngFound = rngFirstFound
            bCondition = True
            Debug.Assert False
            'Find all the cells containing references to the UDF
            Do
                Dim vFormula As Variant
                Dim lPos1 As Long
                Dim lPos2 As Long
                'Replace all references to the UDF from the formula
                vFormula = rngFound.Formula
                lPos2 = InStr(vFormula, sWorkbookName)
                Do While lPos2 > 0
                    lPos1 = InStrRev(vFormula, "'", InStr(lPos2, vFormula, ThisWorkbook.Name))
                    lPos2 = lPos2 + lWorkBookNameLength
                    vFormula = Left$(vFormula, lPos1 - 1) & Right$(vFormula, Len(vFormula) - lPos2 + 1)
                    lPos2 = InStr(vFormula, sWorkbookName)
                Loop
                If rngFound.HasArray Then 'check if the formula is part of a matrix
                    rngFound.FormulaArray = vFormula
                Else
                    rngFound.Formula = vFormula
                End If
                Set rngFound = ws.UsedRange.Cells.FindNext(After:=rngFound)
                If rngFound Is Nothing Then
                    bCondition = False
                ElseIf rngFound.Address = rngFirstFound.Address Then
                    bCondition = False
                End If
            Loop While bCondition
        End If
    Next
End Sub
