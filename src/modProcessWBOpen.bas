Attribute VB_Name = "modProcessWBOpen"

'@Folder("FixLinks2UDF")

Option Explicit

'Counter to keep score of how many workbooks are open
Private mlBookCount As Long
'Counter to check how many times we've looped
Private mlTimesLooped As Long

'==============================================================================
'NOTE: either only adapt links when AddIn (file) name is identical or also allow
'      to adapt links if the AddIn (file) name has another (AddIn) extension
Private Const AllowAllAddInExtensions As Boolean = False
'==============================================================================


Private Property Get TimesLooped() As Long
    TimesLooped = mlTimesLooped
End Property

Public Property Let TimesLooped(ByVal lTimesLooped As Long)
    mlTimesLooped = lTimesLooped
End Property


'@EntryPoint
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
Public Sub ProcessNewBookOpened(ByVal wkb As Workbook)
    If wkb Is Nothing Then Exit Sub
    If wkb Is ThisWorkbook Then Exit Sub
    If wkb.IsInplace Then Exit Sub
    
    '---
    'NOTE: if also links on protected sheets should be adapted
    '      (part 1 of "protected sheets handler")
    GetProtectedWorksheets wkb
    UnprotectWorksheets
    '---
    
    Dim vAddIn As Variant
    For Each vAddIn In AddInCollection
        Dim CurrentAddIn As Workbook
        Set CurrentAddIn = vAddIn
        
        CheckAndFixLinks wkb, CurrentAddIn
        ReplaceMyFunctions wkb, CurrentAddIn
    Next
    
    '---
    'NOTE: (part 2 of "protected sheets handler")
    RestoreProtection
    '---
    
    CountBooks
End Sub

'Check for links to AddIn and fix them if they are not pointing to proper location
Private Sub CheckAndFixLinks( _
    ByVal wkb As Workbook, _
    ByVal CurrentAddIn As Workbook _
)
    
    Dim vLink As Variant
    Dim vLinks As Variant
    'Get all links
    vLinks = wkb.LinkSources(xlExcelLinks)
    'Check if we have any links, if not, exit
    If IsEmpty(vLinks) Then Exit Sub
    
    Dim CurrentAddInNamesCollection As Collection
    Set CurrentAddInNamesCollection = GetCollectionOfCurrentAddInNames(CurrentAddIn.Name)
    
    For Each vLink In vLinks
        If MeetsCriteriaToChangeLink(vLink, CurrentAddIn, CurrentAddInNamesCollection) Then
            'We've found a link to our add-in, redirect it to
            'its current location. Avoid prompts
            Application.DisplayAlerts = False
            'if a worksheet is (still) password protected runtime error 1004 is thrown
            On Error GoTo errPasswordProtected
            wkb.ChangeLink vLink, CurrentAddIn.FullName, xlLinkTypeExcelLinks
            On Error GoTo 0
            Application.DisplayAlerts = True
        End If
    Next
    Exit Sub
    
errPasswordProtected:
    If Err.Number = 1004 Then
        Err.Clear
        Resume Next
    Else
        Err.Raise Err.Number
    End If
    
End Sub

'in case to compare "only" base names
Private Function MeetsCriteriaToChangeLink( _
    ByVal vLink As Variant, _
    ByVal CurrentAddIn As Workbook, _
    ByVal CurrentAddInNamesCollection As Collection _
        ) As Boolean
    
    MeetsCriteriaToChangeLink = False
    
    'the link is already correct
    If vLink = CurrentAddIn.FullName Then Exit Function
    
    Dim vAddInName As Variant
    For Each vAddInName In CurrentAddInNamesCollection
        Dim sAddInName As String
        sAddInName = CStr(vAddInName)
        
        If vLink Like "*" & sAddInName Then
            MeetsCriteriaToChangeLink = True
            Exit Function
        End If
    Next
    
End Function

Private Function GetCollectionOfCurrentAddInNames( _
    ByVal CurrentAddInName As String _
        ) As Collection
    
    If AllowAllAddInExtensions Then
        Dim col As Collection
        Set col = GetCollectionOfCurrentAddInBaseNameWithAllAddInExtensions(CurrentAddInName)
    Else
        Set col = New Collection
        col.Add CurrentAddInName
    End If
    
    Set GetCollectionOfCurrentAddInNames = col
    
End Function

'(add a reference to the "Microsoft Scripting Runtime" library)
Private Function GetCollectionOfCurrentAddInBaseNameWithAllAddInExtensions( _
    ByVal CurrentAddInName As String _
        ) As Collection
    
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    
    Dim AddInBaseName As String
    AddInBaseName = fso.GetBaseName(CurrentAddInName)
    
    Dim col As Collection
    Set col = New Collection
    
    col.Add AddInBaseName & ".xlam"
    col.Add AddInBaseName & ".xla"
    
    Set GetCollectionOfCurrentAddInBaseNameWithAllAddInExtensions = col
    
End Function

'Ensure (relevant) functions point to `CurrentAddIn`
Private Sub ReplaceMyFunctions( _
    ByVal wkb As Workbook, _
    ByVal CurrentAddIn As Workbook _
)
    
    Dim CurrentAddInNamesCollection As Collection
    Set CurrentAddInNamesCollection = GetCollectionOfCurrentAddInNames(CurrentAddIn.Name)
    
    Dim vAddInName As Variant
    For Each vAddInName In CurrentAddInNamesCollection
        Dim AddInName As String
        AddInName = CStr(vAddInName)
        
        ReplaceMyFunctionsHandler wkb, AddInName
    Next
    
End Sub

Private Sub ReplaceMyFunctionsHandler( _
    ByVal wkb As Workbook, _
    ByVal AddInName As String _
)
    
    Dim AddInNamePlusSuffix As String
    AddInNamePlusSuffix = AddInName & "'!"
    
    Dim AddInNamePlusSuffixLength As Long
    AddInNamePlusSuffixLength = Len(AddInNamePlusSuffix)
    
    Dim ws As Worksheet
    For Each ws In wkb.Worksheets
        On Error Resume Next
        Dim rngFirstFound As Range
        Set rngFirstFound = ws.Cells.Find(What:=AddInNamePlusSuffix, LookIn:=xlFormulas, LookAt:=xlPart, _
                SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
        On Error GoTo 0
        
        If Not rngFirstFound Is Nothing Then
            Dim rngFound As Range
            Set rngFound = rngFirstFound
            
            Dim bCondition As Boolean
            bCondition = True
'            Debug.Assert False
            
            'Find all the cells containing references to the UDF
            Do
                'Replace all references to the UDF from the formula
                Dim vFormula As Variant
                vFormula = rngFound.Formula
                
                Dim lPos2 As Long
                lPos2 = InStr(vFormula, AddInNamePlusSuffix)
                Do While lPos2 > 0
                    Dim lPos1 As Long
                    lPos1 = InStrRev(vFormula, "'", InStr(lPos2, vFormula, AddInName))
                    lPos2 = lPos2 + AddInNamePlusSuffixLength
                    vFormula = Left$(vFormula, lPos1 - 1) & Right$(vFormula, Len(vFormula) - lPos2 + 1)
                    lPos2 = InStr(vFormula, AddInNamePlusSuffix)
                Loop
                
                'if a worksheet is (still) password protected runtime error 1004 is thrown
                On Error Resume Next
                If rngFound.HasArray Then 'check if the formula is part of a matrix
                    rngFound.FormulaArray = vFormula
                Else
                    rngFound.Formula = vFormula
                End If
                On Error GoTo 0
                
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
