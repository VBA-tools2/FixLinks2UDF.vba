Attribute VB_Name = "modSheetProtection"

'@Folder("FixLinks2UDF.SheetProtection")

Option Explicit
Option Private Module

Private ProtectedWorksheetsCollection As Collection

Public Sub GetProtectedWorksheets(ByVal wkb As Workbook)
    
    Set ProtectedWorksheetsCollection = New Collection
    
    Dim ws As Worksheet
    For Each ws In wkb.Worksheets
        Dim wks As IWorksheetProtection
        '---
        'NOTE: If you have Excel 2021 or newer use this ...
        Set wks = WorksheetSetUserInterfaceOnly.Create(ws)
        '---
'        'NOTE: ... otherwise you could use this to circumvent some bugs
'        '      (which are at least present in Excel 2016)
'        Set wks = WorksheetRemoveProtection.Create(ws)
'        '---
        
        If wks.IsWorksheetProtected Then
            ProtectedWorksheetsCollection.Add wks
        End If
    Next
    
End Sub

Public Sub UnprotectWorksheets()
    Dim vWorksheet As Variant
    For Each vWorksheet In ProtectedWorksheetsCollection
        vWorksheet.Unprotect
    Next
End Sub

Public Sub RestoreProtection()
    Dim vWorksheet As Variant
    For Each vWorksheet In ProtectedWorksheetsCollection
        vWorksheet.RestoreProtection
    Next
End Sub
