Attribute VB_Name = "modDocumentProperties"

'@Folder("FixLinks2UDF")

Option Explicit

'@EntryPoint
Private Sub SetDocumentProperties()
    
    Dim TodayDate As String
    TodayDate = Format$(Date, "yyyy-mm-dd")
    
    With ThisWorkbook
        'also the name as shown in the Excel's AddIn list
        .BuiltinDocumentProperties("Title") = "FixLinks2UDF"
        .BuiltinDocumentProperties("Author") = "Jan Karel Pieterse;Stefan Pinnow"
        '6 lines can be seen in Excel's AddIn list without scrolling
        .BuiltinDocumentProperties("Comments") = _
                "Fixing Links To UDFs in AddIns" & vbCrLf & _
                "<https://jkp-ads.com/Articles/FixLinks2UDF.asp>" & vbCrLf & _
                "Build: " & TodayDate
        .Save
    End With
    
End Sub
