Attribute VB_Name = "modAddInHandler"

'@Folder("FixLinks2UDF")

Option Explicit

Public AddInCollection As New Collection

'==============================================================================
'NOTE: in case this AddIn should check for old/wrong links itself/standalone,
'      i.e. without having a `FixLinks2UDF.xlam` (AddIn)
Public Const StandAlone As Boolean = False
'==============================================================================

Private Enum eAddInsHandlerError
    [_First] = vbObjectError + 1
    ErrBuiltinDocumentPropertiesTitleIsEmpty = [_First]
    [_Last] = ErrBuiltinDocumentPropertiesTitleIsEmpty
End Enum

Public Sub EnsureBuiltinDocumentPropertiesTitleIsNotEmpty()
    If IsBuiltinDocumentPropertiesTitleEmpty Then RaiseErrorBuiltinDocumentPropertiesTitleIsEmpty
End Sub

Private Function IsBuiltinDocumentPropertiesTitleEmpty() As Boolean
    IsBuiltinDocumentPropertiesTitleEmpty = _
            (ThisWorkbook.BuiltinDocumentProperties("Title") = vbNullString)
End Function

Public Sub AddAddInToCollection(ByVal wkb As Workbook)
    
    If Not AreConditionsMeetToAddAddInToCollection Then Exit Sub
    
    'prevent an error in case *all* loaded/active AddIns should be checked *and*
    'AddIns have registered, i.e. the AddIn name is already in the collection
    On Error Resume Next
    AddInCollection.Add wkb, wkb.Name
    On Error GoTo 0
    
End Sub

Private Function AreConditionsMeetToAddAddInToCollection() As Boolean
    
    AreConditionsMeetToAddAddInToCollection = False
    
    If StandAlone Then
        'fine
    Else
        Dim AddInTitle As String
        AddInTitle = ThisWorkbook.BuiltinDocumentProperties("Title")
        
        If AddIns(AddInTitle).Installed Then
            'fine
        Else
            Exit Function
        End If
    End If
    
    AreConditionsMeetToAddAddInToCollection = True
    
End Function

Public Sub RemoveAddInFromCollection(ByVal wkb As Workbook)
    'ignore 'wkb's that aren't in the collection
    On Error Resume Next
    AddInCollection.Remove wkb.Name
    On Error GoTo 0
End Sub

Private Sub RaiseErrorBuiltinDocumentPropertiesTitleIsEmpty()
    Err.Raise _
        Number:=eAddInsHandlerError.ErrBuiltinDocumentPropertiesTitleIsEmpty, _
        Description:= _
                "The 'BuiltinDocumentProperties(""Title"")' is empty." & vbCrLf & _
                "(It is expected to be 'FixLinks2UDF'.)"
End Sub
