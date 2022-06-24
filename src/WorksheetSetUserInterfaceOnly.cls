VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorksheetSetUserInterfaceOnly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "Rubberduck" ,"Predeclared Class Module"

'@Folder("FixLinks2UDF.SheetProtection")
'@ModuleDescription("Class to store essential protection properties of a worksheet.")
'@PredeclaredId

Option Explicit

Implements IWorksheetProtection

Private Type TWorksheetProtection
    Worksheet As Worksheet
    IsWorksheetProtected As Boolean
    
    UserInterfaceOnly As Boolean
End Type
Private This As TWorksheetProtection

'@Description("Factory method to create the 'Protection' object (itself)")
Public Function Create(ByVal Value As Worksheet) As IWorksheetProtection
    With New WorksheetSetUserInterfaceOnly
        Set .Worksheet = Value
        .CheckIfWorksheetIsProtected
        .GetProtectionStatus
        Set Create = .Self
    End With
End Function

'@Description("Sets the 'Protection' object (itself)")
Friend Property Get Self() As IWorksheetProtection
    Set Self = Me
End Property

'@Description("Sets the 'Worksheet' object (itself)")
Friend Property Set Worksheet(ByVal Value As Worksheet)
    If Value Is Nothing Then RaiseErrorIsNothing
    
    Set This.Worksheet = Value
End Property

Private Function IWorksheetProtection_IsWorksheetProtected() As Boolean
    IWorksheetProtection_IsWorksheetProtected = This.IsWorksheetProtected
End Function

Private Sub IWorksheetProtection_Unprotect()
    'current implementation only tries without a password
    On Error Resume Next
    This.Worksheet.Protect Password:=vbNullString, UserInterfaceOnly:=True
    On Error GoTo 0
End Sub

Private Sub IWorksheetProtection_RestoreProtection()
    With This
        .Worksheet.Protect UserInterfaceOnly:=.UserInterfaceOnly
    End With
End Sub

'==============================================================================
Friend Sub CheckIfWorksheetIsProtected()
    This.IsWorksheetProtected = True
    
    With This.Worksheet
        If .ProtectContents Then Exit Sub
        If .ProtectDrawingObjects Then Exit Sub
        If .ProtectScenarios Then Exit Sub
    End With
    
    This.IsWorksheetProtected = False
End Sub

Friend Sub GetProtectionStatus()
    If Not This.IsWorksheetProtected Then Exit Sub
    
    With This
        .UserInterfaceOnly = .Worksheet.ProtectionMode
    End With
End Sub

'==============================================================================
Private Sub RaiseErrorIsNothing()
    Err.Raise _
            Source:=TypeName(Me), _
            Number:=eWorksheetProtectionError.ErrIsNothing, _
            Description:= _
                    "The 'Worksheet' is 'Nothing'." & vbCrLf & _
                    "Please check beforehand for that."
End Sub
