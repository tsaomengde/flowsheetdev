Attribute VB_Name = "modVBAExport"
Option Explicit

' ======= User helpers =======
Public Sub ExportAllVBAComponents()
    Dim targetRoot As String
    Dim exportPath As String
    
    If Not CheckVBATrustAccess() Then
        MsgBox "Please enable: File ? Options ? Trust Center ? Trust Center Settings ?" & vbCrLf & _
               "Macro Settings ? 'Trust access to the VBA project object model'.", vbExclamation, "Export Aborted"
        Exit Sub
    End If
    
    targetRoot = PickFolder("Choose a folder to receive the exported VBA components:")
    If Len(targetRoot) = 0 Then Exit Sub
    
    exportPath = CreateTimestampedSubfolder(targetRoot, ThisWorkbook.Name)
    If Len(Dir(exportPath, vbDirectory)) = 0 Then
        MsgBox "Failed to create export folder." & vbCrLf & exportPath, vbCritical
        Exit Sub
    End If
    
    Dim countExported As Long
    countExported = ExportVBProjectComponents(ThisWorkbook, exportPath)
    
    MsgBox "Export complete." & vbCrLf & _
           "Components exported: " & countExported & vbCrLf & _
           "Folder: " & exportPath, vbInformation, "VBA Export"
End Sub

' ======= Core export routine =======
Private Function ExportVBProjectComponents(wb As Workbook, ByVal exportPath As String) As Long
    Dim vbComp As Object ' VBIDE.VBComponent
    Dim filePath As String
    Dim baseName As String
    Dim exported As Long
    
    On Error GoTo CleanFail
    
    For Each vbComp In wb.VBProject.VBComponents
        baseName = SanitizeFileName(vbComp.Name)
        If Len(baseName) = 0 Then
            ' Skip unnamed
            GoTo NextComp
        End If
        
        Select Case vbComp.Type
            Case 1 ' vbext_ct_StdModule
                filePath = exportPath & baseName & ".bas"
                SafeDeleteIfExists filePath
                vbComp.Export filePath
                exported = exported + 1
            
            Case 2 ' vbext_ct_ClassModule (includes sheet class if Type=2? No; sheets are 100/3. Handle below.)
                filePath = exportPath & baseName & ".cls"
                SafeDeleteIfExists filePath
                vbComp.Export filePath
                exported = exported + 1
            
            Case 3 ' vbext_ct_MSForm (UserForm)
                filePath = exportPath & baseName & ".frm"
                SafeDeleteIfExists filePath
                vbComp.Export filePath
                ' VBE will also write a matching .frx next to it automatically
                exported = exported + 1
            
            Case 100, 34, 36, 35, 0, 11, 12, 13
                ' Excel exposes Sheet and ThisWorkbook as "document modules".
                ' In the type library: vbext_ct_Document = 100
                ' Treat them like class modules for export.
                filePath = exportPath & baseName & ".cls"
                SafeDeleteIfExists filePath
                vbComp.Export filePath
                exported = exported + 1
            
            Case Else
                ' Unknown or host-specific; attempt class export as a safe fallback
                filePath = exportPath & baseName & ".cls"
                SafeDeleteIfExists filePath
                On Error Resume Next
                vbComp.Export filePath
                If Err.Number = 0 Then
                    exported = exported + 1
                End If
                On Error GoTo CleanFail
        End Select
        
NextComp:
    Next vbComp
    
    ExportVBProjectComponents = exported
    Exit Function

CleanFail:
    MsgBox "Error during export: " & Err.Number & " - " & Err.Description, vbCritical, "VBA Export"
End Function

' ======= Utilities =======
Private Function PickFolder(Optional ByVal prompt As String = "Select folder") As String
    Dim dlg As FileDialog
    On Error GoTo Fallback
    Set dlg = Application.FileDialog(msoFileDialogFolderPicker)
    With dlg
        .title = prompt
        .AllowMultiSelect = False
        If .Show = -1 Then
            PickFolder = EnsureTrailingSlash(.SelectedItems(1))
        Else
            PickFolder = vbNullString
        End If
    End With
    Exit Function

Fallback:
    ' Fallback to default Documents if FileDialog not available
    PickFolder = EnsureTrailingSlash(CreateObject("WScript.Shell").SpecialFolders("MyDocuments"))
End Function

Private Function EnsureTrailingSlash(ByVal path As String) As String
    If Len(path) = 0 Then
        EnsureTrailingSlash = ""
    ElseIf Right$(path, 1) = Application.PathSeparator Then
        EnsureTrailingSlash = path
    Else
        EnsureTrailingSlash = path & Application.PathSeparator
    End If
End Function

Private Function CreateTimestampedSubfolder(ByVal root As String, ByVal wbName As String) As String
    Dim safeWB As String
    Dim ts As String
    safeWB = SanitizeFileName(RemoveFileExtension(wbName))
    ts = Format(Now, "yyyy-mm-dd_hh-nn-ss")
    CreateTimestampedSubfolder = EnsureTrailingSlash(root) & safeWB & "_VBAExport_" & ts & Application.PathSeparator
    On Error Resume Next
    MkDir CreateTimestampedSubfolder
    On Error GoTo 0
End Function

Private Function RemoveFileExtension(ByVal fileName As String) As String
    Dim p As Long
    p = InStrRev(fileName, ".")
    If p > 0 Then
        RemoveFileExtension = Left$(fileName, p - 1)
    Else
        RemoveFileExtension = fileName
    End If
End Function

Private Function SanitizeFileName(ByVal s As String) As String
    Dim badChars As Variant
    Dim i As Long
    badChars = Array("<", ">", ":", """", "/", "\", "|", "?", "*")
    For i = LBound(badChars) To UBound(badChars)
        s = Replace$(s, badChars(i), "_")
    Next i
    ' Also trim whitespace
    SanitizeFileName = Trim$(s)
End Function

Private Sub SafeDeleteIfExists(ByVal fullPath As String)
    On Error Resume Next
    If Len(Dir(fullPath, vbNormal)) > 0 Then
        Kill fullPath
    End If
    On Error GoTo 0
End Sub

Private Function CheckVBATrustAccess() As Boolean
    ' There is no reliable API to read Trust Center settings programmatically.
    ' We attempt a benign access and infer.
    On Error Resume Next
    Dim test As Long
    test = ThisWorkbook.VBProject.VBComponents.Count
    CheckVBATrustAccess = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

