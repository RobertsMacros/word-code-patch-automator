Attribute VB_Name = "modModuleTransport"
'===========================================================================
' modModuleTransport — Export / Import VBA modules with folder picker
'
' Fixes:
'   - No hardcoded Mac UNC or platform-specific default paths
'   - Folder picker uses My Documents as a safe initial folder
'   - Selected folders are trusted as-is; no unnecessary creation
'   - Only _backup subfolder is created when needed
'   - Recursive MkDir for backup path (handles missing parents)
'   - Write-permission pre-flight test before export/import
'   - Clear, specific error messages for each failure mode
'===========================================================================
Option Explicit

Private Const BACKUP_SUBFOLDER As String = "_backup"

'===========================================================================
' Public entry points
'===========================================================================

Public Sub ExportAllModules()
    Dim folderPath As String
    folderPath = PickFolder("Select folder to export modules into")
    If folderPath = "" Then Exit Sub

    ' Validate the chosen folder
    If Not ValidateFolder(folderPath, "export") Then Exit Sub

    Dim proj As Object
    Set proj = ThisDocument.VBProject

    Dim comp As Object
    Dim exported As Long
    exported = 0

    Dim i As Long
    For i = 1 To proj.VBComponents.Count
        Set comp = proj.VBComponents(i)
        Dim ext As String
        ext = FileExtForComponent(comp)
        If ext <> "" Then
            Dim filePath As String
            filePath = folderPath & "\" & comp.Name & ext
            On Error GoTo ExportErr
            comp.Export filePath
            On Error GoTo 0
            exported = exported + 1
        End If
    Next i

    MsgBox "Exported " & exported & " module(s) to:" & vbCrLf & folderPath, _
           vbInformation, "Export Complete"
    Exit Sub

ExportErr:
    MsgBox "Failed to export component '" & comp.Name & "'." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "Target path: " & filePath, _
           vbCritical, "Export Error"
    Resume Next
End Sub

Public Sub ImportAllModules()
    Dim folderPath As String
    folderPath = PickFolder("Select folder to import modules from")
    If folderPath = "" Then Exit Sub

    ' Validate the chosen folder
    If Not ValidateFolder(folderPath, "import") Then Exit Sub

    ' Create backup before importing
    Dim backupPath As String
    backupPath = folderPath & "\" & BACKUP_SUBFOLDER
    If Not EnsureFolderExists(backupPath) Then
        MsgBox "Cannot create backup subfolder:" & vbCrLf & _
               backupPath & vbCrLf & vbCrLf & _
               "Check that you have write permission to the selected folder.", _
               vbCritical, "Backup Folder Error"
        Exit Sub
    End If

    Dim proj As Object
    Set proj = ThisDocument.VBProject

    ' Back up existing modules, then remove and re-import
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim importFile As Object
    Dim importFolder As Object
    Set importFolder = fso.GetFolder(folderPath)

    Dim imported As Long
    imported = 0

    For Each importFile In importFolder.Files
        Dim ext As String
        ext = LCase$(fso.GetExtensionName(importFile.Name))
        If ext = "bas" Or ext = "cls" Or ext = "frm" Then
            Dim modName As String
            modName = fso.GetBaseName(importFile.Name)

            ' Back up the existing component if present
            BackupComponent proj, modName, backupPath

            ' Remove existing component before importing
            On Error Resume Next
            Dim existing As Object
            Set existing = proj.VBComponents(modName)
            If Err.Number = 0 Then
                proj.VBComponents.Remove existing
            End If
            Err.Clear
            On Error GoTo 0

            ' Import
            On Error GoTo ImportErr
            proj.VBComponents.Import importFile.Path
            On Error GoTo 0
            imported = imported + 1
        End If
    Next importFile

    MsgBox "Imported " & imported & " module(s) from:" & vbCrLf & folderPath, _
           vbInformation, "Import Complete"
    Exit Sub

ImportErr:
    MsgBox "Failed to import '" & importFile.Name & "'." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "Import Error"
    Resume Next
End Sub

'===========================================================================
' Folder picker — no platform-specific initial path
'===========================================================================

Private Function PickFolder(prompt As String) As String
    Dim dlg As Object ' FileDialog
    Set dlg = Application.FileDialog(4) ' msoFileDialogFolderPicker = 4

    With dlg
        .Title = prompt
        .AllowMultiSelect = False

        ' Use My Documents as a sensible, cross-platform default.
        ' Environ("USERPROFILE") is always set on Windows; fall back gracefully.
        Dim myDocs As String
        myDocs = GetMyDocumentsPath()
        If myDocs <> "" Then
            ' InitialFileName must end with backslash for folders
            If Right$(myDocs, 1) <> "\" Then myDocs = myDocs & "\"
            .InitialFileName = myDocs
        End If
        ' If myDocs is empty, leave InitialFileName unset — Word will use
        ' its own default (usually the last-used folder), which is safe.

        If .Show = -1 Then
            PickFolder = .SelectedItems(1)
            ' Normalise: strip trailing backslash for consistency
            If Len(PickFolder) > 3 And Right$(PickFolder, 1) = "\" Then
                PickFolder = Left$(PickFolder, Len(PickFolder) - 1)
            End If
        Else
            PickFolder = ""
        End If
    End With
End Function

'===========================================================================
' Resolve My Documents path safely
'===========================================================================

Private Function GetMyDocumentsPath() As String
    ' Try the Shell special-folder approach first (works on all Windows)
    On Error Resume Next
    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")
    If Err.Number = 0 Then
        Dim myDocs As String
        myDocs = sh.SpecialFolders("MyDocuments")
        If Err.Number = 0 And myDocs <> "" Then
            GetMyDocumentsPath = myDocs
            Set sh = Nothing
            On Error GoTo 0
            Exit Function
        End If
    End If
    Err.Clear
    On Error GoTo 0

    ' Fallback: USERPROFILE\Documents
    Dim userProfile As String
    userProfile = Environ("USERPROFILE")
    If userProfile <> "" Then
        Dim candidate As String
        candidate = userProfile & "\Documents"
        If FolderExists(candidate) Then
            GetMyDocumentsPath = candidate
            Exit Function
        End If
    End If

    ' Give up — caller will leave InitialFileName unset
    GetMyDocumentsPath = ""
End Function

'===========================================================================
' Folder validation and write-permission pre-flight
'===========================================================================

Private Function ValidateFolder(folderPath As String, _
                                 operation As String) As Boolean
    ValidateFolder = False

    ' 1. Does the folder exist?
    If Not FolderExists(folderPath) Then
        MsgBox "The selected folder does not exist:" & vbCrLf & _
               folderPath & vbCrLf & vbCrLf & _
               "Please choose a folder that already exists.", _
               vbCritical, "Folder Not Found"
        Exit Function
    End If

    ' 2. Is the folder accessible? (try to list it)
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim fld As Object
    Set fld = fso.GetFolder(folderPath)
    If Err.Number <> 0 Then
        MsgBox "The selected folder is inaccessible:" & vbCrLf & _
               folderPath & vbCrLf & vbCrLf & _
               "Error: " & Err.Description & vbCrLf & _
               "Check network connectivity and permissions.", _
               vbCritical, "Folder Inaccessible"
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    ' 3. Write-permission test: create and delete a temp file
    Dim testFile As String
    testFile = folderPath & "\" & "_write_test_" & Format$(Now, "yyyymmddhhnnss") & ".tmp"

    Dim fNum As Integer
    fNum = FreeFile
    On Error Resume Next
    Open testFile For Output As #fNum
    If Err.Number <> 0 Then
        MsgBox "Cannot write to the selected folder:" & vbCrLf & _
               folderPath & vbCrLf & vbCrLf & _
               "Error: " & Err.Description & vbCrLf & _
               "You need write permission to " & operation & " modules.", _
               vbCritical, "Write Permission Denied"
        On Error GoTo 0
        Exit Function
    End If
    Print #fNum, "write test"
    Close #fNum

    ' Clean up test file
    Kill testFile
    Err.Clear
    On Error GoTo 0

    ValidateFolder = True
End Function

'===========================================================================
' Folder existence check — works for local and UNC paths
'===========================================================================

Private Function FolderExists(folderPath As String) As Boolean
    On Error Resume Next
    Dim attr As Long
    attr = GetAttr(folderPath)
    If Err.Number <> 0 Then
        FolderExists = False
    Else
        FolderExists = ((attr And vbDirectory) = vbDirectory)
    End If
    Err.Clear
    On Error GoTo 0
End Function

'===========================================================================
' Recursive folder creation — handles missing parent directories
'===========================================================================

Private Function EnsureFolderExists(folderPath As String) As Boolean
    If FolderExists(folderPath) Then
        EnsureFolderExists = True
        Exit Function
    End If

    ' Use FSO.CreateFolder via building the path from the root
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Err.Number <> 0 Then
        EnsureFolderExists = False
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    ' Walk the path segments and create each missing level
    Dim parts() As String
    Dim sep As String
    sep = "\"

    ' Handle UNC paths: \\server\share is the root, cannot be created
    Dim startSegment As Long
    Dim basePath As String

    If Left$(folderPath, 2) = "\\" Then
        ' UNC: \\server\share — find the share (3rd and 4th backslash)
        Dim thirdSlash As Long
        thirdSlash = InStr(3, folderPath, sep)
        If thirdSlash = 0 Then
            ' Only \\server — invalid
            EnsureFolderExists = False
            Exit Function
        End If
        Dim fourthSlash As Long
        fourthSlash = InStr(thirdSlash + 1, folderPath, sep)
        If fourthSlash = 0 Then
            ' \\server\share with no deeper path — check if it exists
            EnsureFolderExists = FolderExists(folderPath)
            Exit Function
        End If
        basePath = Left$(folderPath, fourthSlash - 1)
        Dim remainder As String
        remainder = Mid$(folderPath, fourthSlash + 1)
        parts = Split(remainder, sep)
        startSegment = 0
    Else
        ' Local path: drive letter is the root (e.g. C:\)
        parts = Split(folderPath, sep)
        basePath = parts(0) ' e.g. "C:"
        startSegment = 1
    End If

    ' Build path incrementally
    Dim current As String
    current = basePath
    Dim i As Long
    For i = startSegment To UBound(parts)
        If parts(i) <> "" Then
            current = current & sep & parts(i)
            If Not FolderExists(current) Then
                On Error Resume Next
                MkDir current
                If Err.Number <> 0 Then
                    EnsureFolderExists = False
                    Err.Clear
                    On Error GoTo 0
                    Exit Function
                End If
                On Error GoTo 0
            End If
        End If
    Next i

    EnsureFolderExists = FolderExists(folderPath)
End Function

'===========================================================================
' Back up a single component to the backup folder
'===========================================================================

Private Sub BackupComponent(proj As Object, modName As String, _
                             backupPath As String)
    On Error Resume Next
    Dim comp As Object
    Set comp = proj.VBComponents(modName)
    If Err.Number <> 0 Then
        ' Component does not exist — nothing to back up
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Dim ext As String
    ext = FileExtForComponent(comp)
    If ext = "" Then Exit Sub

    Dim destFile As String
    destFile = backupPath & "\" & modName & ext

    On Error Resume Next
    comp.Export destFile
    ' Silently ignore backup failures — the import will still proceed
    Err.Clear
    On Error GoTo 0
End Sub

'===========================================================================
' Map component type to file extension
'===========================================================================

Private Function FileExtForComponent(comp As Object) As String
    Select Case comp.Type
        Case 1  ' vbext_ct_StdModule
            FileExtForComponent = ".bas"
        Case 2  ' vbext_ct_ClassModule
            FileExtForComponent = ".cls"
        Case 3  ' vbext_ct_MSForm
            FileExtForComponent = ".frm"
        Case Else
            ' Document modules (ThisDocument) and other types — skip
            FileExtForComponent = ""
    End Select
End Function
