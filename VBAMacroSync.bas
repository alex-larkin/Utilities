Attribute VB_Name = "VBAMacroSync"
' ========================================================================
' VBA MACRO SYNC SYSTEM
' Two-way synchronization between Word templates and local folders
' Multi-template support - works with Normal.dotm, Utilities.dotm, and custom templates
' ========================================================================
'
' This module resides in Utilities.dotm and provides sync services for all templates.
' Each template that wants to use sync must implement simple AutoExec/AutoExit macros
' that call VMS_AutoExec() and VMS_AutoExit() in this module.
'
' See VBAMAcroSync_SRD.md for complete specification.
' ========================================================================

' ========================================================================
' WINDOWS API DECLARATION
' ========================================================================
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' ========================================================================
' CONFIGURATION - Paths are read from environment variables
' ========================================================================

' Get the sync folder path for a specific template
' Template name should be the base name without extension (e.g., "Normal", "Utilities")
' Returns the path from environment variable MACROS_<TEMPLATENAME>
Private Function GetSyncFolderPath(templateName As String) As String
    Dim envVarName As String
    Dim envPath As String

    ' Construct environment variable name: MACROS_<TEMPLATENAME>
    envVarName = "MACROS_" & UCase(templateName)
    envPath = Environ(envVarName)

    If envPath <> "" Then
        GetSyncFolderPath = envPath
    Else
        Debug.Print "WARNING: Environment variable '" & envVarName & "' is not set."
        Debug.Print "Skipping sync for " & templateName & ".dotm"
        GetSyncFolderPath = ""
    End If
End Function

' Extract template base name from VBProject filename
' E.g., "C:\...\Normal.dotm" -> "Normal"
Private Function GetTemplateBaseName(vbProj As Object) As String
    Dim fullPath As String
    Dim fileName As String
    Dim baseName As String

    fullPath = vbProj.fileName
    fileName = Dir(fullPath) ' Get just the filename

    ' Remove .dotm or .dotx extension
    If Right(fileName, 5) = ".dotm" Then
        baseName = Left(fileName, Len(fileName) - 5)
    ElseIf Right(fileName, 5) = ".dotx" Then
        baseName = Left(fileName, Len(fileName) - 5)
    Else
        baseName = fileName
    End If

    GetTemplateBaseName = baseName
End Function

' ========================================================================
' AUTO-RUN ENTRY POINTS (Called by templates via Application.Run)
' ========================================================================

' This is called by AutoExec() macros in templates when Word starts
' It syncs all templates that have environment variables configured
Public Sub VMS_AutoExec()
    Debug.Print "========================================="
    Debug.Print "=== VMS_AutoExec START ==="
    Debug.Print "Current Time: " & Now
    Debug.Print "Word Version: " & Application.VERSION
    Debug.Print "========================================="

    On Error Resume Next

    Dim vbProj As Object
    Dim templateName As String
    Dim syncPath As String
    Dim templateCount As Integer
    Dim syncCount As Integer

    templateCount = 0
    syncCount = 0

    ' Iterate through all loaded templates
    Debug.Print "Scanning for templates to sync..."
    For Each vbProj In Application.VBE.VBProjects
        templateCount = templateCount + 1

        ' Get template base name
        templateName = GetTemplateBaseName(vbProj)
        Debug.Print ""
        Debug.Print "--- Template #" & templateCount & ": " & templateName & " ---"
        Debug.Print "Full path: " & vbProj.fileName

        ' Get sync folder path for this template
        syncPath = GetSyncFolderPath(templateName)

        If syncPath <> "" Then
            Debug.Print "Sync folder: " & syncPath
            Debug.Print ">>> IMPORTING macros from folder to " & templateName & ".dotm <<<"

            ' Import macros from folder (folder is source of truth at startup)
            ImportMacrosFromFolder vbProj, syncPath

            If Err.Number <> 0 Then
                Debug.Print "ERROR in import: " & Err.Number & " - " & Err.Description
                Err.Clear
            Else
                syncCount = syncCount + 1
                Debug.Print "Import completed for " & templateName
            End If
        Else
            Debug.Print "No environment variable - skipping sync for this template"
        End If
    Next vbProj

    Debug.Print ""
    Debug.Print "========================================="
    Debug.Print "Templates scanned: " & templateCount
    Debug.Print "Templates synced: " & syncCount
    Debug.Print "=== VMS_AutoExec END ==="
    Debug.Print "========================================="

    ' Show a brief status message
    Application.StatusBar = "Macro sync complete"
    Sleep 2000 ' Show for 2 seconds
    Application.StatusBar = False ' Clear status bar

    On Error GoTo 0
End Sub

' This is called by AutoExit() macros in templates when Word closes
' It syncs all templates that have environment variables configured
Public Sub VMS_AutoExit()
    Debug.Print "========================================="
    Debug.Print "=== VMS_AutoExit START ==="
    Debug.Print "Current Time: " & Now
    Debug.Print "========================================="

    On Error Resume Next

    Dim vbProj As Object
    Dim templateName As String
    Dim syncPath As String
    Dim templateCount As Integer
    Dim syncCount As Integer

    templateCount = 0
    syncCount = 0

    ' Iterate through all loaded templates
    Debug.Print "Scanning for templates to sync..."
    For Each vbProj In Application.VBE.VBProjects
        templateCount = templateCount + 1

        ' Get template base name
        templateName = GetTemplateBaseName(vbProj)
        Debug.Print ""
        Debug.Print "--- Template #" & templateCount & ": " & templateName & " ---"
        Debug.Print "Full path: " & vbProj.fileName

        ' Get sync folder path for this template
        syncPath = GetSyncFolderPath(templateName)

        If syncPath <> "" Then
            Debug.Print "Sync folder: " & syncPath
            Debug.Print ">>> EXPORTING macros from " & templateName & ".dotm to folder <<<"

            ' Export macros to folder (template is source of truth at shutdown)
            ExportMacrosToFolder vbProj, syncPath

            If Err.Number <> 0 Then
                Debug.Print "ERROR in export: " & Err.Number & " - " & Err.Description
                Err.Clear
            Else
                syncCount = syncCount + 1
                Debug.Print "Export completed for " & templateName
            End If
        Else
            Debug.Print "No environment variable - skipping sync for this template"
        End If
    Next vbProj

    Debug.Print ""
    Debug.Print "========================================="
    Debug.Print "Templates scanned: " & templateCount
    Debug.Print "Templates synced: " & syncCount
    Debug.Print "=== VMS_AutoExit END ==="
    Debug.Print "========================================="

    On Error GoTo 0
End Sub

' ========================================================================
' EXPORT FUNCTIONALITY
' ========================================================================

' Export all modules from a VBProject to its sync folder
Private Sub ExportMacrosToFolder(vbProj As Object, syncPath As String)
    Debug.Print "  --- ExportMacrosToFolder START ---"
    On Error Resume Next

    Dim vbComp As Object ' VBComponent
    Dim exportPath As String
    Dim exportCount As Integer
    Dim fileExt As String
    Dim totalComponents As Integer

    Debug.Print "  Checking sync folder: " & syncPath

    ' Make sure the sync folder exists
    If Dir(syncPath, vbDirectory) = "" Then
        Debug.Print "  Sync folder does not exist, attempting to create..."
        MkDir syncPath
        If Err.Number <> 0 Then
            Debug.Print "  ERROR creating folder: " & Err.Number & " - " & Err.Description
            Err.Clear
            Exit Sub
        End If
        Debug.Print "  Created sync folder: " & syncPath
    Else
        Debug.Print "  Sync folder exists"
    End If

    exportCount = 0
    totalComponents = 0

    Debug.Print "  Accessing VBProject.VBComponents..."
    If Err.Number <> 0 Then
        Debug.Print "  ERROR accessing VBProject: " & Err.Number & " - " & Err.Description
        Debug.Print "  VBA Project access may be disabled. Check Trust Center settings."
        Err.Clear
        Exit Sub
    End If

    ' Loop through all VBA components in the VBProject
    For Each vbComp In vbProj.VBComponents
        totalComponents = totalComponents + 1
        Debug.Print "  Component #" & totalComponents & ": " & vbComp.Name & " (Type: " & vbComp.Type & ")"

        ' Determine the file extension based on component type
        Select Case vbComp.Type
            Case 1 ' vbext_ct_StdModule - Standard module
                fileExt = ".bas"
                Debug.Print "    -> Standard Module"
            Case 2 ' vbext_ct_ClassModule - Class module
                fileExt = ".cls"
                Debug.Print "    -> Class Module"
            Case 3 ' vbext_ct_MSForm - UserForm
                fileExt = ".frm"
                Debug.Print "    -> UserForm"
            Case Else
                fileExt = "" ' Skip document modules and other types
                Debug.Print "    -> Skipping (Type " & vbComp.Type & " not exportable)"
        End Select

        ' Only export if we have a valid file extension
        If fileExt <> "" Then
            exportPath = syncPath & vbComp.Name & fileExt
            Debug.Print "    -> Exporting to: " & exportPath

            ' Export the component to a file
            vbComp.Export exportPath
            If Err.Number <> 0 Then
                Debug.Print "    -> ERROR exporting: " & Err.Number & " - " & Err.Description
                Err.Clear
            Else
                exportCount = exportCount + 1
                Debug.Print "    -> Exported successfully: " & vbComp.Name & fileExt
            End If
        End If
    Next vbComp

    Debug.Print "  Total components found: " & totalComponents
    Debug.Print "  Total exported: " & exportCount & " module(s)"
    Debug.Print "  --- ExportMacrosToFolder END ---"
End Sub

' ========================================================================
' IMPORT FUNCTIONALITY
' ========================================================================

' Import modules from the sync folder into a VBProject
Private Sub ImportMacrosFromFolder(vbProj As Object, syncPath As String)
    Debug.Print "  --- ImportMacrosFromFolder START ---"
    On Error Resume Next

    Dim fileName As String
    Dim fullPath As String
    Dim moduleName As String
    Dim importCount As Integer
    Dim basFileCount As Integer
    Dim clsFileCount As Integer
    Dim frmFileCount As Integer

    Debug.Print "  Checking sync folder: " & syncPath

    ' Check if sync folder exists
    If Dir(syncPath, vbDirectory) = "" Then
        Debug.Print "  ERROR: Sync folder does not exist: " & syncPath
        Exit Sub
    Else
        Debug.Print "  Sync folder exists"
    End If

    Debug.Print "  Accessing VBProject for import..."
    If Err.Number <> 0 Then
        Debug.Print "  ERROR accessing VBProject: " & Err.Number & " - " & Err.Description
        Debug.Print "  VBA Project access may be disabled. Check Trust Center settings."
        Err.Clear
        Exit Sub
    End If

    importCount = 0
    basFileCount = 0
    clsFileCount = 0
    frmFileCount = 0

    ' Process .bas files (standard modules)
    Debug.Print "  Searching for .bas files..."
    fileName = Dir(syncPath & "*.bas")
    If fileName = "" Then
        Debug.Print "  No .bas files found in folder"
    End If

    Do While fileName <> ""
        basFileCount = basFileCount + 1
        fullPath = syncPath & fileName
        moduleName = Left(fileName, Len(fileName) - 4) ' Remove .bas extension

        Debug.Print "  Found .bas file #" & basFileCount & ": " & fileName
        Debug.Print "    Full path: " & fullPath
        Debug.Print "    Module name: " & moduleName

        ' Import from folder (Git-aware: folder is source of truth at startup)
        If ProcessImport(vbProj, fullPath, moduleName, ".bas", syncPath) Then
            importCount = importCount + 1
        End If

        If Err.Number <> 0 Then
            Debug.Print "    ERROR processing: " & Err.Number & " - " & Err.Description
            Err.Clear
        End If

        fileName = Dir() ' Get next file
    Loop
    Debug.Print "  Total .bas files found: " & basFileCount

    ' Process .cls files (class modules)
    Debug.Print "  Searching for .cls files..."
    fileName = Dir(syncPath & "*.cls")
    If fileName = "" Then
        Debug.Print "  No .cls files found in folder"
    End If

    Do While fileName <> ""
        clsFileCount = clsFileCount + 1
        fullPath = syncPath & fileName
        moduleName = Left(fileName, Len(fileName) - 4) ' Remove .cls extension

        Debug.Print "  Found .cls file #" & clsFileCount & ": " & fileName

        If ProcessImport(vbProj, fullPath, moduleName, ".cls", syncPath) Then
            importCount = importCount + 1
        End If

        If Err.Number <> 0 Then
            Debug.Print "    ERROR processing: " & Err.Number & " - " & Err.Description
            Err.Clear
        End If

        fileName = Dir()
    Loop
    Debug.Print "  Total .cls files found: " & clsFileCount

    ' Process .frm files (UserForms)
    Debug.Print "  Searching for .frm files..."
    fileName = Dir(syncPath & "*.frm")
    If fileName = "" Then
        Debug.Print "  No .frm files found in folder"
    End If

    Do While fileName <> ""
        frmFileCount = frmFileCount + 1
        fullPath = syncPath & fileName
        moduleName = Left(fileName, Len(fileName) - 4) ' Remove .frm extension

        Debug.Print "  Found .frm file #" & frmFileCount & ": " & fileName

        If ProcessImport(vbProj, fullPath, moduleName, ".frm", syncPath) Then
            importCount = importCount + 1
        End If

        If Err.Number <> 0 Then
            Debug.Print "    ERROR processing: " & Err.Number & " - " & Err.Description
            Err.Clear
        End If

        fileName = Dir()
    Loop
    Debug.Print "  Total .frm files found: " & frmFileCount

    Debug.Print "  Total imported: " & importCount & " module(s)"
    Debug.Print "  --- ImportMacrosFromFolder END ---"
End Sub

' Process a single import (Git-aware: folder is source of truth at startup)
Private Function ProcessImport(vbProj As Object, filePath As String, moduleName As String, fileExt As String, syncPath As String) As Boolean
    Debug.Print "    -> ProcessImport START for: " & moduleName & fileExt
    On Error Resume Next

    Dim vbComp As Object
    Dim moduleExists As Boolean
    Dim filesIdentical As Boolean
    Dim tempExportPath As String

    ProcessImport = False ' Default to False

    ' Check if module already exists in the VBProject
    Debug.Print "    -> Checking if module already exists in template..."
    moduleExists = False
    For Each vbComp In vbProj.VBComponents
        If vbComp.Name = moduleName Then
            moduleExists = True
            Debug.Print "    -> Module EXISTS in template: " & moduleName
            Exit For
        End If
    Next vbComp

    If Not moduleExists Then
        Debug.Print "    -> Module does NOT exist in template (will import new module)"
    End If

    ' If module exists, check if files are identical (optimization to skip unnecessary imports)
    If moduleExists Then
        Debug.Print "    -> Comparing with existing module..."
        ' Export current version to a temp file for comparison
        tempExportPath = syncPath & "~temp_" & moduleName & fileExt
        Debug.Print "    -> Exporting current version to temp file: " & tempExportPath

        vbComp.Export tempExportPath
        If Err.Number <> 0 Then
            Debug.Print "    -> ERROR exporting to temp file: " & Err.Number & " - " & Err.Description
            Err.Clear
            Exit Function
        End If

        ' Compare the two files
        Debug.Print "    -> Comparing files..."
        filesIdentical = FilesAreIdentical(filePath, tempExportPath)

        ' Delete temp file
        Debug.Print "    -> Deleting temp file..."
        Kill tempExportPath
        If Err.Number <> 0 Then
            Debug.Print "    -> ERROR deleting temp file: " & Err.Number & " - " & Err.Description
            Err.Clear
        End If

        ' If files are identical, skip import
        If filesIdentical Then
            Debug.Print "    -> Files are IDENTICAL - skipping import (already in sync)"
            ProcessImport = False
            Exit Function
        Else
            Debug.Print "    -> Files are DIFFERENT - will import folder version (Git-managed)"
        End If

        ' Files differ: Remove existing module to import new version
        Debug.Print "    -> Removing existing module from template..."
        vbProj.VBComponents.Remove vbComp
        If Err.Number <> 0 Then
            Debug.Print "    -> ERROR removing module: " & Err.Number & " - " & Err.Description
            Err.Clear
            Exit Function
        End If
        Debug.Print "    -> Module removed successfully"
    End If

    ' Import the module from file (folder is source of truth)
    Debug.Print "    -> Importing module from file: " & filePath
    vbProj.VBComponents.Import filePath
    If Err.Number <> 0 Then
        Debug.Print "    -> ERROR importing module: " & Err.Number & " - " & Err.Description
        Err.Clear
        Exit Function
    End If

    Debug.Print "    -> Import SUCCESS: " & moduleName & fileExt
    ProcessImport = True
    Debug.Print "    -> ProcessImport END (success)"
End Function

' ========================================================================
' HELPER FUNCTIONS
' ========================================================================

' Compare two files to see if they're identical
Private Function FilesAreIdentical(file1 As String, file2 As String) As Boolean
    Debug.Print "      -> FilesAreIdentical comparing:"
    Debug.Print "         File1: " & file1
    Debug.Print "         File2: " & file2
    On Error Resume Next

    Dim fso As Object
    Dim f1 As Object
    Dim f2 As Object
    Dim content1 As String
    Dim content2 As String
    Dim size1 As Long
    Dim size2 As Long

    FilesAreIdentical = False ' Default to False

    ' Create FileSystemObject for file operations
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Err.Number <> 0 Then
        Debug.Print "      -> ERROR creating FileSystemObject: " & Err.Number & " - " & Err.Description
        Err.Clear
        Exit Function
    End If

    ' Check if both files exist
    If Not fso.FileExists(file1) Then
        Debug.Print "      -> File1 does NOT exist"
        Exit Function
    End If
    If Not fso.FileExists(file2) Then
        Debug.Print "      -> File2 does NOT exist"
        Exit Function
    End If
    Debug.Print "      -> Both files exist"

    ' Quick check: if file sizes differ, they're different
    size1 = fso.GetFile(file1).Size
    size2 = fso.GetFile(file2).Size
    Debug.Print "      -> File1 size: " & size1 & " bytes"
    Debug.Print "      -> File2 size: " & size2 & " bytes"

    If size1 <> size2 Then
        Debug.Print "      -> Files are DIFFERENT (size mismatch)"
        Exit Function
    End If

    ' Read and compare file contents
    Debug.Print "      -> Reading file contents for comparison..."
    Set f1 = fso.OpenTextFile(file1, 1) ' 1 = ForReading
    If Err.Number <> 0 Then
        Debug.Print "      -> ERROR opening File1: " & Err.Number & " - " & Err.Description
        Err.Clear
        Exit Function
    End If

    Set f2 = fso.OpenTextFile(file2, 1)
    If Err.Number <> 0 Then
        Debug.Print "      -> ERROR opening File2: " & Err.Number & " - " & Err.Description
        f1.Close
        Err.Clear
        Exit Function
    End If

    content1 = f1.ReadAll
    content2 = f2.ReadAll

    f1.Close
    f2.Close

    ' Compare content
    If content1 = content2 Then
        Debug.Print "      -> Files are IDENTICAL (content matches)"
        FilesAreIdentical = True
    Else
        Debug.Print "      -> Files are DIFFERENT (content differs)"
        FilesAreIdentical = False
    End If
End Function

' ========================================================================
' MANUAL TRIGGER SUBS (For testing - syncs all templates)
' ========================================================================

' Manually export all templates (same as VMS_AutoExit)
Public Sub ManualExport()
    Debug.Print ""
    Debug.Print "========================================="
    Debug.Print "MANUAL EXPORT TRIGGERED"
    Debug.Print "========================================="

    VMS_AutoExit

    MsgBox "Manual export complete! Check VBA Immediate Window (Ctrl+G) for details.", vbInformation, "VBA Macro Sync"
End Sub

' Manually import all templates (same as VMS_AutoExec)
Public Sub ManualImport()
    Debug.Print ""
    Debug.Print "========================================="
    Debug.Print "MANUAL IMPORT TRIGGERED"
    Debug.Print "========================================="

    VMS_AutoExec

    MsgBox "Manual import complete! Check VBA Immediate Window (Ctrl+G) for details.", vbInformation, "VBA Macro Sync"
End Sub
