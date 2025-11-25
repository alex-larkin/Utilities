Attribute VB_Name = "VBAMacroSync"
' ========================================================================
' WORD MACRO SYNC SYSTEM
' Two-way synchronization between Normal.dotm macros and a local folder
' ========================================================================

' ========================================================================
' WINDOWS API DECLARATION
' ========================================================================
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' ========================================================================
' CONFIGURATION - Path is read from environment variable
' ========================================================================
Private Function GetSyncFolderPath() As String
    Dim envPath As String
    envPath = Environ("MACROS_UTILITIES")

    If envPath <> "" Then
        GetSyncFolderPath = envPath
    Else
        MsgBox "ERROR: Environment variable 'MACROS_UTILITIES' is not set." & vbCrLf & vbCrLf & _
               "Please set this variable to your sync folder path." & vbCrLf & _
               "See README.md for setup instructions.", vbCritical, "VBA Macro Sync"
        GetSyncFolderPath = ""
    End If
End Function

' Legacy constant for reference (not used)
' Const SYNC_FOLDER_PATH As String = "C:\Users\YourUsername\Your\Path\Here\"
' Don't use a folder that autosyncs with Dropbox, OneDrive, or the like (recommended).
' Make sure the path ends with a backslash (\)

' ========================================================================
' AUTO-RUN PROCEDURES
' ========================================================================

' This runs automatically when Word starts
Sub AutoExec()
    Debug.Print "=== AutoExec START ==="
    Debug.Print "Current Time: " & Now
    Debug.Print "Word Version: " & Application.VERSION
    Debug.Print "Normal Template Path: " & Application.NormalTemplate.FullName

    On Error Resume Next

    ' Log the sync operation
    Debug.Print "=== Word Startup: Importing macros from folder ==="
    Debug.Print "Sync Folder Path: " & GetSyncFolderPath()

    ' Import any new or modified macros from the sync folder
    Debug.Print "Calling ImportMacrosFromFolder..."
    ImportMacrosFromFolder

    If Err.Number <> 0 Then
        Debug.Print "ERROR in AutoExec: " & Err.Number & " - " & Err.Description
        Err.Clear
    Else
        Debug.Print "ImportMacrosFromFolder completed without errors"
    End If

    ' Show a brief status message
    Application.StatusBar = "Macro sync complete"
    Sleep 2000 ' Show for 2 seconds (2000 milliseconds)
    Application.StatusBar = False ' Clear status bar
    
    ' Load and hide all templates to make VBA projects accessible
    Debug.Print "Getting ready to call OpenAllTemplates"

    ' First, explicitly load Utilities.dotm so we can call its macros
    On Error Resume Next
    Dim utilDoc As Document
    Dim utilPath As String
    utilPath = Environ("APPDATA") & "\Microsoft\Templates\Utilities.dotm"

    Debug.Print "Loading Utilities.dotm from: " & utilPath
    Set utilDoc = Documents.Open(fileName:=utilPath, _
                                  ReadOnly:=False, _
                                  AddToRecentFiles:=False, _
                                  Visible:=False)

    If Err.Number <> 0 Then
        Debug.Print "ERROR loading Utilities.dotm: " & Err.Number & " - " & Err.Description
        Err.Clear
    Else
        Debug.Print "Utilities.dotm loaded successfully"

        ' === DIAGNOSTIC SECTION ===
        Debug.Print "--- DIAGNOSTICS: Examining loaded VBA projects ---"

        ' Check if the document is actually loaded
        Debug.Print "utilDoc.Name: " & utilDoc.Name
        Debug.Print "utilDoc.Path: " & utilDoc.Path

        ' List all available VBA projects
        Debug.Print "Total VBA Projects available: " & Application.VBE.VBProjects.Count
        Dim i As Integer
        Dim proj As Object
        For i = 1 To Application.VBE.VBProjects.Count
            Set proj = Application.VBE.VBProjects(i)
            Debug.Print "  Project " & i & ": " & proj.Name & " (FileName: " & proj.fileName & ")"
        Next i

        ' Try to access the Utilities.dotm VBProject
        Dim utilProj As Object
        Set utilProj = utilDoc.VBProject
        Debug.Print "Utilities VBProject.Name: " & utilProj.Name

        ' List all components (modules) in Utilities.dotm
        Debug.Print "Components in Utilities VBProject: " & utilProj.VBComponents.Count
        Dim comp As Object
        For Each comp In utilProj.VBComponents
            Debug.Print "  Component: " & comp.Name & " (Type: " & comp.Type & ")"
        Next comp

        Debug.Print "--- END DIAGNOSTICS ---"

        ' Now we can call its macro
        ' Use the actual project name from diagnostics: "Project"
        Debug.Print "Calling LoadTemplates..."
        Application.Run "Utilities.OpenAllTemplates_v2.LoadTemplates"

        If Err.Number <> 0 Then
            Debug.Print "ERROR calling LoadTemplates: " & Err.Number & " - " & Err.Description
            Err.Clear
        Else
            Debug.Print "LoadTemplates completed successfully"
        End If
    End If
    On Error GoTo 0


    Debug.Print "=== AutoExec END ==="
    
    
End Sub

' This runs automatically when Word closes
Sub AutoExit()
    Debug.Print "=== AutoExit START ==="
    Debug.Print "Current Time: " & Now

    On Error Resume Next

    ' Log the sync operation
    Debug.Print "=== Word Shutdown: Exporting macros to folder ==="

    ' Export all macros to the sync folder
    Debug.Print "Calling ExportMacrosToFolder..."
    ExportMacrosToFolder

    If Err.Number <> 0 Then
        Debug.Print "ERROR in AutoExit: " & Err.Number & " - " & Err.Description
        Err.Clear
    End If

    Debug.Print "=== AutoExit END ==="
End Sub

' ========================================================================
' EXPORT FUNCTIONALITY
' ========================================================================

' Export all modules from Normal.dotm to the sync folder
Sub ExportMacrosToFolder()
    Debug.Print "--- ExportMacrosToFolder START ---"
    On Error Resume Next

    Dim vbComp As Object ' VBComponent
    Dim exportPath As String
    Dim exportCount As Integer
    Dim fileExt As String
    Dim totalComponents As Integer
    Dim syncPath As String

    syncPath = GetSyncFolderPath()
    If syncPath = "" Then Exit Sub ' Exit if environment variable not set

    Debug.Print "Checking sync folder: " & syncPath

    ' Make sure the sync folder exists
    If Dir(syncPath, vbDirectory) = "" Then
        Debug.Print "Sync folder does not exist, attempting to create..."
        MkDir syncPath
        If Err.Number <> 0 Then
            Debug.Print "ERROR creating folder: " & Err.Number & " - " & Err.Description
            Err.Clear
            Exit Sub
        End If
        Debug.Print "Created sync folder: " & syncPath
    Else
        Debug.Print "Sync folder exists"
    End If

    exportCount = 0
    totalComponents = 0

    Debug.Print "Accessing VBProject.VBComponents..."
    If Err.Number <> 0 Then
        Debug.Print "ERROR accessing VBProject: " & Err.Number & " - " & Err.Description
        Debug.Print "VBA Project access may be disabled. Check Trust Center settings."
        Err.Clear
        Exit Sub
    End If

    ' Loop through all VBA components in Normal.dotm
    For Each vbComp In Application.NormalTemplate.VBProject.VBComponents
        totalComponents = totalComponents + 1
        Debug.Print "Component #" & totalComponents & ": " & vbComp.Name & " (Type: " & vbComp.Type & ")"

        ' Determine the file extension based on component type
        Select Case vbComp.Type
            Case 1 ' vbext_ct_StdModule - Standard module
                fileExt = ".bas"
                Debug.Print "  -> Standard Module"
            Case 2 ' vbext_ct_ClassModule - Class module
                fileExt = ".cls"
                Debug.Print "  -> Class Module"
            Case 3 ' vbext_ct_MSForm - UserForm
                fileExt = ".frm"
                Debug.Print "  -> UserForm"
            Case Else
                fileExt = "" ' Skip document modules and other types
                Debug.Print "  -> Skipping (Type " & vbComp.Type & " not exportable)"
        End Select

        ' Only export if we have a valid file extension
        If fileExt <> "" Then
            exportPath = syncPath & vbComp.Name & fileExt
            Debug.Print "  -> Exporting to: " & exportPath

            ' Export the component to a file
            vbComp.Export exportPath
            If Err.Number <> 0 Then
                Debug.Print "  -> ERROR exporting: " & Err.Number & " - " & Err.Description
                Err.Clear
            Else
                exportCount = exportCount + 1
                Debug.Print "  -> Exported successfully: " & vbComp.Name & fileExt
            End If
        End If
    Next vbComp

    Debug.Print "Total components found: " & totalComponents
    Debug.Print "Total exported: " & exportCount & " module(s)"
    Debug.Print "--- ExportMacrosToFolder END ---"
End Sub

' ========================================================================
' IMPORT FUNCTIONALITY
' ========================================================================

' Import modules from the sync folder into Normal.dotm
Sub ImportMacrosFromFolder()
    Debug.Print "--- ImportMacrosFromFolder START ---"
    On Error Resume Next

    Dim fileName As String
    Dim fullPath As String
    Dim moduleName As String
    Dim fileExt As String
    Dim importCount As Integer
    Dim basFileCount As Integer
    Dim clsFileCount As Integer
    Dim frmFileCount As Integer
    Dim syncPath As String

    syncPath = GetSyncFolderPath()
    If syncPath = "" Then Exit Sub ' Exit if environment variable not set

    Debug.Print "Checking sync folder: " & syncPath

    ' Check if sync folder exists
    If Dir(syncPath, vbDirectory) = "" Then
        Debug.Print "ERROR: Sync folder does not exist: " & syncPath
        Exit Sub
    Else
        Debug.Print "Sync folder exists"
    End If

    Debug.Print "Accessing VBProject for import..."
    If Err.Number <> 0 Then
        Debug.Print "ERROR accessing VBProject: " & Err.Number & " - " & Err.Description
        Debug.Print "VBA Project access may be disabled. Check Trust Center settings."
        Err.Clear
        Exit Sub
    End If

    importCount = 0
    basFileCount = 0
    clsFileCount = 0
    frmFileCount = 0

    ' Process .bas files (standard modules)
    Debug.Print "Searching for .bas files..."
    fileName = Dir(syncPath & "*.bas")
    If fileName = "" Then
        Debug.Print "No .bas files found in folder"
    End If

    Do While fileName <> ""
        basFileCount = basFileCount + 1
        fullPath = syncPath & fileName
        moduleName = Left(fileName, Len(fileName) - 4) ' Remove .bas extension

        Debug.Print "Found .bas file #" & basFileCount & ": " & fileName
        Debug.Print "  Full path: " & fullPath
        Debug.Print "  Module name: " & moduleName

        ' Import from folder (Git-aware: folder is source of truth)
        If ProcessImport(fullPath, moduleName, ".bas") Then
            importCount = importCount + 1
        End If

        If Err.Number <> 0 Then
            Debug.Print "  ERROR processing: " & Err.Number & " - " & Err.Description
            Err.Clear
        End If

        fileName = Dir() ' Get next file
    Loop
    Debug.Print "Total .bas files found: " & basFileCount

    ' Process .cls files (class modules)
    Debug.Print "Searching for .cls files..."
    fileName = Dir(syncPath & "*.cls")
    If fileName = "" Then
        Debug.Print "No .cls files found in folder"
    End If

    Do While fileName <> ""
        clsFileCount = clsFileCount + 1
        fullPath = syncPath & fileName
        moduleName = Left(fileName, Len(fileName) - 4) ' Remove .cls extension

        Debug.Print "Found .cls file #" & clsFileCount & ": " & fileName

        If ProcessImport(fullPath, moduleName, ".cls") Then
            importCount = importCount + 1
        End If

        If Err.Number <> 0 Then
            Debug.Print "  ERROR processing: " & Err.Number & " - " & Err.Description
            Err.Clear
        End If

        fileName = Dir()
    Loop
    Debug.Print "Total .cls files found: " & clsFileCount

    ' Process .frm files (UserForms)
    Debug.Print "Searching for .frm files..."
    fileName = Dir(syncPath & "*.frm")
    If fileName = "" Then
        Debug.Print "No .frm files found in folder"
    End If

    Do While fileName <> ""
        frmFileCount = frmFileCount + 1
        fullPath = syncPath & fileName
        moduleName = Left(fileName, Len(fileName) - 4) ' Remove .frm extension

        Debug.Print "Found .frm file #" & frmFileCount & ": " & fileName

        If ProcessImport(fullPath, moduleName, ".frm") Then
            importCount = importCount + 1
        End If

        If Err.Number <> 0 Then
            Debug.Print "  ERROR processing: " & Err.Number & " - " & Err.Description
            Err.Clear
        End If

        fileName = Dir()
    Loop
    Debug.Print "Total .frm files found: " & frmFileCount

    Debug.Print "Total imported: " & importCount & " module(s)"
    Debug.Print "--- ImportMacrosFromFolder END ---"
End Sub

' Process a single import (Git-aware: folder is source of truth)
Function ProcessImport(filePath As String, moduleName As String, fileExt As String) As Boolean
    Debug.Print "  -> ProcessImport START for: " & moduleName & fileExt
    On Error Resume Next

    Dim vbComp As Object
    Dim moduleExists As Boolean
    Dim filesIdentical As Boolean
    Dim tempExportPath As String

    ProcessImport = False ' Default to False

    ' Check if module already exists in Normal.dotm
    Debug.Print "  -> Checking if module already exists in Normal.dotm..."
    moduleExists = False
    For Each vbComp In Application.NormalTemplate.VBProject.VBComponents
        If vbComp.Name = moduleName Then
            moduleExists = True
            Debug.Print "  -> Module EXISTS in Normal.dotm: " & moduleName
            Exit For
        End If
    Next vbComp

    If Not moduleExists Then
        Debug.Print "  -> Module does NOT exist in Normal.dotm (will import new module)"
    End If

    ' If module exists, check if files are identical (optimization to skip unnecessary imports)
    If moduleExists Then
        Debug.Print "  -> Comparing with existing module..."
        ' Export current version to a temp file for comparison
        tempExportPath = GetSyncFolderPath() & "~temp_" & moduleName & fileExt
        Debug.Print "  -> Exporting current version to temp file: " & tempExportPath

        vbComp.Export tempExportPath
        If Err.Number <> 0 Then
            Debug.Print "  -> ERROR exporting to temp file: " & Err.Number & " - " & Err.Description
            Err.Clear
            Exit Function
        End If

        ' Compare the two files
        Debug.Print "  -> Comparing files..."
        filesIdentical = FilesAreIdentical(filePath, tempExportPath)

        ' Delete temp file
        Debug.Print "  -> Deleting temp file..."
        Kill tempExportPath
        If Err.Number <> 0 Then
            Debug.Print "  -> ERROR deleting temp file: " & Err.Number & " - " & Err.Description
            Err.Clear
        End If

        ' If files are identical, skip import
        If filesIdentical Then
            Debug.Print "  -> Files are IDENTICAL - skipping import (already in sync)"
            ProcessImport = False
            Exit Function
        Else
            Debug.Print "  -> Files are DIFFERENT - will import folder version (Git-managed)"
        End If

        ' Files differ: Remove existing module to import new version
        Debug.Print "  -> Removing existing module from Normal.dotm..."
        Application.NormalTemplate.VBProject.VBComponents.Remove vbComp
        If Err.Number <> 0 Then
            Debug.Print "  -> ERROR removing module: " & Err.Number & " - " & Err.Description
            Err.Clear
            Exit Function
        End If
        Debug.Print "  -> Module removed successfully"
    End If

    ' Import the module from file (folder is source of truth)
    Debug.Print "  -> Importing module from file: " & filePath
    Application.NormalTemplate.VBProject.VBComponents.Import filePath
    If Err.Number <> 0 Then
        Debug.Print "  -> ERROR importing module: " & Err.Number & " - " & Err.Description
        Err.Clear
        Exit Function
    End If

    Debug.Print "  -> Import SUCCESS: " & moduleName & fileExt
    ProcessImport = True
    Debug.Print "  -> ProcessImport END (success)"
End Function

' ========================================================================
' HELPER FUNCTIONS
' ========================================================================

' Compare two files to see if they're identical
Function FilesAreIdentical(file1 As String, file2 As String) As Boolean
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
' MANUAL TRIGGER SUBS (Optional - for testing)
' ========================================================================

' You can run these manually to test the sync without restarting Word
Sub ManualExport()
    ExportMacrosToFolder
    MsgBox "Manual export complete!", vbInformation
End Sub

Sub ManualImport()
    ImportMacrosFromFolder
    MsgBox "Manual import complete!", vbInformation
End Sub






