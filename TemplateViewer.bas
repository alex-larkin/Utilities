Attribute VB_Name = "TemplateViewer"
Option Explicit

' This File used to be named OpenAllTemplates.bas
' It has been renamed to TemplateViewer.bas

' ============================================================================
' Module: TemplateViewer
' Purpose: Automatically load and hide all Word template files (.dotm/.dot)
'          to make their VBA projects visible in the VBA Editor
' ============================================================================

' Module-level variables to persist throughout Word session
Private m_colOpenTemplates As Collection      ' Tracks all opened template documents
Private m_objAppEvents As TemplateViewerAppEvents     ' Application event handler (prevents garbage collection)

' ============================================================================
' LoadTemplates
' Purpose: Opens all .dotm and .dot template files from the Templates folder,
'          hides their windows, and makes their VBA projects accessible
' Called by: AutoExec() in Normal.dotm at Word startup
' ============================================================================
Public Sub LoadTemplates()
    Debug.Print "=== Starting LoadTemplates() ==="
    
    On Error GoTo ErrorHandler

    Dim strTemplatesPath As String
    Dim strFileName As String
    Dim docTemplate As Document
    Dim intCount As Integer

    ' Initialize the collection if not already created
    If m_colOpenTemplates Is Nothing Then
        Set m_colOpenTemplates = New Collection
    End If

    ' Get the Templates folder path
    On Error Resume Next
    strTemplatesPath = Application.Options.DefaultFilePath(wdUserTemplatesPath)
    If Err.Number <> 0 Or strTemplatesPath = "" Then
        ' Fallback to environment variable
        strTemplatesPath = Environ("APPDATA") & "\Microsoft\Templates"
        Err.Clear
    End If
    On Error GoTo ErrorHandler

    ' Ensure path ends with backslash
    If Right(strTemplatesPath, 1) <> "\" Then
        strTemplatesPath = strTemplatesPath & "\"
    End If

    Debug.Print "=== Template Loading Started ==="
    Debug.Print "Templates Path: " & strTemplatesPath
    Debug.Print "Timestamp: " & Now

    intCount = 0

    ' Enumerate .dotm files first
    strFileName = Dir(strTemplatesPath & "*.dotm")
    Do While strFileName <> ""
        ' Skip Normal.dotm (already loaded by Word)
        If LCase(strFileName) <> "normal.dotm" Then
            If OpenAndHideTemplate(strTemplatesPath & strFileName, docTemplate) Then
                m_colOpenTemplates.Add docTemplate
                intCount = intCount + 1
            End If
        End If
        strFileName = Dir()
    Loop

    ' Enumerate .dot files (legacy format)
    strFileName = Dir(strTemplatesPath & "*.dot")
    Do While strFileName <> ""
        ' Skip Normal.dot if it exists
        If LCase(strFileName) <> "normal.dot" Then
            If OpenAndHideTemplate(strTemplatesPath & strFileName, docTemplate) Then
                m_colOpenTemplates.Add docTemplate
                intCount = intCount + 1
            End If
        End If
        strFileName = Dir()
    Loop

    ' Initialize Application event handler to detect Word shutdown
    If m_objAppEvents Is Nothing Then
        Set m_objAppEvents = New TemplateViewerAppEvents
    End If

    ' Log success
    Debug.Print "Templates Loaded: " & intCount
    Debug.Print "=== Template Loading Complete ==="

    ' Notify user (optional - comment out if too intrusive)
    If intCount > 0 Then
        Debug.Print "Successfully loaded " & intCount & " template(s). All VBA projects are now accessible in the VBA Editor."
    End If

    Exit Sub

ErrorHandler:
    Dim strError As String
    strError = "Error in LoadTemplates: " & Err.Number & " - " & Err.Description
    Debug.Print strError
    MsgBox strError, vbExclamation, "Template Loading Error"
    Resume Next
End Sub

' ============================================================================
' OpenAndHideTemplate
' Purpose: Opens a single template file and hides its window
' Parameters:
'   strFilePath - Full path to template file
'   docTemplate - Output parameter: Document object of opened template
' Returns: True if successful, False if error occurred
' ============================================================================
Private Function OpenAndHideTemplate(ByVal strFilePath As String, ByRef docTemplate As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim strFileName As String
    strFileName = Dir(strFilePath)

    Debug.Print "Opening: " & strFileName

    ' Open the template file for editing (not as a template to create new docs)
    Set docTemplate = Documents.Open(FileName:=strFilePath, _
                                     ReadOnly:=False, _
                                     AddToRecentFiles:=False, _
                                     Visible:=False)

    ' Hide the template window
    If Not docTemplate Is Nothing Then
        If Not docTemplate.ActiveWindow Is Nothing Then
            docTemplate.ActiveWindow.Visible = False
        End If
        Debug.Print "  -> Successfully opened and hidden"
        OpenAndHideTemplate = True
    End If

    Exit Function

ErrorHandler:
    Dim strError As String
    strError = "Error opening '" & strFileName & "': " & Err.Number & " - " & Err.Description
    Debug.Print "  -> ERROR: " & strError

    ' Show user-friendly error message
    Select Case Err.Number
        Case 5174 ' File is locked or in use
            MsgBox "Cannot open '" & strFileName & "' - file is locked or in use by another process.", vbExclamation, "Template Loading Warning"
        Case 5273, 53 ' File not found
            MsgBox "Cannot find template file: '" & strFileName & "'", vbExclamation, "Template Loading Warning"
        Case Else
            MsgBox strError, vbExclamation, "Template Loading Warning"
    End Select

    OpenAndHideTemplate = False
End Function

' ============================================================================
' CloseTemplates
' Purpose: Closes all hidden templates, prompting to save if modified
' Called by: TemplateViewerAppEvents.Application_Quit event handler
' ============================================================================
Public Sub CloseTemplates()
    On Error Resume Next ' Continue closing even if errors occur

    Dim docTemplate As Document
    Dim intCount As Integer
    Dim strError As String

    Debug.Print "=== Template Closing Started ==="
    Debug.Print "Timestamp: " & Now

    If Not m_colOpenTemplates Is Nothing Then
        intCount = 0

        ' Close each template in the collection
        For Each docTemplate In m_colOpenTemplates
            If Not docTemplate Is Nothing Then
                Debug.Print "Closing: " & docTemplate.Name

                ' Close with save prompt (Word handles the save dialog automatically)
                docTemplate.Close SaveChanges:=wdPromptToSaveChanges

                If Err.Number <> 0 Then
                    strError = "Error closing '" & docTemplate.Name & "': " & Err.Number & " - " & Err.Description
                    Debug.Print "  -> ERROR: " & strError
                    Err.Clear
                Else
                    Debug.Print "  -> Closed successfully"
                    intCount = intCount + 1
                End If
            End If
        Next docTemplate

        ' Clear the collection
        Set m_colOpenTemplates = Nothing

        Debug.Print "Templates Closed: " & intCount
    Else
        Debug.Print "No templates to close (collection is Nothing)"
    End If

    ' Clean up event handler
    Set m_objAppEvents = Nothing

    Debug.Print "=== Template Closing Complete ==="
End Sub

' ============================================================================
' ManualCloseTemplates (Optional utility)
' Purpose: Allows user to manually close templates without quitting Word
' Usage: Can be called from Ribbon button or keyboard shortcut
' ============================================================================
Public Sub ManualCloseTemplates()
    Dim intResponse As VbMsgBoxResult

    intResponse = MsgBox("Close all hidden templates?" & vbCrLf & vbCrLf & _
                        "You will be prompted to save any templates with unsaved changes." & vbCrLf & _
                        "VBA projects will no longer be visible in the VBA Editor until you restart Word.", _
                        vbQuestion + vbYesNo, "Close Hidden Templates")

    If intResponse = vbYes Then
        CloseTemplates
        MsgBox "Hidden templates have been closed.", vbInformation, "Templates Closed"
    End If
End Sub

' ============================================================================
' GetLoadedTemplateCount (Optional utility)
' Purpose: Returns the number of currently loaded hidden templates
' Usage: For debugging or status display
' ============================================================================
Public Function GetLoadedTemplateCount() As Integer
    If Not m_colOpenTemplates Is Nothing Then
        GetLoadedTemplateCount = m_colOpenTemplates.Count
    Else
        GetLoadedTemplateCount = 0
    End If
End Function
