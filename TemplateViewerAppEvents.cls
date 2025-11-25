VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END

' This file used to be named clsAppEvents_v2.cls
' It has been renamed to TemplateViewerAppEvents.cls

Attribute VB_Name = "TemplateViewerAppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ============================================================================
' Class: TemplateViewerAppEvents
' Purpose: Application event handler for Word Template Management System
'          Intercepts the Application.Quit event to close hidden templates
'          gracefully before Word shuts down
' ============================================================================

' WithEvents allows this class to respond to Application events
Private WithEvents m_wdApp As Word.Application

' ============================================================================
' Class_Initialize
' Purpose: Automatically sets up event handling when class is instantiated
' ============================================================================
Private Sub Class_Initialize()
    ' Connect to the Word Application object to monitor its events
    Set m_wdApp = Word.Application
    Debug.Print "TemplateViewerAppEvents: Event handler initialized"
End Sub

' ============================================================================
' Class_Terminate
' Purpose: Cleanup when class is destroyed
' ============================================================================
Private Sub Class_Terminate()
    ' Disconnect from Application events
    Set m_wdApp = Nothing
    Debug.Print "TemplateViewerAppEvents: Event handler terminated"
End Sub

' ============================================================================
' m_wdApp_Quit
' Purpose: Event handler triggered when Word application is quitting
'          Calls CloseTemplates() to properly close all hidden templates
'          and prompt user to save any unsaved changes
' ============================================================================
Private Sub m_wdApp_Quit()
    On Error Resume Next
    Debug.Print "TemplateViewerAppEvents: Application.Quit event detected"

    ' Call the CloseTemplates procedure from TemplateViewer module
    TemplateViewer.CloseTemplates

    If Err.Number <> 0 Then
        Debug.Print "TemplateViewerAppEvents: Error in Quit handler - " & Err.Description
        Err.Clear
    End If
End Sub
