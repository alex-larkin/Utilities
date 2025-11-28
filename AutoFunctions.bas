Attribute VB_Name = "AutoFunctions"

Sub AutoExec()
    Debug.Print "=== AutoExec START ==="
    Debug.Print "Current Time: " & Now
    Debug.Print "Word Version: " & Application.VERSION
    Debug.Print "Normal Template Path: " & Application.NormalTemplate.FullName
    On Error Resume Next
    
    ' Don't try to update yourself
     'Application.Run "'VBAMacroSync.dotm'!VBAMacroSync.VMS_AutoExec"
   
    If Err.Number <> 0 Then
    Debug.Print "WARNING: VMS_AutoExec failed - " & Err.Description
    End If

    On Error GoTo 0

    Debug.Print "=== AutoExec END ==="

End Sub

Sub AutoExit()
    Debug.Print "=== AutoExit START ==="
    Debug.Print "Current Time: " & Now
    On Error Resume Next
    
'    Application.Run "VBAMacroSync.VMS_AutoExit"
    
    If Err.Number <> 0 Then
        Debug.Print "WARNING: VMS_AutoExit failed - " & Err.Description
        MsgBox "Error during AutoExit: " & Err.Description, vbCritical
    End If
    
    On Error GoTo 0
    
    Debug.Print "=== AutoExit END ==="
End Sub
