# Word VBA Template Management System
## Installation and Usage Instructions

## Overview
This system automatically loads all your .dotm and .dot template files at Word startup, hides their windows, and makes their VBA projects visible in the VBA Editor for easy editing.

---

## Installation Steps

### 1. Create/Update Utilities.dotm

**Step 1a: Create the Template File**
1. Open Word
2. Go to **File → Save As**
3. Choose **Word Macro-Enabled Template (*.dotm)** as the file type
4. Name it `Utilities.dotm`
5. Save it in your Templates folder (typically `%APPDATA%\Microsoft\Templates`)

**Step 1b: Import the Code Modules into Utilities.dotm**
1. Press `Alt+F11` to open the VBA Editor
2. In the Project Explorer, locate **Utilities (Utilities.dotm)**
3. Right-click on **Utilities** → **Import File...**
4. Browse to and select **TemplateViewer.bas**
5. Click **Open**
6. Right-click on **Utilities** → **Import File...**
7. Browse to and select **TemplateViewerAppEvents.cls**
8. Click **Open**

You should now see:
- ** Project (Utilities)**
  - **Modules**
    - TemplateViewer
  - **Class Modules**
    - TemplateViewerAppEvents

**Step 1c: Save Utilities.dotm**
1. Press `Ctrl+S` to save
2. Close the VBA Editor

---

### 2. Install Utilities.dotm as a Global Add-in

1. In Word, go to **File → Options → Add-ins**
2. At the bottom, select **Manage: Word Add-ins** → click **Go...**
3. Click **Add...**
4. Browse to your Templates folder and select `Utilities.dotm`
5. Ensure the checkbox next to **Utilities** is checked
6. Click **OK**

Utilities.dotm will now load automatically every time Word starts.

---

### 3. Add AutoExec to Normal.dotm

**Locate or create your AutoExec macro:**
1. Press `Alt+F11` to open the VBA Editor
2. In the Project Explorer, locate **Normal (Normal.dotm)**
3. If you have an existing AutoExec macro, find it in your modules
4. If not, insert a new standard module

**Add the LoadTemplates call:**

```vba
Public Sub AutoExec()
    ' Your existing AutoExec code here (if any)
    ' ...

    ' Load and hide template files for VBA Editor access
    On Error Resume Next
    Application.Run "Utilities.LoadTemplates"
    If Err.Number <> 0 Then
        MsgBox "Could not load templates. Make sure Utilities.dotm is installed as an add-in.", _
               vbExclamation, "Template Manager"
    End If
    On Error GoTo 0
End Sub
```

**Note:** The `Application.Run` method calls the LoadTemplates procedure from the Utilities.dotm add-in.

---

### 4. Save Normal.dotm and Test

1. In the VBA Editor, press `Ctrl+S` or click **File → Save Normal**
2. Close Word completely (ensure all Word windows are closed)
3. Restart Word to test the system

---

## How It Works

### At Startup:
1. Word starts and loads Normal.dotm
2. Word loads Utilities.dotm as a global add-in (configured in Word Add-ins)
3. AutoExec() in Normal.dotm runs automatically
4. AutoExec calls `Application.Run "Utilities.LoadTemplates"` to invoke the macro from Utilities.dotm
5. `LoadTemplates()` scans your Templates folder (`%APPDATA%\Microsoft\Templates`)
6. Opens all .dotm and .dot files (except Normal.dotm/Normal.dot)
7. Hides their windows (they're invisible but remain open)
8. Shows a summary message of how many templates loaded

### During Your Session:
- All template VBA projects are visible in the VBA Editor's Project Explorer
- You can navigate, edit, and save any macro in any template
- All macros from all templates remain globally available
- Hidden templates stay loaded even when you close regular documents
- Module-level variables in Utilities.dotm persist throughout the Word session

### At Shutdown:
- When you close Word (File → Exit), the Application_Quit event fires
- TemplateViewerAppEvents (in Utilities.dotm) detects this event
- The system automatically closes hidden templates via CloseTemplates()
- If any template has unsaved changes, Word prompts you to save (standard save dialog)
- Templates close cleanly with no orphaned processes

---

## Manual Control

### Manually Close All Templates
You can close hidden templates without restarting Word:

**Option 1: Run from VBA Editor**
1. Press `Alt+F11` to open VBA Editor
2. Press `Ctrl+G` to open Immediate Window
3. Type: `Utilities.ManualCloseTemplates` and press Enter

**Option 2: Assign to Keyboard Shortcut**
1. File → Options → Customize Ribbon
2. Click **Keyboard shortcuts: Customize...**
3. Categories: **Macros**
4. Macros: Select **Utilities.TemplateViewer.ManualCloseTemplates**
5. Press new shortcut key (e.g., `Alt+Shift+U`)
6. Click **Assign**, then **Close**

**Option 3: Add to Quick Access Toolbar**
1. Right-click Quick Access Toolbar → **Customize Quick Access Toolbar...**
2. Choose commands from: **Macros**
3. Select **Utilities.TemplateViewer.ManualCloseTemplates**
4. Click **Add >>**, then **OK**

---

## Troubleshooting

### Templates Not Loading
**Check if Utilities.dotm is loaded:**
1. Go to **File → Options → Add-ins**
2. Look for "Utilities" in the Active Application Add-ins list
3. If not there, click **Manage: Word Add-ins → Go...** and add it (see Step 2 above)
4. Ensure the checkbox next to "Utilities" is checked

**Check the Immediate Window for errors:**
1. Press `Alt+F11` to open VBA Editor
2. Press `Ctrl+G` to open Immediate Window
3. Look for error messages from the Template Manager

**Common issues:**
- **"Could not load templates"** - Utilities.dotm is not installed as a global add-in
- **"Unable to determine Templates folder path"** - Your Templates folder path couldn't be found
- **"(not found)"** - Template file was deleted or moved
- **"(file in use)"** - Template is already open in another Word instance
- **"(corrupted or invalid)"** - Template file is damaged

### VBA Projects Not Visible
- Make sure macro security allows macros: File → Options → Trust Center → Trust Center Settings → Macro Settings
- Choose "Disable all macros with notification" or "Enable all macros"
- Restart Word after changing security settings

### Templates Appear in Taskbar/Alt+Tab
This shouldn't happen, but if it does:
- The `.Visible = False` setting should hide windows
- Check if you have any custom add-ins that might interfere with window visibility

### Word Won't Close / Hangs on Exit
- Check Immediate Window for errors during shutdown
- One of your templates might have an issue preventing clean closing
- Try `ManualCloseTemplates` manually before closing Word

### Performance Issues
- The system is designed for up to 20 templates
- Each template uses ~2-5 MB RAM
- If you have many large templates, consider organizing them and loading only essential ones

---

## File Locations

### Your Templates Folder:
```
C:\Users\[YourUsername]\AppData\Roaming\Microsoft\Templates
```

### What Gets Loaded:
- ✅ All .dotm files (Word 2007+ macro-enabled templates)
- ✅ All .dot files (Word 97-2003 legacy templates)
- ❌ Normal.dotm (excluded - already loaded by Word)
- ❌ Normal.dot (excluded - legacy Normal template)

---

## Debugging

### View Detailed Logs
All operations are logged to the Immediate Window:

1. Open VBA Editor (`Alt+F11`)
2. Open Immediate Window (`Ctrl+G`)
3. Review load/close operations and any errors

**Example log output:**
```
=== Template Manager: Loading Templates ===
Templates Path: C:\Users\Alex\AppData\Roaming\Microsoft\Templates\
Timestamp: 11/24/2025 10:30:45 AM
Loading: MyMacros.dotm
  Status: Loaded and hidden successfully
Loading: ProjectTools.dotm
  Status: Loaded and hidden successfully
---
Templates loaded successfully: 2
Templates failed to load: 0
==========================================
```

---

## Uninstalling

To remove the system:

1. **Remove AutoExec call from Normal.dotm:**
   - Open VBA Editor (`Alt+F11`)
   - Find your AutoExec macro in Normal.dotm
   - Delete or comment out the lines: `Application.Run "Utilities.LoadTemplates"` and related error handling
   - Press `Ctrl+S` to save Normal.dotm

2. **Remove Utilities.dotm as a global add-in:**
   - Go to **File → Options → Add-ins**
   - Select **Manage: Word Add-ins → Go...**
   - Uncheck **Utilities** or select it and click **Remove**
   - Click **OK**

3. **Delete Utilities.dotm file (optional):**
   - Navigate to your Templates folder (`%APPDATA%\Microsoft\Templates`)
   - Delete **Utilities.dotm**
   - This will permanently remove the TemplateViewer and TemplateViewerAppEvents code

---

## Version Control

### Backing Up Your Code
Export modules from Utilities.dotm before making changes:
1. Open VBA Editor (`Alt+F11`)
2. Expand **Utilities (Utilities.dotm)** in Project Explorer
3. Right-click **TemplateViewer** → **Export File...**
4. Save to a backup location with date in filename (e.g., `TemplateViewer_2025-11-24.bas`)
5. Repeat for **TemplateViewerAppEvents** class module

### Restoring from Backup
1. Open VBA Editor (`Alt+F11`)
2. In Utilities.dotm project, right-click → **Import File...**
3. Select your backup .bas or .cls file
4. If module already exists, remove it first then import

---

## Advanced Configuration

### Loading Templates from Other Folders
By default, the system loads templates from the standard Templates folder. To load from additional locations, modify the `LoadTemplates()` procedure:

```vba
' After loading from standard folder, load from custom folder:
templatesPath = "C:\MyCustomTemplates\"
fileName = Dir(templatesPath & "*.dotm")
' ... (same loading logic)
```

### Excluding Specific Templates
To skip certain templates, add exclusions in the `LoadTemplates()` procedure:

```vba
' Skip specific files
If LCase(strFileName) = "oldtemplate.dotm" Then
    strFileName = Dir()
    Continue Do ' Skip to next file
End If
```

### Load Order
Templates load in the order returned by `Dir()`, which is typically alphabetical. To enforce a specific order, you could:
1. Rename templates with numeric prefixes (01_First.dotm, 02_Second.dotm)
2. Or modify the code to use an array of specific filenames in your desired order

---

## Architecture Notes

### Why Separate Utilities.dotm from Normal.dotm?

The system is designed with a two-template architecture for several important reasons:

1. **Keeps Normal.dotm Clean**
   - Normal.dotm only contains a simple AutoExec call
   - Minimal code footprint in your core Word template
   - Reduces risk of corrupting Normal.dotm

2. **Module-Level Variables Persist**
   - The `m_colOpenTemplates` collection and `m_objAppEvents` in TemplateViewer.bas need to persist throughout the Word session
   - When code runs from a global add-in (Utilities.dotm), module-level variables remain in memory
   - This ensures templates can be tracked and closed properly

3. **Easier Maintenance and Updates**
   - Update Utilities.dotm without modifying Normal.dotm
   - Can be distributed to other users as a single .dotm file
   - Easy to enable/disable via Word Add-ins without code changes

4. **Better Organization**
   - Utilities.dotm can contain other utility macros beyond just Template Manager
   - Separates general utilities from user-specific Normal.dotm customizations
   - Professional development practice (separation of concerns)

### How AutoExec Calls Across Templates

The `Application.Run` method allows VBA code in one template to call public procedures in another:

```vba
Application.Run "TemplateName.ProcedureName"
```

In our case:
- **Normal.dotm**: Contains `AutoExec()` which runs at startup
- **Utilities.dotm**: Contains `LoadTemplates()` in the TemplateViewer module
- **Connection**: `Application.Run "Utilities.LoadTemplates"` bridges them

This pattern is standard for Word add-in development and ensures clean separation between startup logic and utility functions.

---

## Support

### What to Check When Issues Occur:
1. ✅ Utilities.dotm is installed as a global add-in (File → Options → Add-ins)
2. ✅ Macros are enabled (Trust Center settings)
3. ✅ AutoExec in Normal.dotm has correct Application.Run call
4. ✅ Templates folder path is correct
5. ✅ Template files are not corrupt
6. ✅ You have read/write permissions to Templates folder
7. ✅ No other Word instances have templates open
8. ✅ Check Immediate Window for error details

### Error Codes Reference:
- **5174** - File not found or path invalid
- **5941** - File is locked or in use by another process
- **5152** - File is corrupted or invalid format

---

## System Requirements

- Microsoft Word 2013 or later
- Windows operating system
- Macros enabled in Trust Center
- 4+ GB RAM (for multiple templates)
- Read/write access to Templates folder

---

**Document Version:** 1.0  
**Date:** 2025-11-24  
**Compatibility:** Word 2013, 2016, 2019, 2021, Microsoft 365
