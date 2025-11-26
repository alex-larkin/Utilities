# VBA Macro Sync System

Git-aware two-way synchronization between Word template macros and local folders for collaborative development. Supports multiple templates (Normal.dotm, Utilities.dotm, custom templates) with automatic discovery and sync.

## Architecture Overview

**Centralized Design:**
- VBAMacroSync.bas resides in **Utilities.dotm** as the central sync engine
- Each template implements simple AutoExec/AutoExit hooks to trigger sync
- Templates are discovered automatically via environment variables (`MACROS_<TEMPLATENAME>`)
- No code duplication—one sync module serves all templates

**For complete technical specification, see:** [VBAMAcroSync_SRD.md](VBAMAcroSync_SRD.md)

## Quick Start

1. **Enable VBA Project Access** (one-time setup):
   - Word → File → Options → Trust Center → Trust Center Settings
   - Enable "Trust access to the VBA project object model"

2. **Set Environment Variables** (one-time setup):
   - Press `Windows + R` to open Run dialog
   - Type `sysdm.cpl` and press Enter
   - Click the **"Advanced"** tab
   - Click **"Environment Variables"** button at the bottom
   - Under "User variables" (top section), click **"New..."** for each template you want to sync
   - Create variables using the pattern `MACROS_<TEMPLATENAME>`:
     - For Normal.dotm:
       - **Variable name:** `MACROS_NORMAL`
       - **Variable value:** `C:\Your\Path\To\Macros\Normal\` (ending with `\`)
     - For Utilities.dotm:
       - **Variable name:** `MACROS_UTILITIES`
       - **Variable value:** `C:\Your\Path\To\Macros\Utilities\` (ending with `\`)
     - For custom templates, use the same pattern (e.g., `MACROS_MYTHEME` for MyTheme.dotm)
   - Click OK on all dialogs
   - Restart any open applications for the change to take effect

3. **Set Up VBAMacroSync Module** (one-time setup):
   - VBAMacroSync.bas should reside in **Utilities.dotm** (the central sync engine)
   - Each template that wants sync must implement simple AutoExec/AutoExit hooks that call the centralized VMS_AutoExec() and VMS_AutoExit() functions
   - See VBAMAcroSync_SRD.md for complete implementation details
   - Close and reopen Word to activate automatic sync

## How It Works

**Multi-Template Architecture:**
- VBAMacroSync.bas resides in **Utilities.dotm** and provides sync services for all templates
- Each template that wants sync implements simple AutoExec/AutoExit hooks
- On Word startup, VMS_AutoExec() automatically discovers and syncs all configured templates
- Templates are identified by environment variables (e.g., `MACROS_NORMAL`, `MACROS_UTILITIES`)

**Sync Behavior:**
- **On Word Open:** Imports all .bas/.cls/.frm files from each template's folder → template
- **On Word Close:** Exports all modules from each template → corresponding folder

**Git-Aware Design:**
   - When opening Word, folder is source of truth after Git operations (pull/push)
   - When closing Word, template is source of truth—changes overwrite folder contents, ready for Git commit
   - IMPORTANT: No conflict detection in macros. Git handles merge conflicts.

   (Note: Don't edit macros in external editors like VS Code while Word is open—VMS_AutoExit() will overwrite changes on Word close)

## Daily Workflow with GitHub Desktop

1. **Morning:** Pull latest changes (GitHub Desktop)
2. **Open Word:** Macros auto-import from folder
3. **Edit macros** in Word VBA Editor during the day
4. **Close Word:** Macros auto-export to folder
5. **Review changes** in GitHub Desktop
6. **Commit and push** your changes
7. **If Git conflicts occur:** Resolve in GitHub Desktop's merge tool, then reopen Word

## Editing .bas Files in VS Code

You can edit .bas files directly in VS Code while Word is closed:

1. Edit .bas file in VS Code
2. Save changes (Ctrl+S)
3. Commit to Git via GitHub Desktop
4. Open Word → changes automatically import

**Important:**
- Preserve the `Attribute VB_Name = "ModuleName"` header line
- Use CRLF line endings (Windows format)
   - To check this, open a .bas file in VSC and click somewhere in the file. In the lower-right corner you should see "CRLF". If you see "LF", click "LF" and then select CRLF from the menu that appears at the top of the screen.
- For special characters, save as ANSI/Windows-1252 encoding

## Deleting Modules

Deletions are **not** automatically synced. To delete a module completely:

1. Delete from the template (VBA Editor)
2. Delete corresponding .bas/.cls/.frm file from the sync folder
3. Commit deletion to Git

**Note:** If you only delete the file from the folder, it will reappear on Word close (exported from template). If you only delete from the template, it will reappear on Word open (imported from folder).

## Manual Testing

Run these macros in VBA Editor for immediate sync without restarting Word:

- `ManualExport` - Export all configured templates → their folders
- `ManualImport` - Import all configured templates from their folders

**Location:** Both macros are in Utilities.dotm's VBAMacroSync module

**Important:** Close the VBA editor before running these macros. Otherwise Word will create duplicate modules with "1" appended to their names.

View debug output in Immediate Window (Ctrl+G in VBA Editor). The output shows which templates were scanned and synced.

## File Types Supported

- `.bas` - Standard modules
- `.cls` - Class modules
- `.frm` - UserForms

## Troubleshooting

**Macros not importing on Word open:**
- Check Immediate Window (Ctrl+G) for debug messages
- Verify VBA project access is enabled
- Confirm environment variables are set correctly (e.g., `MACROS_NORMAL`, `MACROS_UTILITIES`)
- Ensure environment variable names match template names (uppercase, without .dotm extension)
- Check that Utilities.dotm is loaded and contains VBAMacroSync.bas
- Verify AutoExec/AutoExit hooks are implemented in the template

**Import failed after editing in VS Code:**
- Verify `Attribute VB_Name` matches filename
- Check line endings are CRLF (not LF)
- Run `ManualImport` to see detailed error messages

**Changes not syncing:**
- When opening Word, folder is source of truth—Git changes always override template
- When closing Word, template is source of truth—changes in template always override folder
- If files are identical, import is skipped (optimization)
- Check that you're editing the correct sync folder for that template
- Verify the template has an environment variable configured
- Check that AutoExec/AutoExit hooks are calling VMS_AutoExec/VMS_AutoExit

## Configuration

Sync folder paths are configured via environment variables using the pattern `MACROS_<TEMPLATENAME>`:
- `MACROS_NORMAL` - for Normal.dotm
- `MACROS_UTILITIES` - for Utilities.dotm
- `MACROS_<CUSTOMNAME>` - for any custom template

**Each user sets their own local paths** - paths are not stored in code or Git repository. This keeps your personal folder structure private.

**Automatic Template Discovery:** VBAMacroSync automatically scans all loaded templates and syncs those with configured environment variables. No code changes needed to add new templates—just set the environment variable.

**Recommendation:** Don't use auto-syncing cloud folders (Dropbox, OneDrive). Use Git for version control instead.

## Syncing Multiple Templates

The refactored VBAMacroSync now has **built-in multi-template support**. Adding new templates is simple:

1. **Create environment variable for the new template**:
   - Follow Quick Start step 2
   - Use the naming pattern: `MACROS_<TEMPLATENAME>`
   - Example: For MyCustom.dotm, create `MACROS_MYCUSTOM` → `C:\Path\To\Macros\MyCustom\`

2. **Add AutoExec/AutoExit hooks in the template**:
   - Each template needs simple hooks that call the centralized sync functions
   - Example code for your template:
   ```vba
   ' In MyCustom.dotm
   Public Sub AutoExec()
       On Error Resume Next
       Application.Run "Utilities.VBAMacroSync.VMS_AutoExec"
   End Sub

   Public Sub AutoExit()
       On Error Resume Next
       Application.Run "Utilities.VBAMacroSync.VMS_AutoExit"
   End Sub
   ```

3. **Restart Word** - the new template will be automatically discovered and synced

**Key Benefits:**
- VBAMacroSync.bas stays in one place (Utilities.dotm)
- No code duplication across templates
- Easy to update—change VBAMacroSync once, all templates benefit
- Templates are discovered automatically based on environment variables
