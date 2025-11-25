# VBA Macro Sync System

Git-aware two-way synchronization between Word Normal.dotm macros and a local folder for collaborative development.

## Quick Start

1. **Enable VBA Project Access** (one-time setup):
   - Word → File → Options → Trust Center → Trust Center Settings
   - Enable "Trust access to the VBA project object model"

2. **Set Environment Variable** (one-time setup):
   - Press `Windows + R` to open Run dialog
   - Type `sysdm.cpl` and press Enter
   - Click the **"Advanced"** tab
   - Click **"Environment Variables"** button at the bottom
   - Under "User variables" (top section), click **"New..."**
   - Enter:
     - **Variable name:** `VBA_MACRO_SYNC_PATH`
     - **Variable value:** `C:\Your\Path\To\SyncFolder\` (use your actual path, ending with `\`)
   - Click OK on all dialogs
   - Restart any open applications for the change to take effect

3. **Import into Normal.dotm**:
   - Open Word VBA Editor (Alt+F11)
   - Import VBAMacroSync.bas into Normal template
   - Close and reopen Word to activate

## How It Works

**On Word Open:** Imports all .bas/.cls/.frm files from folder → Normal.dotm
**On Word Close:** Exports all modules from Normal.dotm → folder

**Git-Aware Design:** 
   - Folder is source of truth after Git operations (pull/push). 
   - IMPORTANT: These macros have no conflict detection. Git handles merge conflicts.

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
   - To check this, open a .bas file in VSC and click somewhere in the file. In the lower-right corder you should see "CRFL". If you see "LF", click "LF" and then select CRFL from the menu tat appears at the top of the screen.
- For special characters, save as ANSI/Windows-1252 encoding

## Deleting Modules

Deletions are **not** automatically synced. To delete a module completely:

1. Delete from Normal.dotm (VBA Editor)
2. Delete corresponding .bas/.cls/.frm file from folder
3. Commit deletion to Git

**Note:** If you only delete the file from the folder, it will reappear on Word close (exported from Normal.dotm). If you only delete from Normal.dotm, it will reappear on Word open (imported from folder).

## Manual Testing

Run these macros in VBA Editor for immediate sync without restarting Word:

- `ManualExport` - Export Normal.dotm → folder
- `ManualImport` - Import folder → Normal.dotm

Important: close the VBA editor before running these macros. Otherwise Word will create duplicate modules of your existing VBA modules with the number 1 appended to them.

View debug output in Immediate Window (Ctrl+G in VBA Editor).

## File Types Supported

- `.bas` - Standard modules
- `.cls` - Class modules
- `.frm` - UserForms

## Troubleshooting

**Macros not importing on Word open:**
- Check Immediate Window (Ctrl+G) for debug messages
- Verify VBA project access is enabled
- Confirm `VBA_MACRO_SYNC_PATH` environment variable is set correctly

**Import failed after editing in VS Code:**
- Verify `Attribute VB_Name` matches filename
- Check line endings are CRLF (not LF)
- Run `ManualImport` to see detailed error messages

**Changes not syncing:**
- Folder is source of truth—Git changes always override Normal.dotm
- If files are identical, import is skipped (optimization)
- Check that you're editing the correct sync folder

## Configuration

The sync folder path is configured via the `VBA_MACRO_SYNC_PATH` environment variable (see Quick Start, step 2).

**Each user sets their own local path** - the path is not stored in the code or Git repository. This keeps your personal folder structure private.

**Recommendation:** Don't use auto-syncing cloud folders (Dropbox, OneDrive). Use Git for version control instead.

## Syncing Multiple Templates

If you have additional .dotm or .dot template files beyond Normal.dotm, you can sync each independently:

1. **Copy VBAMacroSync.bas into each template** (via VBA Editor → File → Import)
2. **Create a separate environment variable for each template**:
   - Each template needs its own unique environment variable
   - Follow the same steps as Quick Start (step 2) for each variable:
     - Example variables:
       - `VBA_MACRO_SYNC_PATH_NORMAL` → `C:\Your\Path\WordMacros\Normal\`
       - `VBA_MACRO_SYNC_PATH_OTHER` → `C:\Your\Path\WordMacros\OtherTemplate\`
3. **Modify the `GetSyncFolderPath()` function** in each template's copy of VBAMacroSync.bas to use the appropriate variable name:
   ```vba
   ' In Normal.dotm copy:
   envPath = Environ("VBA_MACRO_SYNC_PATH_NORMAL")

   ' In OtherTemplate.dotm copy:
   envPath = Environ("VBA_MACRO_SYNC_PATH_OTHER")
   ```
4. **Each template syncs independently** - modules export to their own subdirectories
5. **Note:** If you update VBAMacroSync.bas, you'll need to manually copy the updated version to each template.

This approach keeps templates isolated and is ideal when you have a small number of template files (2-5) to manage.
