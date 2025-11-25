# Software Requirements Document
## VBA Macro Sync System

**Version:** 1.0
**Date:** 2025-11-19
**Status:** Implemented

---

## 1. Introduction

### 1.1 Purpose
This document specifies the requirements for the VBA Macro Sync System, a two-way synchronization solution that enables collaborative development of Microsoft Word VBA macros through Git version control.

### 1.2 Scope
The VBA Macro Sync System synchronizes VBA code modules between Microsoft Word templates and designated filesystem folders specified by environment variables. Each template that implements the system (such as Normal.dotm, Utilities.dotm, or custom templates) synchronizes its VBA code independently to its own designated folder. The system is designed to integrate with Git-based workflows, enabling multiple developers to collaborate on Word macros using standard version control practices.

### 1.3 Intended Audience
VBA developers working collaboratively on Word macros

---

## 2. System Overview

### 2.1 System Description
The VBA Macro Sync System operates as a VBA module (VBAMacroSync.bas) that resides in Utilities.dotm. The core functionality resides in `VMS_AutoExec()` and `VMS_AutoExit()` subroutines which are called by simple `AutoExec()` and `AutoExit()` macros in each template that wants to use the sync system. The system automatically:
- Exports all VBA components from a template to individual files in a designated folder when Word closes
- Imports VBA components from the designated folder into the template when Word opens
- When opening Word, treats the filesystem folder as the source of truth for Git-managed code
- When closing Word, treats the template file as the source of truth for Git-managed code


### 2.2 Context
The system supports workflows where:
- VBA macro files are stored in a Git repository (e.g., GitHub)
- Multiple users collaborate on macros using Git for version control
- Users edit macros either in Word's VBA Editor or in external text editors (e.g., VS Code)
- Git handles all merge conflict resolution

---

## 3. Functional Requirements

### 3.1 Export Functionality

**FR-1.1** The system shall export all standard modules (.bas), class modules (.cls), and UserForms (.frm) from the calling template to its designated sync folder when Word closes.

**FR-1.2** The system shall execute the export operation automatically via the `VMS_AutoExit()` subroutine, called by the `AutoExit()` macro in each template, when Microsoft Word terminates.

**FR-1.3** The system shall create the sync folder if it does not exist at export time.

**FR-1.4** The system shall skip exporting document modules and other non-exportable VBA component types.

**FR-1.5** The system shall provide debug logging to the VBA Immediate Window detailing:
- Total components found
- Components exported successfully
- Export errors encountered
- File paths used

### 3.2 Import Functionality

**FR-2.1** The system shall import all .bas, .cls, and .frm files from the designated sync folder into the calling template when Word opens.

**FR-2.2** The system shall execute the import operation automatically via the `VMS_AutoExec()` subroutine, called by the `AutoExec()` macro in each template, when Microsoft Word starts.

**FR-2.3** When opening Word, the system shall treat the filesystem folder as the source of truth:
- If a module exists in both the template and the folder, the folder version shall replace the template version
- No conflict resolution prompts shall be presented to the user
- Git is responsible for all merge conflict resolution

**FR-2.4** When closing Word, the system shall treat the template file as the source of truth:
- If a module exists in both the template file and the folder, the template file version shall replace the folder version
- No conflict resolution prompts shall be presented to the user
- Git is responsible for all merge conflict resolution

**FR-2.5** The system shall optimize imports by comparing file contents:
- If the folder file and template module are identical, the import shall be skipped
- File comparison shall be byte-for-byte content matching

**FR-2.6** The system shall handle new modules:
- Modules present in the folder but not in the template shall be imported
- Modules present in the template but not in the folder shall be exported

**FR-2.7** The system shall provide debug logging to the VBA Immediate Window detailing:
- Files found in sync folder by type
- Comparison results
- Import operations performed
- Import errors encountered

### 3.3 Configuration

**FR-3.1** The system shall read the sync folder path from a Windows environment variable specific to each template. The environment variable name shall follow the pattern `MACROS_<TemplateName>`, where `<TemplateName>` is the base name of the template file without the .dotm extension (e.g., `MACROS_UTILITIES` for Utilities.dotm, `MACROS_NORMAL` for Normal.dotm).

**FR-3.2** The system shall display an error message if the required environment variable is not set for the calling template, and exit without performing sync operations for that template.

### 3.4 Template-Specific Integration

**FR-4.1** The VBAMacroSync.bas module shall reside in Utilities.dotm as the single authoritative implementation.

**FR-4.2** Each template that wants to use the sync system shall implement simple `AutoExec()` and `AutoExit()` macros that call `VMS_AutoExec()` and `VMS_AutoExit()` in Utilities.dotm respectively.

**FR-4.3** The `VMS_AutoExec()` and `VMS_AutoExit()` subroutines shall determine which template is calling them and use the appropriate environment variable for that template's sync folder.

**FR-4.4** The system shall provide diagnostic output about which template is being synchronized to assist with troubleshooting.

### 3.5 Manual Operations

**FR-5.1** The system shall provide a `ManualExport()` subroutine for testing export functionality without closing Word.

**FR-5.2** The system shall provide a `ManualImport()` subroutine for testing import functionality without reopening Word.

**FR-5.3** Manual operations shall display a message box confirming completion.

### 3.6 User Feedback

**FR-6.1** The system shall display "Macro sync complete" in the Word status bar for 2 seconds after import operations complete.

**FR-6.2** The system shall provide comprehensive debug logging viewable in the VBA Immediate Window (Ctrl+G).

---

## 4. Non-Functional Requirements

### 4.1 Performance

**NFR-1.1** Import operations shall skip file processing when content is identical to optimize startup time.

**NFR-1.2** The status bar message shall display for exactly 2 seconds using the Windows API Sleep function.

### 4.2 Reliability

**NFR-2.1** The system shall use error handling (`On Error Resume Next`) to prevent VBA errors from interrupting Word startup or shutdown.

**NFR-2.2** All errors shall be logged to the debug output with error number and description.

**NFR-2.3** The system shall continue processing remaining files if an individual file import or export fails.

### 4.3 Usability

**NFR-3.1** The system shall operate transparently without user interaction during normal Word startup and shutdown.

**NFR-3.2** Error messages shall provide actionable guidance (e.g., "See README.md for setup instructions").

**NFR-3.3** Debug output shall be structured and verbose to assist troubleshooting.

### 4.4 Maintainability

**NFR-4.1** The code shall include comprehensive comments explaining functionality for novice programmers.

**NFR-4.2** The code shall use descriptive variable names and structured logging.

**NFR-4.3** The code shall be organized into logical sections with header comments.

---

## 5. Technical Specifications

### 5.1 Platform Requirements
- Microsoft Word 365 (or compatible version supporting VBA)
- Windows operating system (uses Windows API Sleep function)
- VBA project object model access enabled in Trust Center settings

### 5.2 File Format Support
| Component Type | File Extension | VBA Type Constant |
|----------------|----------------|-------------------|
| Standard Module | .bas | vbext_ct_StdModule (1) |
| Class Module | .cls | vbext_ct_ClassModule (2) |
| UserForm | .frm | vbext_ct_MSForm (3) |

### 5.3 Environment Variables
The system uses environment variables following the pattern `MACROS_<TemplateName>` where `<TemplateName>` is the base name of the template file without extension (converted to uppercase):
- **MACROS_NORMAL**: Sync folder path for Normal.dotm
- **MACROS_UTILITIES**: Sync folder path for Utilities.dotm
- **MACROS_<CustomTemplate>**: Sync folder path for any custom template

Each path should be a full path to the sync folder with a trailing backslash.

### 5.4 API Dependencies
- Windows kernel32.dll Sleep function (PtrSafe declaration)
- Scripting.FileSystemObject for file comparison operations

### 5.5 VBA Object Model
- Application.VBE.VBProjects (for accessing template VBA projects)
- VBProject.VBComponents (for accessing modules within a template)
- VBComponent.Export and VBComponents.Import methods
- VBProject.Filename (for determining which template is calling the sync routines)

---

## 6. Constraints and Assumptions

### 6.1 Constraints

**C-1** The system does not synchronize deletions automatically. Users must manually:
1. Delete the module from the template
2. Delete the corresponding file from the sync folder
3. Commit the deletion to Git

**C-2** The system requires VBA project access to be enabled, which reduces security. This is a necessary trade-off for programmatic VBA manipulation.

**C-3** When opening Word, the system uses the folder as the absolute source of truth. Local changes in the template will be overwritten by folder contents if they differ.

**C-4** When closing Word, the system uses the template file as the absolute source of truth. Local changes in the folder contents will be overwritten by the template file if they differ.

**C-5** The system does not handle Git merge conflicts. Users must resolve conflicts using Git tools before opening Word.

**C-6** Each template requires its own environment variable to be configured separately.

### 6.2 Assumptions

**A-1** Each template's sync folder is managed by Git version control.

**A-2** Users understand Git workflows (pull, commit, push, merge conflict resolution).

**A-3** The appropriate environment variable (e.g., `MACROS_NORMAL`, `MACROS_UTILITIES`) is set before Word starts for each template that uses the sync system.

**A-4** Each template's sync folder path is accessible and has appropriate read/write permissions.

**A-5** UserForms that include .frx files (binary resources) will have those files exported automatically by the VBA Export method.

**A-6** Users will not simultaneously edit the same macro in multiple Word instances.

**A-7** Users will close Word before performing Git pull operations to avoid conflicts between in-memory VBA and incoming file changes.

**A-8** Utilities.dotm is loaded before other templates attempt to call `VMS_AutoExec()` or `VMS_AutoExit()`. This ensures the VBAMacroSync module is available when other templates need it.

---

## 7. Dependencies

### 7.1 External Dependencies
- **Git**: Version control system for managing macro files
- **Utilities.dotm**: Template containing the VBAMacroSync.bas module with the core sync functionality
- **Windows Environment Variables**: Configuration storage mechanism specifying sync folder paths for each template

### 7.2 VBA References
- Microsoft Scripting Runtime (for FileSystemObject)
- Microsoft Visual Basic for Applications Extensibility (for VBProject access)

---

## 8. Version Control Integration

### 8.1 Git Workflow
The system is designed to support the following workflow:
1. User pulls latest changes from Git repository
2. User opens Word → macros auto-import from folder (Git-managed files)
3. User edits macros in Word VBA Editor
4. User closes Word → macros auto-export to folder
5. User reviews changes in Git client (e.g., GitHub Desktop)
6. User commits and pushes changes
7. If conflicts occur, user resolves in Git before reopening Word



### 8.2 Design Philosophy
- **Source of truth**: When opening Word, Git-managed files are the source of truth and always override template file contents. When closing Word, the template files are the source of truth and always override the git-managed files. 
- **No application-level conflict resolution**: Git handles all merging and conflict detection
- **Optimistic synchronization**: Assumes users coordinate to avoid simultaneous editing

---

## 9. Implementation Notes

### 9.1 Key Functions

| Function | Purpose |
|----------|---------|
| `AutoExec()` | Template-specific macro that calls VMS_AutoExec() at Word startup (implemented in each template) |
| `AutoExit()` | Template-specific macro that calls VMS_AutoExit() at Word shutdown (implemented in each template) |
| `VMS_AutoExec()` | Core entry point for Word startup operations, determines calling template and imports its macros |
| `VMS_AutoExit()` | Core entry point for Word shutdown operations, determines calling template and exports its macros |
| `GetSyncFolderPath()` | Retrieves sync folder path from the appropriate environment variable based on calling template |
| `ExportMacrosToFolder()` | Exports all VBA components from the calling template to its sync folder |
| `ImportMacrosFromFolder()` | Imports VBA component files from sync folder into the calling template |
| `ProcessImport()` | Handles import of a single module with optimization |
| `FilesAreIdentical()` | Compares two files for byte-for-byte equality |
| `ManualExport()` | User-triggered export for testing |
| `ManualImport()` | User-triggered import for testing |

### 9.2 Template Implementation Pattern

The VBAMacroSync.bas module resides only in Utilities.dotm as the single authoritative implementation. Each template that uses VBAMacroSync must implement the following pattern:

1. **Ensure Utilities.dotm is loaded**: Utilities.dotm must be loaded in Word for other templates to access its VBAMacroSync module
2. **Implement AutoExec()**: Create a simple `AutoExec()` macro that calls `VMS_AutoExec()` in Utilities.dotm
3. **Implement AutoExit()**: Create a simple `AutoExit()` macro that calls `VMS_AutoExit()` in Utilities.dotm
4. **Set environment variable**: Define the `MACROS_<TemplateName>` environment variable pointing to the template's sync folder

Example implementation in any template (e.g., Normal.dotm or custom templates):
```vba
Sub AutoExec()
    ' Call the VMS_AutoExec function in Utilities.dotm
    Application.Run "Utilities.VMS_AutoExec"
End Sub

Sub AutoExit()
    ' Call the VMS_AutoExit function in Utilities.dotm
    Application.Run "Utilities.VMS_AutoExit"
End Sub
```

The `VMS_AutoExec()` and `VMS_AutoExit()` subroutines in Utilities.dotm automatically determine which template called them and use the corresponding environment variable.

This design provides:
- **Single source of truth**: Only one copy of VBAMacroSync.bas exists in Utilities.dotm
- **Easy maintenance**: Updates to the sync logic only need to be made in one location
- **Minimal template overhead**: Other templates only need two simple wrapper macros

### 9.3 Template Detection Mechanism

When `VMS_AutoExec()` or `VMS_AutoExit()` is called from another template via `Application.Run`, the system determines which template initiated the call by examining the call stack or using VBA's `Application.VBE.ActiveVBProject` to identify the calling template's VBProject. The template name is extracted from the VBProject filename, converted to uppercase, and used to construct the environment variable name (e.g., "Normal.dotm" → "MACROS_NORMAL").

This cross-template calling mechanism allows:
- A single VBAMacroSync module in Utilities.dotm to serve all templates
- Each template to maintain its own independent sync folder
- No hard-coded template names in the core sync logic
- Easy maintenance with updates in only one location

### 9.4 Error Handling Strategy
The system uses permissive error handling (`On Error Resume Next`) to ensure Word startup/shutdown is never blocked. All errors are logged but do not halt execution, allowing Word to function normally even if sync operations fail.

---

## 10. Future Considerations

### 10.1 Out of Scope for Current Version
- Automatic deletion synchronization
- Conflict detection and resolution UI
- Synchronization to cloud storage services
- Support for non-Windows platforms
- Encrypted or compressed module storage

### 10.2 Potential Enhancements
- Configuration file alternative to environment variables
- Synchronization status dashboard
- Selective module synchronization (inclusion/exclusion filters)
- Automatic backup before import operations
- Integration with additional version control systems

---

## 11. References

### 11.1 Related Documents
- [VBAMacroSync ReadME.md](VBAMacroSync%20ReadME.md) - User guide and setup instructions for `VBAMacroSync - Backup.bas`
- VBAMacroSync - Backup.bas - Implementation source code for previous working version that only worked with Normal.dotm

### 11.2 External Resources
- Microsoft VBA Language Reference
- Git Documentation
- Windows Environment Variables Configuration Guide
