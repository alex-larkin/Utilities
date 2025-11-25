This file used to be named Word_VBA_Template_Management_SRD.md
It is now named TemplateViewer_SRD.md

# System Requirements Document: Word VBA Template Management System

## 1. Executive Summary

This system provides an automated solution for organizing and accessing multiple Word VBA macro templates. It addresses the limitation of Word's VBA Editor, which only displays Normal.dotm's modules by default, by automatically loading multiple .dotm templates as hidden documents at startup and making their VBA projects visible in the Project Explorer for easy editing.

The system consists of template management macros stored in **Utilities.dotm**, with **AutoExec()** in **Normal.dotm** calling these macros at Word startup to initialize the system.

## 2. Use Case

### 2.1 Current Situation
The user has hundreds of VBA macros organized across multiple .dotm template files stored in `C:\Users\larka\AppData\Roaming\Microsoft\Templates` (alongside Normal.dotm). These macros:
- Call each other frequently across templates
- Are available globally to all Word documents
- Require frequent editing and maintenance

### 2.2 Problem Statement
Word's VBA Editor only displays VBA projects for:
- Normal.dotm (always visible)
- The active document
- Templates that are explicitly opened for editing

This creates a workflow problem: the user must manually open each template file via Windows Explorer every time they need to edit macros in that template, which is inefficient when managing hundreds of macros across multiple templates.

### 2.3 Solution Overview
The system automatically:
1. Opens all .dotm template files at Word startup
2. Hides these template windows (making them invisible to the user)
3. Keeps templates loaded throughout the Word session
4. Makes all template VBA projects visible and editable in the VBA Editor
5. Closes hidden templates cleanly when Word quits

### 2.4 User Workflow

**Startup:**
1. User launches Word
2. System automatically opens and hides all templates
3. User sees normal Word interface (no visible clutter)

**During Work:**
1. User presses Alt+F11 to open VBA Editor
2. All template projects are visible in Project Explorer
3. User can navigate and edit any macro in any template
4. User works with regular documents normally
5. All macros from all templates are available globally

**Shutdown:**
1. User closes Word (File → Exit or closes application)
2. System automatically closes all hidden templates
3. Prompts to save if any templates have unsaved changes

## 3. Functional Requirements

### 3.1 Template Loading (Startup)

**FR-1.1: Automatic Template Discovery**
- System shall scan the directory `%APPDATA%\Microsoft\Templates`
- System shall identify all files with .dotm extension
- System shall exclude Normal.dotm from processing (Normal.dotm is always loaded)
- System shall include Utilities.dotm in processing (so its VBA project is visible for editing)

**FR-1.2: Template Opening**
- System shall open each discovered .dotm file using `Documents.Open`
- System shall open templates for editing (not as document-creation templates)
- Opening shall occur during Word startup via AutoExec() in Normal.dotm calling `Utilities.LoadTemplates()`

**FR-1.3: Window Hiding**
- System shall set `.Visible = False` on each opened template's window
- Hidden templates shall not appear in taskbar
- Hidden templates shall not appear in Alt+Tab window switcher
- Hidden templates shall not be visible anywhere in the UI

**FR-1.4: VBA Project Visibility**
- Once templates are opened, their VBA projects shall be visible in VBA Editor's Project Explorer
- Projects shall be editable (modules, procedures, forms accessible)
- This visibility shall persist throughout the Word session

### 3.2 Template Management (During Session)

**FR-2.1: Template Persistence**
- Hidden templates shall remain open for the entire Word session
- Templates shall remain loaded even when no visible documents are open
- Templates shall remain loaded when user opens/closes regular documents

**FR-2.2: Global Macro Availability**
- All macros in all templates shall be available for execution
- Macros can call procedures across templates
- Macros are available to all Word documents

**FR-2.3: VBA Editor Integration**
- All template projects visible in Project Explorer at all times
- User can expand/collapse project trees
- User can edit code in any template
- User can save individual templates (Ctrl+S while in that project)

### 3.3 Template Closing (Shutdown)

**FR-3.1: Quit Event Detection**
- System shall detect when Word application is quitting
- Detection shall occur via Application.Quit event handler

**FR-3.2: Graceful Closing**
- System shall close all hidden templates when Word quits
- If templates have unsaved changes, system shall prompt user to save
- Prompts shall specify which template has changes
- System shall respect user's save/don't save choices

**FR-3.3: Cleanup**
- All template document objects shall be properly closed
- No orphaned Word processes shall remain after quit
- Resources shall be released properly

### 3.4 User Control

**FR-4.1: Manual Close Function**
- User shall be able to manually close hidden templates via macro
- Manual close macro shall be accessible (e.g., via ribbon or keyboard shortcut)
- Manual close shall follow same save-prompt behavior as automatic close

**FR-4.2: Error Handling**
- If template file is missing, system shall log error and continue with remaining templates
- If template is locked/in use, system shall notify user and skip that file
- Errors shall not prevent Word from starting

## 4. Technical Requirements

### 4.1 Environment
- **Platform:** Microsoft Word for Windows (2013 or later)
- **VBA Version:** VBA 7.0+ (64-bit compatible)
- **File System:** NTFS (Windows)
- **User Permissions:** Read/write access to Templates folder

### 4.2 Code Organization

**TR-2.1: Module Structure**

The system is organized across two templates:

**Utilities.dotm** (contains all template management code):
- Standard Module: `TemplateViewer.bas` - Contains `LoadTemplates()` and `CloseTemplates()` procedures
- Class Module: `TemplateViewerAppEvents.cls` - For Application event handling
- This is where all the template management logic resides

**Normal.dotm** (minimal integration point):
- Contains AutoExec() macro that calls `Utilities.LoadTemplates()` at startup
- No other template management code in Normal.dotm
- Keeps Normal.dotm clean and focused

**TR-2.2: AutoExec Integration**
- AutoExec() in Normal.dotm shall invoke `Utilities.LoadTemplates()` at startup
- The call shall be: `Application.Run "Utilities.LoadTemplates"` (or equivalent)
- Shall not replace or interfere with other AutoExec functionality
- User's existing AutoExec code remains intact

**TR-2.3: Variable Scoping**
- Module-level variable to store Application event handler reference
- Must persist throughout Word session (prevent garbage collection)
- Global collection or array to track opened template document objects

### 4.3 File Handling

**TR-3.1: Path Resolution**
- Use `Application.Options.DefaultFilePath(wdUserTemplatesPath)` to get Templates folder
- Fallback to `Environ("APPDATA") & "\Microsoft\Templates"` if needed
- Handle paths with spaces and special characters

**TR-3.2: File Enumeration**
- Use `Dir()` function to enumerate .dotm files
- Filter for .dotm extension only
- Skip Normal.dotm explicitly

**TR-3.3: File Access**
- Use `Documents.Open` with ReadOnly:=False for editing capability
- Handle file-not-found errors gracefully
- Handle file-locked errors (file in use by another process)

### 4.4 Performance

**TR-4.1: Startup Time**
- Template loading shall complete within 5 seconds for up to 10 templates
- Shall not significantly delay Word startup
- User can begin working before all templates fully load (if needed)

**TR-4.2: Memory Usage**
- Each hidden template consumes ~2-5 MB RAM
- System shall function with up to 20 templates
- Total memory overhead: <100 MB

**TR-4.3: VBA Editor Performance**
- Project Explorer shall remain responsive with all templates loaded
- Expanding/collapsing projects shall be instantaneous
- No lag when switching between projects

## 5. Non-Functional Requirements

### 5.1 Reliability
- System shall function across Word restarts
- System shall handle template file corruption gracefully
- System shall not cause Word crashes or hangs

### 5.2 Maintainability
- Code shall be well-commented
- Variable and procedure names shall be descriptive
- Error messages shall be clear and actionable

### 5.3 Usability
- System shall be transparent to user (no visible UI changes)
- VBA Editor experience shall feel native
- No learning curve for basic usage

### 5.4 Compatibility
- Shall work with Word 2013, 2016, 2019, 2021, Microsoft 365
- Shall work on both 32-bit and 64-bit Office installations
- Shall not conflict with other add-ins or templates

## 6. Constraints and Assumptions

### 6.1 Constraints
- Cannot modify Windows API behavior (virtual desktop manipulation ruled out)
- Must use native VBA capabilities only
- Cannot use external DLLs or dependencies
- Limited by Word VBA security settings (macros must be enabled)

### 6.2 Assumptions
- User has macro security set to allow macros in Normal.dotm
- Templates folder location follows standard Windows conventions
- User has adequate RAM (4+ GB) for multiple open templates
- Template files are not password-protected
- Templates use .dotm extension (not legacy .dot)

### 6.3 Legacy Considerations
- User has existing .dot files (pre-2007 format)
- These should be converted to .dotm for macro support
- System shall handle both formats if .dot files contain macros

## 7. Success Criteria

The system shall be considered successful when:
1. User can start Word and all templates load automatically without manual intervention
2. User opens VBA Editor and sees all template projects in Project Explorer
3. User can edit macros in any template without opening template files manually
4. Hidden templates cause no visible UI clutter (no taskbar items, no visible windows)
5. Templates close cleanly when Word quits, with proper save prompts
6. System functions reliably across multiple Word sessions
7. No performance degradation in normal Word operations
8. User reports improved workflow efficiency for macro management

## 8. Future Enhancements (Out of Scope)

The following features are not part of initial implementation but could be considered later:
- Configuration UI to select which templates to auto-load
- Templates organized in subfolders
- Lazy loading (load templates on-demand when VBA Editor opens)
- Template load order specification
- Load/unload individual templates without restarting Word
- Status indicator showing which templates are currently loaded
- Integration with version control systems

## 9. Risks and Mitigation

| Risk | Impact | Likelihood | Mitigation |
|------|--------|------------|------------|
| Template file corruption | High | Low | Graceful error handling, skip corrupted files |
| Memory overhead with many templates | Medium | Medium | Document memory requirements, test with 20+ templates |
| Conflict with existing AutoExec code | High | Low | Design as modular addition to existing AutoExec |
| User accidentally closes hidden template | Low | Low | Not possible - templates are hidden from UI |
| Word version compatibility issues | Medium | Low | Test on multiple Word versions |
| Slow startup with many templates | Medium | Medium | Optimize file enumeration, consider lazy loading |

## 10. Acceptance Testing

The system shall pass the following tests:

**Test 1: Basic Functionality**
- Start Word → All .dotm files auto-load and hide
- Open VBA Editor → All projects visible
- Close Word → Templates close with save prompts

**Test 2: Cross-Template Macro Calls**
- Create macro in Template A that calls macro in Template B
- Execute macro successfully
- Verify cross-template functionality works

**Test 3: Editing and Saving**
- Open VBA Editor
- Modify code in hidden template
- Save template (Ctrl+S)
- Close and restart Word
- Verify changes persisted

**Test 4: Error Handling**
- Remove one template file
- Start Word
- Verify system handles missing file gracefully
- Verify other templates still load

**Test 5: Clean Shutdown**
- Make changes to hidden template without saving
- Close Word
- Verify save prompt appears
- Test both Save and Don't Save options

## 11. Implementation Notes

**Development Priority:**
1. LoadTemplates() macro (core functionality)
2. Application.Quit event handler and CloseTemplates()
3. Error handling and logging
4. Testing and refinement

**Code Location:**
- Template management code (`LoadTemplates()`, `CloseTemplates()`, `TemplateViewerAppEvents`) resides in Utilities.dotm
  - `TemplateViewer.bas` - Standard module with LoadTemplates() and CloseTemplates() procedures
  - `TemplateViewerAppEvents.cls` - Class module for Application event handling
- AutoExec() in Normal.dotm calls `Utilities.LoadTemplates()` at startup
- No modifications to user's individual template files required

**Deployment:**
- Export code modules as .bas files for version control
- Document import instructions for user
- Provide backup/restore procedure for Normal.dotm

---

## Document History

| Version | Date | Author | Changes |
|---------|------|--------|---------|
| 1.0 | 2025-11-23 | Assistant | Initial requirements document |
