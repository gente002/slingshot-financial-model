# Mac Compatibility Guide

## Status: Partial Support

The RDK kernel and domain modules are VBA — they work on both Windows and Mac Excel. However, several utility features use Windows-only COM objects. The core workflow (bootstrap, run model, formula tabs, input validation) works on Mac with manual setup.

## What Works on Mac

| Feature | Status | Notes |
|---------|--------|-------|
| BootstrapWorkbook (VBA) | Works | All tab creation, config loading, formula generation |
| Run Model | Works | Domain computation, derived fields, Detail/CSV output |
| Formula tab creation | Works | All placeholder resolution, quarterly layouts |
| Input validation | Works | Data Validation rules applied via VBA |
| Health formatting | Works | Conditional formatting via VBA |
| Named ranges | Works | Created via VBA |
| ProveIt audit | Works | Formula-based checks |
| Dev mode toggle | Works | Tab visibility switching |
| Cover Page / User Guide | Works | Config-driven content |
| Pipeline config | Works | Step enable/disable |

## What Doesn't Work on Mac

| Feature | Reason | Workaround |
|---------|--------|------------|
| Setup.bat / Bootstrap.ps1 | Windows batch/PowerShell | Manual: copy config_insurance/ to config/, open workbook, run BootstrapWorkbook from VBA editor |
| SHA256 hashing (snapshots) | Uses WScript.Shell + PowerShell | Snapshots save without hash verification. Integrity checking skipped. |
| SaveConfigToSnapshot | Uses Scripting.FileSystemObject | Manual: copy config/ directory into snapshot folder |
| RestoreConfigFromSnapshot | Uses Scripting.FileSystemObject | Manual: copy config/ from snapshot back to config/ |
| BuildConfigHash | Uses Scripting.FileSystemObject | Config drift detection unavailable. Model still runs. |
| KernelWorkspace (if using FSO) | Uses Scripting.FileSystemObject | Workspace save/load requires manual file management |
| Workbook_Open injection | Uses VBProject.VBComponents | Disabled by default (BUG-081). No impact. |

## Mac Setup Instructions

1. Copy `config_insurance/` (or your model's config directory) to `config/` at the project root
2. Open the .xlsm workbook in Excel for Mac
3. Enable macros when prompted
4. Open VBA editor (Opt+F11) and run `KernelBootstrap.BootstrapWorkbook`
5. Close VBA editor. The model is ready.

To run the model: click "Run Model" on the Dashboard tab (or run `KernelEngine.RunModel` from VBA editor).

## Developer Notes

If you're building a production version targeting Mac:
- Replace `Scripting.FileSystemObject` with VBA's native `Dir()`, `MkDir`, `FileCopy`, `Kill` functions
- Replace `WScript.Shell` PowerShell calls with a VBA-native hash function or skip hashing
- Replace `Environ("TEMP")` with `Application.DefaultFilePath` or Mac-compatible temp path
- The `Application.Run` dispatch pattern works identically on Mac
