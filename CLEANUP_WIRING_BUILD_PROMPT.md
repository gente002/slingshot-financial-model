# RDK Cleanup & Config Wiring — Claude Code Build Prompt

**Date:** April 2026
**Prior:** v1.3.0 post-TD/investor-demo cycle
**Goal:** 7 cleanup/wiring items in one session. Module split, deprecated code removal, and config wiring.

## BEFORE WRITING ANY CODE

1. Read `CLAUDE.md` — project bible.
2. Read `SESSION_NOTES.md` — **APPEND ONLY — do not truncate.**
3. Read `docs/DESIGN_Kernel_vs_Domain.md` — kernel/domain boundary rules.
4. Read `data/anti_patterns.csv` and `data/patterns.csv`.

---

## ITEM 1: Split KernelSnapshot.bas (P0)

KernelSnapshot.bas is at ~62.9KB — 1.1KB from the 64KB VBA hard limit. Split it.

**Extract into new module `KernelSnapshotIO.bas`:**
- `ExportDetailToFile` (and any Private helpers it uses exclusively)
- `ExportInputsToFile`
- `ExportSettingsToFile`
- `ExportErrorLogToFile`
- `SaveConfigToSnapshot`
- `ImportInputsFromCsv`
- `ImportDetailFromCsv` (if it exists)
- `LoadCsvToArray` (if it's a Private helper used by import/export)
- `DetectEntityCount`
- `FormatISOTimestamp`
- Any other Private helpers used exclusively by the above export/import functions

**KernelSnapshot.bas retains:**
- `SaveSnapshot` / `LoadSnapshot` / `LoadSnapshotInputsOnly` (orchestration)
- `DeleteSnapshot` / `RenameSnapshot` / `ArchiveSnapshot` / `RestoreFromArchive`
- `ListSnapshots` / `ListArchivedSnapshots`
- `WriteWAL` / `PurgeWAL`
- `ComputeSHA256` / `BuildConfigHash` / `BuildExecutionFingerprint`
- `AcquireLock` / `ReleaseLock`
- `GetProjectRoot` / `EnsureDirectoryExists`
- `GetInputsSheet` (the TAB_INPUTS/TAB_ASSUMPTIONS fallback helper)

**Update callers:** SaveSnapshot, LoadSnapshot, and KernelWorkspace call the export/import functions — update references from `KernelSnapshot.ExportDetailToFile` to `KernelSnapshotIO.ExportDetailToFile`, etc. Also check KernelTests, KernelCompare, KernelTabIO for calls.

**Public helpers that both modules need** (e.g., `GetProjectRoot`, `EnsureDirectoryExists`, `FormatISOTimestamp`): keep in KernelSnapshot since it's the established location. KernelSnapshotIO calls `KernelSnapshot.GetProjectRoot()` etc.

**Target sizes:** KernelSnapshot < 40KB, KernelSnapshotIO < 30KB.

**Update CLAUDE.md:** Kernel modules 27 → 28. Add KernelSnapshotIO to the list.

---

## ITEM 2: Remove deprecated tab references

The following tabs were removed from tab_registry: Summary, Exhibits, Charts, CumulativeView, Analysis. But kernel modules still reference them. Remove or gate all references.

**Search pattern:** `grep -rn "TAB_SUMMARY\|TAB_EXHIBITS\|TAB_CHARTS\|TAB_CUMULATIVE_VIEW\|TAB_ANALYSIS" engine/*.bas`

**For each reference found:**
- If the code block creates/populates/formats the deprecated tab → **delete the entire block** (the tab doesn't exist)
- If the code block references the tab conditionally (e.g., hide/show, protect) → **wrap in `If SheetExists(TAB_xxx) Then`** or delete
- If it's a Const declaration in KernelConstants → **keep the Const** (other modules may reference it for compatibility) but add a comment: `' DEPRECATED — tab removed from registry`

**Known locations to check:**
- KernelBootstrap.bas (~line 1163, 1369) — Summary references
- KernelOutput.bas (~line 88) — Summary reference
- KernelTabs.bas (~lines 135, 148, 447-458, 470, 711-738, 1058-1063, 1194-1209) — Charts, Exhibits, CumulativeView references
- KernelTransform.bas (~line 293-297) — Analysis reference

**Do NOT delete the tab-building subroutines themselves** (GenerateExhibits, GenerateCharts, etc.) — these are kernel capabilities that future models may use. Just gate them behind `If SheetExists()` checks so they no-op when the tab isn't in the registry.

---

## ITEM 3: Wire workspace_config.csv

KernelWorkspace.bas currently hardcodes workspace behavior. Wire it to read from workspace_config.csv via `KernelConfig.GetWorkspaceSetting()`.

**Settings to wire:**

| Setting | Where to Use | Current Hardcoded Default |
|---|---|---|
| WorkspacesEnabled | Gate `SaveWorkspace`/`LoadWorkspace` — if FALSE, show "Workspaces disabled" and exit | Always enabled |
| AutoSaveOnRun | At end of `KernelEngine.RunModel` — if TRUE, auto-save current workspace | Never auto-saves |
| MaxVersionsPerWorkspace | In `SaveWorkspace` — if version count exceeds max, warn or auto-archive oldest | No limit |
| DefaultWorkspaceName | In `ResolveWorkspaceName` — use as default when no name provided | Probably "Main" or prompts user |

**Implementation:** At the top of each relevant function, read the setting:
```vba
Dim enabled As String
enabled = KernelConfig.GetWorkspaceSetting("WorkspacesEnabled")
If StrComp(enabled, "TRUE", vbTextCompare) <> 0 Then
    MsgBox "Workspaces are disabled.", vbInformation, "RDK"
    Exit Sub
End If
```

---

## ITEM 4: Wire msgbox_config.csv (12 defined entries)

Replace hardcoded MsgBox calls with config-driven lookups for the 12 entries defined in msgbox_config.csv. The remaining ~150+ edge-case MsgBox calls stay hardcoded.

**How to use:**
```vba
Dim msg As Variant
msg = KernelConfig.GetMsgBox("RUN_COMPLETE")
If Not IsEmpty(msg) Then
    Dim txt As String: txt = CStr(msg(0))  ' Message with placeholders
    Dim ttl As String: ttl = CStr(msg(1))  ' Title
    Dim icon As String: icon = CStr(msg(2)) ' Icon
    ' Replace placeholders
    txt = Replace(txt, "{ENTITIES}", CStr(entityCount))
    txt = Replace(txt, "{PERIODS}", CStr(periodCount))
    txt = Replace(txt, "{ELAPSED}", Format(elapsed, "0.0"))
    txt = Replace(txt, "\n", vbCrLf)
    MsgBox txt, IIf(icon = "Information", vbInformation, IIf(icon = "Exclamation", vbExclamation, vbCritical)), ttl
End If
```

**Create a helper function** in KernelConfig (or KernelFormHelpers) to simplify:
```vba
Public Sub ShowConfigMsgBox(msgID As String, ParamArray replacements() As Variant)
```
This helper resolves the config, applies placeholder replacements, maps icon strings to VBA constants, and shows the MsgBox. Falls back to a generic MsgBox if the config entry is missing.

**Entries to wire (find the corresponding hardcoded MsgBox and replace):**

| MsgBoxID | Where to Find Current Hardcoded Call |
|---|---|
| BOOTSTRAP_COMPLETE | KernelBootstrap — end of bootstrap |
| RUN_COMPLETE | KernelEngine — end of RunModel |
| RUN_COMPLETE_ERRORS | KernelEngine — end of RunModel (when errors logged) |
| VALIDATION_FAILED | KernelEngine — validation failure |
| SNAPSHOT_SAVED | KernelSnapshot.SaveSnapshot |
| SNAPSHOT_LOADED | KernelSnapshot.LoadSnapshot |
| SNAPSHOT_LOADED_STALE | KernelSnapshot.LoadSnapshot (stale warning) |
| SNAPSHOT_NOT_FOUND | KernelSnapshot — not found error |
| SNAPSHOT_CORRUPTED | KernelSnapshot — corruption error |
| NO_FORMULA_CONFIG | KernelFormulaWriter — no formula config |
| NO_NAMED_RANGES | KernelFormula — no named range config |

**Important:** Also wire the workspace equivalents if KernelWorkspace has similar MsgBox calls (workspace saved, loaded, not found). Add new msgbox_config entries if needed:
```csv
"WORKSPACE_SAVED","RDK","Workspace saved: {NAME} {VERSION}","Information","OK"
"WORKSPACE_LOADED","RDK","Workspace loaded: {NAME} {VERSION}","Information","OK"
"WORKSPACE_NOT_FOUND","RDK","Workspace not found: {NAME}","Exclamation","OK"
```

---

## ITEM 5: Wire display_aliases.csv

Use display aliases in VBA-generated tab content where metric labels are shown to the user.

**Create a helper** (if `KernelConfig.GetDisplayAlias` doesn't already exist — it does):
```vba
' Already exists: KernelConfig.GetDisplayAlias(internalID) → display name or internalID if not found
```

**Wire into these modules:**

| Module | Where to Use |
|---|---|
| Ins_QuarterlyAgg.bas | When writing metric labels in QuarterlySummary row headers. Currently uses the raw column_registry metric name (e.g., "G_WP"). Replace with `KernelConfig.GetDisplayAlias("G_WP")` → "Gross Written Premium". |
| Ins_Triangles.bas | Triangle section headers and metric labels. |
| Ins_Presentation.bas | User Guide content if it references metric names. |
| KernelProveIt.bas | Prove-It check labels. |

**Do NOT change formula_tab_config.csv labels** — those are already human-readable ("Gross Written Premium") and are the authoritative display names for formula tabs. Display aliases are for VBA-generated content that currently uses internal IDs.

**Verify the alias table is complete:** Check that every metric shown to the user in QuarterlyAgg and Triangles has an entry in display_aliases.csv. Add missing entries if needed (e.g., CQ_WP, XC_EP, count metrics like G_ClsCt, G_RptCt).

---

## ITEM 6: Add regression tab capture to workspace saves (SC-09)

In `engine/KernelWorkspace.bas`, in the `ExportStateToFolder` subroutine, add this line after `KernelTabIO.ExportAllInputTabs verDir`:

```vba
KernelTabIO.ExportRegressionTabs verDir
```

This captures all formula tab outputs (per regression_config.csv) into each workspace version.

---

## ITEM 7: Add Compare Workspaces button stub

Add to `config_insurance/button_config.csv`:
```csv
"Dashboard","COMPARE_WS","Compare Workspaces","KernelFormHelpers.ShowWorkspaceCompare","FALSE","25","TRUE","9","2"
```

Add stub subroutine in `engine/KernelFormHelpers.bas`:
```vba
Public Sub ShowWorkspaceCompare()
    MsgBox "Compare Workspaces: coming soon.", vbInformation, "RDK"
End Sub
```

---

## VALIDATION GATES

1. KernelSnapshot.bas < 40KB after split
2. KernelSnapshotIO.bas exists with all export/import functions, < 30KB
3. No runtime errors when calling SaveSnapshot/LoadSnapshot (callers updated)
4. No references to TAB_SUMMARY, TAB_EXHIBITS, TAB_CHARTS, TAB_CUMULATIVE_VIEW, TAB_ANALYSIS in active code paths (only in `If SheetExists` guards or deprecated Const comments)
5. KernelWorkspace reads all 4 workspace_config settings
6. At least 11 of 12 msgbox_config entries wired to their callers via ShowConfigMsgBox
7. QuarterlyAgg uses GetDisplayAlias for metric labels
8. Ins_Triangles uses GetDisplayAlias for metric labels
9. Workspace saves include regression_tabs/ directory
10. Compare Workspaces button appears on Dashboard (user-visible)
11. All existing tests still pass
12. BS_CHECK = 0, CFS_CHECK = 0
13. SESSION_NOTES.md APPENDED only
14. CLAUDE.md kernel module count updated (27 → 28)

## LOGGING

Append to SESSION_NOTES.md. Log any bugs found. Sync config/ → config_insurance/ before ZIP delivery. Include full directory structure.
