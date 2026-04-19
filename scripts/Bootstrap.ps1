# =============================================================================
# Bootstrap.ps1 - RDK Excel COM Automation
# Called by Setup.bat. Creates workbook, imports VBA modules, runs bootstrap.
# =============================================================================

param(
    [string]$WorkbookName = "RDK_Model"
)

$ErrorActionPreference = "Stop"

$scriptDir = $PSScriptRoot
if (-not $scriptDir) { $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
$rootDir = Split-Path -Parent $scriptDir

$wbDir     = Join-Path $rootDir "workbook"
$engineDir = Join-Path $rootDir "engine"
$configDir = Join-Path $rootDir "config"
$xlsmPath  = Join-Path $wbDir   "$WorkbookName.xlsm"

# Verify prerequisites
if (-not (Test-Path $engineDir)) {
    Write-Host "ERROR: engine/ directory not found at $engineDir" -ForegroundColor Red
    exit 1
}
if (-not (Test-Path $configDir)) {
    Write-Host "ERROR: config/ directory not found at $configDir" -ForegroundColor Red
    Write-Host "Run Setup.bat to select a config seeder first." -ForegroundColor Yellow
    exit 1
}
if (-not (Test-Path $wbDir)) {
    New-Item -ItemType Directory -Path $wbDir | Out-Null
}

# --- Win32 helper for PID lookup from COM window handle ---
Add-Type @"
    using System;
    using System.Runtime.InteropServices;
    public class RdkWin32Helper {
        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
    }
"@

# --- Check for file lock on existing workbook ---
if (Test-Path $xlsmPath) {
    try {
        $lockTest = [System.IO.File]::Open($xlsmPath, 'Open', 'ReadWrite', 'None')
        $lockTest.Close()
        $lockTest.Dispose()
    } catch {
        Write-Host "ERROR: $xlsmPath is locked by another process." -ForegroundColor Red
        Write-Host "Please close $WorkbookName.xlsm in Excel and try again." -ForegroundColor Yellow
        Write-Host "(Other Excel workbooks can stay open.)" -ForegroundColor Yellow
        exit 1
    }
}

# --- Kill zombie Excel processes from prior failed runs (BUG-004, AP-54) ---
# Only kill truly headless Excel processes (no window title AND no visible window).
# Visible Excel processes with user windows are never touched.
$zombies = Get-Process Excel -ErrorAction SilentlyContinue |
           Where-Object { $_.MainWindowTitle -eq "" -and $_.MainWindowHandle -eq 0 }
if ($zombies) {
    Write-Host "Found $($zombies.Count) zombie Excel process(es) from prior run. Cleaning up..." -ForegroundColor Yellow
    $zombies | Stop-Process -Force
    Start-Sleep -Seconds 2
}

# --- COM add-in protection (BUG-016 follow-up) ---
# The /automation flag suppresses add-in loading. When Excel quits, it creates
# incomplete HKCU stub entries (LoadBehavior only, no Manifest/FriendlyName)
# that override the complete HKLM registrations, permanently breaking add-ins.
# We record which HKCU keys exist before launch so we only clean up new stubs
# created by this session (not pre-existing user overrides).
$comAddinPaths = @(
    "HKCU:\Software\Microsoft\Office\Excel\Addins\SNL.Clients.Office.Excel.ExcelAddIn",
    "HKCU:\Software\Microsoft\Office\Excel\Addins\SPGMI.ExcelShell"
)
$preExistingHkcuKeys = @{}
foreach ($cap in $comAddinPaths) {
    $preExistingHkcuKeys[$cap] = (Test-Path $cap)
}

# --- XLSTART / Personal.xlsb protection ---
# The /automation flag suppresses XLSTART loading (including Personal.xlsb).
# When Excel quits after /automation, it may add entries to
# Resiliency\DisabledItems that prevent XLSTART items from loading on next
# normal launch. Snapshot the key now; remove new entries after quit.
$disabledItemsPath = $null
$preDisabledNames = @()
foreach ($officeVer in @("16.0", "15.0")) {
    $diPath = "HKCU:\Software\Microsoft\Office\$officeVer\Excel\Resiliency\DisabledItems"
    $resParent = "HKCU:\Software\Microsoft\Office\$officeVer\Excel\Resiliency"
    if (Test-Path $diPath) {
        $disabledItemsPath = $diPath
        $props = (Get-Item $diPath -ErrorAction SilentlyContinue).Property
        if ($props) { $preDisabledNames = @($props) }
        break
    } elseif (Test-Path $resParent) {
        # DisabledItems subkey doesn't exist yet but may be created by /automation quit
        $disabledItemsPath = $diPath
        break
    }
}

# --- Excel COM automation ---
Write-Host "Starting Excel COM automation..." -ForegroundColor Cyan

$excel = $null
$workbook = $null
$automationPid = $null

try {
    # Create a new Excel COM instance directly. New-Object -ComObject creates an
    # independent Excel process and returns a direct COM reference — no Running
    # Object Table lookup needed. This coexists with any existing Excel windows,
    # so the user does not need to close other workbooks first (BUG-050 fix).
    Write-Host "Creating Excel COM instance..." -ForegroundColor Cyan
    try {
        $excel = New-Object -ComObject Excel.Application
    } catch {
        # COM creation can fail if a problematic add-in blocks startup
        # (e.g., S&P Capital IQ — COM error 80080005). Fall back to
        # /automation mode which suppresses add-in loading but requires
        # no other Excel instances to be running.
        Write-Host "Direct COM creation failed: $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Host "Retrying with /automation mode (add-ins suppressed)..." -ForegroundColor Yellow

        $preExistingExcel = Get-Process Excel -ErrorAction SilentlyContinue
        if ($preExistingExcel) {
            Write-Host "" -ForegroundColor Red
            Write-Host "ERROR: A COM add-in is blocking Excel startup, and /automation" -ForegroundColor Red
            Write-Host "       fallback requires all Excel windows to be closed." -ForegroundColor Red
            Write-Host "Please close all Excel windows and try again." -ForegroundColor Yellow
            Write-Host "If Excel is hidden, run:  taskkill /F /IM EXCEL.EXE" -ForegroundColor Yellow
            exit 1
        }

        $excelPath = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe" -ErrorAction SilentlyContinue).'(Default)'
        if (-not $excelPath) {
            $excelPath = (Get-ItemProperty "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\excel.exe" -ErrorAction SilentlyContinue).'(Default)'
        }
        if (-not ($excelPath -and (Test-Path $excelPath))) {
            throw "Cannot find Excel executable and direct COM creation failed."
        }

        $automationProc = Start-Process $excelPath -ArgumentList "/automation" -PassThru
        $automationPid = $automationProc.Id
        Write-Host "  /automation PID: $automationPid" -ForegroundColor Gray
        $excel = $null
        for ($attempt = 1; $attempt -le 10; $attempt++) {
            Start-Sleep -Seconds 2
            try {
                $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
                break
            } catch {
                Write-Host "  Waiting for Excel (attempt $attempt/10)..." -ForegroundColor Yellow
            }
        }
        if (-not $excel) {
            throw "Could not create Excel instance via either COM or /automation."
        }
    }

    # Track our Excel process PID for targeted cleanup (only kill ours, not user's)
    if (-not $automationPid) {
        $hwnd = [IntPtr]$excel.Hwnd
        $ourPid = [uint32]0
        [RdkWin32Helper]::GetWindowThreadProcessId($hwnd, [ref]$ourPid)
        $automationPid = [int]$ourPid
    }
    Write-Host "  Excel PID: $automationPid" -ForegroundColor Gray
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1  # msoAutomationSecurityLow

    Write-Host "Excel version: $($excel.Version)" -ForegroundColor Cyan

    # Step 1: Create seed workbook in memory and save as .xlsm
    Write-Host "Creating workbook..." -ForegroundColor Cyan
    $workbook = $excel.Workbooks.Add()

    # Strip down to a single sheet
    while ($workbook.Sheets.Count -gt 1) {
        $workbook.Sheets.Item($workbook.Sheets.Count).Delete()
    }
    $workbook.Sheets.Item(1).Name = "Sheet1"

    # Save directly as macro-enabled .xlsm (format 52)
    if (Test-Path $xlsmPath) { Remove-Item $xlsmPath -Force }
    $workbook.SaveAs($xlsmPath, 52)
    Write-Host "Created: $xlsmPath" -ForegroundColor Green

    # Step 2: Import VBA modules from engine/
    Write-Host ""
    Write-Host "Importing VBA modules..." -ForegroundColor Cyan

    # --- Conditional module import (BUG-082 fix) ---
    # Read DomainModule from granularity_config.csv to determine which domain
    # modules are needed. Importing unused domain modules inflates the VBA project
    # size (830KB with all 30 modules), which triggers dual-VBA-project crashes
    # with PERSONAL.XLSB on certain Office builds. Conditional import drops
    # unnecessary modules (~100KB / 3 modules for sample config).
    $domainModule = ""
    $granCfgPath = Join-Path $configDir "granularity_config.csv"
    if (Test-Path $granCfgPath) {
        foreach ($line in (Get-Content $granCfgPath)) {
            if ($line -match '"DomainModule"\s*,\s*"([^"]+)"') {
                $domainModule = $Matches[1]
                break
            }
        }
    }
    if ($domainModule) {
        Write-Host "  DomainModule: $domainModule (conditional import)" -ForegroundColor Cyan
    } else {
        Write-Host "  DomainModule: not found -- importing all modules (safe fallback)" -ForegroundColor Yellow
    }

    # Build skip list based on active domain engine
    $skipModules = @()
    switch ($domainModule) {
        "SampleDomainEngine" {
            $skipModules = @("InsuranceDomainEngine", "Ins_GranularCSV", "Ins_QuarterlyAgg", "Ins_Triangles", "Ins_Tests")
        }
        "InsuranceDomainEngine" {
            $skipModules = @("SampleDomainEngine")
        }
        # Default (blank/unknown): import everything
    }

    $vbaProject = $workbook.VBProject
    $allBasFiles = Get-ChildItem -Path $engineDir -Filter "*.bas" |
                Where-Object { $_.Name -like "Kernel*" -or $_.Name -like "*DomainEngine*" -or $_.Name -like "Ext_*" -or $_.Name -like "Ins_*" } |
                Sort-Object Name
    $basFiles = $allBasFiles | Where-Object {
        $modName = [System.IO.Path]::GetFileNameWithoutExtension($_.Name)
        $modName -notin $skipModules
    }

    $skippedCount = $allBasFiles.Count - $basFiles.Count
    if ($skippedCount -gt 0) {
        Write-Host "  Skipping $skippedCount module(s) not needed for $domainModule`: $($skipModules -join ', ')" -ForegroundColor Yellow
    }

    if ($basFiles.Count -eq 0) {
        Write-Host "WARNING: No .bas files found in $engineDir" -ForegroundColor Yellow
    }

    foreach ($basFile in $basFiles) {
        $moduleName = [System.IO.Path]::GetFileNameWithoutExtension($basFile.Name)
        try {
            $existingModule = $vbaProject.VBComponents.Item($moduleName)
            if ($existingModule) {
                $vbaProject.VBComponents.Remove($existingModule)
            }
        } catch {}

        $vbaProject.VBComponents.Import($basFile.FullName) | Out-Null
        Write-Host "  $($basFile.Name) OK" -ForegroundColor Green
    }

    Write-Host "Imported $($basFiles.Count) modules." -ForegroundColor Green

    # Verify VBA modules loaded
    Write-Host ""
    Write-Host "VBA project modules:" -ForegroundColor Cyan
    $moduleCount = 0
    foreach ($comp in $vbaProject.VBComponents) {
        if ($comp.Type -eq 1) {  # vbext_ct_StdModule
            Write-Host "  [Module] $($comp.Name) ($($comp.CodeModule.CountOfLines) lines)" -ForegroundColor Gray
            $moduleCount++
        }
    }
    Write-Host "Total standard modules: $moduleCount" -ForegroundColor Cyan
    if ($moduleCount -lt $basFiles.Count) {
        Write-Host "WARNING: Expected $($basFiles.Count) modules but found $moduleCount!" -ForegroundColor Red
    }

    # Step 3: Normalize config CSV line endings (VBA needs CRLF)
    Write-Host "Normalizing CSV line endings..." -ForegroundColor Cyan
    Get-ChildItem -Path $configDir -Filter "*.csv" | ForEach-Object {
        $content = [System.IO.File]::ReadAllText($_.FullName)
        $content = $content -replace "`r`n", "`n"
        $content = $content -replace "`n", "`r`n"
        [System.IO.File]::WriteAllText($_.FullName, $content)
    }

    # Step 4: Run BootstrapWorkbook macro
    Write-Host ""
    Write-Host "Running BootstrapWorkbook macro..." -ForegroundColor Cyan
    $workbook.Save()

    try {
        $excel.Run("BootstrapWorkbook")
        Write-Host "Bootstrap completed successfully!" -ForegroundColor Green
    } catch {
        $diagMsg = ""
        try { $diagMsg = $workbook.Sheets.Item(1).Cells.Item(1, 1).Text } catch {}
        if ($diagMsg) { Write-Host "DIAG: $diagMsg" -ForegroundColor Yellow }

        Write-Host "Sheets in workbook:" -ForegroundColor Yellow
        for ($si = 1; $si -le $workbook.Sheets.Count; $si++) {
            Write-Host "  $($workbook.Sheets.Item($si).Name)" -ForegroundColor Yellow
        }
        throw
    }

    # Step 5: Save and close
    $excel.Visible = $false
    $workbook.Save()
    Write-Host ""
    Write-Host "Workbook saved: $xlsmPath" -ForegroundColor Green

} catch {
    Write-Host ""
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red

    if ($_.Exception.Message -like "*Excel.Application*" -or
        $_.Exception.Message -like "*Cannot create*") {
        Write-Host "Microsoft Excel 2016+ is required." -ForegroundColor Red
    }

    if ($_.Exception.Message -like "*programmatic access*" -or
        $_.Exception.Message -like "*VBProject*") {
        Write-Host ""
        Write-Host "VBA project access is blocked. To fix:" -ForegroundColor Yellow
        Write-Host "  1. Open Excel" -ForegroundColor Yellow
        Write-Host "  2. File > Options > Trust Center > Trust Center Settings" -ForegroundColor Yellow
        Write-Host "  3. Macro Settings > Check 'Trust access to the VBA project object model'" -ForegroundColor Yellow
    }

    exit 1

} finally {
    if ($workbook) {
        try { $workbook.Close($false) } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null } catch {}
    }
    if ($excel) {
        try {
            $excel.DisplayAlerts = $false
            $excel.Visible = $false
        } catch {}
        # BUG-081: Graceful quit instead of force-kill. Force-kill leaves
        # Resiliency\StartupItems sentinels that trigger the safe mode prompt.
        # The quit-cycle may write DisabledItems (XLSTART suppression) but
        # the cleanup below removes those.
        try { $excel.Quit() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    # --- Wait for graceful exit, force-kill only as last resort ---
    if ($automationPid) {
        $ourExcel = Get-Process -Id $automationPid -ErrorAction SilentlyContinue
        if ($ourExcel) {
            try { $ourExcel.WaitForExit(10000) } catch {}
            $ourExcel = Get-Process -Id $automationPid -ErrorAction SilentlyContinue
            if ($ourExcel) {
                # Graceful quit failed -- force kill as last resort
                Stop-Process -Id $automationPid -Force -ErrorAction SilentlyContinue
                try { $ourExcel.WaitForExit(5000) } catch {}
                Write-Host "Excel did not quit gracefully -- force killed PID $automationPid." -ForegroundColor Yellow
            } else {
                Write-Host "Excel exited gracefully (PID $automationPid)." -ForegroundColor Green
            }
        }
    }

    # Brief pause for OS to flush pending registry writes after process exit
    Start-Sleep -Seconds 1

    # --- Remove HKCU COM add-in stubs created by /automation (BUG-016 follow-up) ---
    # Excel's /automation + Quit cycle creates incomplete HKCU stubs that override
    # the full HKLM registrations. Remove any HKCU keys that did not exist before
    # this session so the complete HKLM entries take effect on next Excel launch.
    foreach ($cap in $comAddinPaths) {
        try {
            if ((Test-Path $cap) -and -not $preExistingHkcuKeys[$cap]) {
                Remove-Item $cap -Force -ErrorAction Stop
                Write-Host "Removed HKCU stub: $(Split-Path $cap -Leaf)" -ForegroundColor Yellow
            }
        } catch {}
    }

    # --- Remove new Resiliency\DisabledItems entries (BUG-045 comprehensive scan) ---
    # Scan ALL Office versions for DisabledItems, not just pre-snapshotted paths.
    # The pre-check may have found no Resiliency key, but /automation could create one.
    foreach ($officeVer in @("16.0", "15.0")) {
        $diScanPath = "HKCU:\Software\Microsoft\Office\$officeVer\Excel\Resiliency\DisabledItems"
        if (Test-Path $diScanPath) {
            $currentProps = (Get-Item $diScanPath -ErrorAction SilentlyContinue).Property
            if ($currentProps) {
                foreach ($prop in @($currentProps)) {
                    if ($prop -notin $preDisabledNames) {
                        try {
                            Remove-ItemProperty -Path $diScanPath -Name $prop -Force -ErrorAction Stop
                            Write-Host "Restored XLSTART item (removed $officeVer DisabledItems/$prop)" -ForegroundColor Yellow
                        } catch {}
                    }
                }
            }
        }
    }

    # --- Clean up orphaned DocumentRecovery entries (BUG-047) ---
    # Process kill leaves DocumentRecovery entries that cause Book counter to
    # increment across sessions (Book2, Book3... instead of Book1).
    foreach ($officeVer in @("16.0", "15.0")) {
        $drPath = "HKCU:\Software\Microsoft\Office\$officeVer\Excel\Resiliency\DocumentRecovery"
        if (Test-Path $drPath) {
            $drEntries = Get-ChildItem $drPath -ErrorAction SilentlyContinue
            if ($drEntries) {
                foreach ($drEntry in $drEntries) {
                    Remove-Item $drEntry.PSPath -Force -ErrorAction SilentlyContinue
                }
                Write-Host "Cleared $($drEntries.Count) DocumentRecovery entries ($officeVer)." -ForegroundColor Yellow
            }
        }
    }

    # --- Clear crash-detection StartupItems (BUG-071) ---
    # Excel writes a sentinel to Resiliency\StartupItems on launch and removes
    # it on graceful exit. Process kill leaves the sentinel, causing Excel to
    # offer Safe Mode on the next normal launch. Remove all StartupItems
    # entries created by this session.
    foreach ($officeVer in @("16.0", "15.0")) {
        $siPath = "HKCU:\Software\Microsoft\Office\$officeVer\Excel\Resiliency\StartupItems"
        if (Test-Path $siPath) {
            $siProps = (Get-Item $siPath -ErrorAction SilentlyContinue).Property
            if ($siProps) {
                foreach ($prop in @($siProps)) {
                    try {
                        Remove-ItemProperty -Path $siPath -Name $prop -Force -ErrorAction Stop
                    } catch {}
                }
                Write-Host "Cleared $($siProps.Count) StartupItems entry(ies) -- Safe Mode prompt prevented." -ForegroundColor Yellow
            }
        }
    }

    # --- Remove stale AutoRecover files for RDK_Model (BUG-081) ---
    # Force-kill (or unclean quit) leaves recovery files on disk.
    # Corrupt recovery files crash Excel when clicked in Document Recovery pane,
    # creating a crash loop (crash -> StartupItems -> safe mode -> repeat).
    $excelAppData = [Environment]::GetFolderPath('ApplicationData') + "\Microsoft\Excel"
    if (Test-Path $excelAppData) {
        $rdkRecovery = Get-ChildItem $excelAppData -Directory -Filter "RDK_Model*" -ErrorAction SilentlyContinue
        foreach ($rf in $rdkRecovery) {
            try {
                Remove-Item $rf.FullName -Recurse -Force -ErrorAction Stop
                Write-Host "Removed recovery folder: $($rf.Name)" -ForegroundColor Yellow
            } catch {}
        }
    }

    # --- Remove stale XLSTART lock files (BUG-047) ---
    # Process kill can leave ~$PERSONAL.XLSB owner files that block loading.
    $xlstartPath = [Environment]::GetFolderPath('ApplicationData') + "\Microsoft\Excel\XLSTART"
    if (Test-Path $xlstartPath) {
        $lockFiles = Get-ChildItem $xlstartPath -Force -Filter '~$*' -ErrorAction SilentlyContinue
        foreach ($lf in $lockFiles) {
            try {
                Remove-Item $lf.FullName -Force -ErrorAction Stop
                Write-Host "Removed stale XLSTART lock file: $($lf.Name)" -ForegroundColor Yellow
            } catch {}
        }
    }
}
