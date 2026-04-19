# =============================================================================
# Repair-ComAddins.ps1 - Diagnose and repair S&P Capital IQ COM add-in issues
#
# This script fixes the most common failure mode: corrupted DPAPI-encrypted
# IsolatedStorage + stale HKCU registry stubs that override HKLM.
#
# Root cause (BUG-017): Killing Excel mid-run corrupts the S&P add-in's
# encrypted cache in %APPDATA%\Roaming\IsolatedStorage. The /automation flag
# or Excel crashes also cause HKCU registry stubs (LoadBehavior=2, no Manifest)
# to override the correct HKLM registrations.
#
# What this script does:
#   1. Verifies Excel is not running
#   2. Removes corrupted Roaming IsolatedStorage for the SNL publisher
#   3. Removes incomplete HKCU registry stubs so HKLM takes effect
#   4. Optionally runs MSI repair on the S&P Office installer
#
# Usage:
#   .\Repair-ComAddins.ps1              # Standard repair (storage + registry)
#   .\Repair-ComAddins.ps1 -Full        # Also run MSI repair
#   .\Repair-ComAddins.ps1 -Diagnose    # Check status without changing anything
#
# After running: restart Excel and re-login to S&P Capital IQ.
#
# This script is SAFE to run from any directory -- it does not depend on the
# RDK project structure. Copy it anywhere you need it.
# =============================================================================

param(
    [switch]$Full,
    [switch]$Diagnose
)

$ErrorActionPreference = "Stop"

# --- S&P add-in identifiers ---
$snlProgId  = "SNL.Clients.Office.Excel.ExcelAddIn"
$spgmiProgId = "SPGMI.ExcelShell"
$hkcuBase   = "HKCU:\Software\Microsoft\Office\Excel\Addins"
$hklmBase   = "HKLM:\Software\Microsoft\Office\Excel\Addins"
$isoPublisher = "Publisher.rzts3tkwo03sjqp1f4dfa2tfj1uydzrp"
$isoDir     = "$env:APPDATA\IsolatedStorage\$isoPublisher"
$snlMsiId   = "{075F6A24-7407-49A0-A7FA-E89DC1FA61B3}"

# --- Functions ---
function Show-AddinStatus {
    param([string]$ProgId)
    $hkcu = "$hkcuBase\$ProgId"
    $hklm = "$hklmBase\$ProgId"
    $hkcuExists = Test-Path $hkcu
    $hklmExists = Test-Path $hklm

    Write-Host "  $ProgId" -ForegroundColor Cyan
    if ($hklmExists) {
        $props = Get-ItemProperty $hklm
        Write-Host "    HKLM: LoadBehavior=$($props.LoadBehavior), FriendlyName=$($props.FriendlyName)" -ForegroundColor Green
    } else {
        Write-Host "    HKLM: NOT REGISTERED" -ForegroundColor Red
    }
    if ($hkcuExists) {
        $props = Get-ItemProperty $hkcu
        $propCount = ($props.PSObject.Properties | Where-Object { $_.Name -notlike "PS*" }).Count
        Write-Host "    HKCU: LoadBehavior=$($props.LoadBehavior) ($propCount properties)" -ForegroundColor Yellow
        if ($propCount -le 1) {
            Write-Host "    >> PROBLEM: HKCU stub has no Manifest/FriendlyName -- overrides HKLM" -ForegroundColor Red
        }
    } else {
        Write-Host "    HKCU: not present (HKLM active)" -ForegroundColor Green
    }
}

# --- Diagnose mode ---
Write-Host ""
Write-Host "======================================" -ForegroundColor Cyan
Write-Host "  S&P COM Add-in Repair Tool" -ForegroundColor Cyan
Write-Host "======================================" -ForegroundColor Cyan
Write-Host ""

# Check Excel
$excelProcs = Get-Process Excel -ErrorAction SilentlyContinue
if ($excelProcs) {
    Write-Host "WARNING: Excel is running ($($excelProcs.Count) process(es))." -ForegroundColor Red
    if (-not $Diagnose) {
        Write-Host "Please close Excel before running repair." -ForegroundColor Red
        exit 1
    }
}

# Registry status
Write-Host "Registry status:" -ForegroundColor Cyan
Show-AddinStatus $snlProgId
Show-AddinStatus $spgmiProgId

# IsolatedStorage status
Write-Host ""
Write-Host "Encrypted storage:" -ForegroundColor Cyan
if (Test-Path $isoDir) {
    $fileCount = (Get-ChildItem $isoDir -Recurse -File -ErrorAction SilentlyContinue).Count
    Write-Host "  IsolatedStorage: EXISTS ($fileCount files)" -ForegroundColor Yellow
    Write-Host "    Path: $isoDir"
} else {
    Write-Host "  IsolatedStorage: not present (clean)" -ForegroundColor Green
}

# COM instantiation test
Write-Host ""
Write-Host "COM load test:" -ForegroundColor Cyan
try {
    $obj = New-Object -ComObject $snlProgId 2>$null
    Write-Host "  SNL add-in: LOADS OK" -ForegroundColor Green
    if ($obj) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) | Out-Null }
} catch {
    $msg = $_.Exception.Message
    if ($msg -match "CryptographicException|Padding") {
        Write-Host "  SNL add-in: FAILS (CryptographicException -- corrupted storage)" -ForegroundColor Red
    } elseif ($msg -match "80131604|ModuleInitialize") {
        Write-Host "  SNL add-in: FAILS (initialization error)" -ForegroundColor Red
    } else {
        Write-Host "  SNL add-in: FAILS ($($msg.Substring(0, [Math]::Min(120, $msg.Length))))" -ForegroundColor Red
    }
}

if ($Diagnose) {
    Write-Host ""
    Write-Host "Diagnose mode -- no changes made." -ForegroundColor Gray
    Write-Host "Run without -Diagnose to repair." -ForegroundColor Gray
    exit 0
}

# --- Repair ---
Write-Host ""
Write-Host "--- Repairing ---" -ForegroundColor Cyan

# Step 1: Clear corrupted IsolatedStorage
Write-Host ""
Write-Host "Step 1: Clear corrupted IsolatedStorage" -ForegroundColor Cyan
if (Test-Path $isoDir) {
    $backupName = "IsolatedStorage_SNL_backup_$(Get-Date -Format yyyyMMdd_HHmmss)"
    $backupPath = "$env:APPDATA\$backupName"
    Copy-Item $isoDir $backupPath -Recurse -Force
    Remove-Item $isoDir -Recurse -Force
    Write-Host "  Removed: $isoDir" -ForegroundColor Yellow
    Write-Host "  Backup:  $backupPath" -ForegroundColor Gray
} else {
    Write-Host "  Already clean." -ForegroundColor Green
}

# Also clear Local IsolatedStorage if present
$localIsoBase = "$env:LOCALAPPDATA\IsolatedStorage"
if (Test-Path $localIsoBase) {
    Get-ChildItem $localIsoBase -Directory -ErrorAction SilentlyContinue |
        Where-Object {
            # Look for SNL strong name directories
            Test-Path "$($_.FullName)\*\StrongName.bc4ly5pvcgxe2oetrztw1tyqiunlyrf4"
        } | ForEach-Object {
            Remove-Item $_.FullName -Recurse -Force
            Write-Host "  Removed Local IsolatedStorage: $($_.Name)" -ForegroundColor Yellow
        }
}

# Step 2: Remove HKCU stubs
Write-Host ""
Write-Host "Step 2: Remove HKCU registry stubs" -ForegroundColor Cyan
foreach ($progId in @($snlProgId, $spgmiProgId)) {
    $hkcu = "$hkcuBase\$progId"
    if (Test-Path $hkcu) {
        Remove-Item $hkcu -Force
        Write-Host "  Removed: $progId" -ForegroundColor Yellow
    } else {
        Write-Host "  Clean: $progId" -ForegroundColor Green
    }
}

# Step 3: Clear SNL Office cache (EdgeUserData, etc.)
Write-Host ""
Write-Host "Step 3: Clear SNL Office cache" -ForegroundColor Cyan
$snlOffice = "$env:APPDATA\SNL\Office"
if (Test-Path $snlOffice) {
    $backupName2 = "SNL_Office_backup_$(Get-Date -Format yyyyMMdd_HHmmss)"
    $backupPath2 = "$env:APPDATA\$backupName2"
    Copy-Item $snlOffice $backupPath2 -Recurse -Force
    Remove-Item $snlOffice -Recurse -Force
    Write-Host "  Removed: $snlOffice" -ForegroundColor Yellow
    Write-Host "  Backup:  $backupPath2" -ForegroundColor Gray
} else {
    Write-Host "  Already clean." -ForegroundColor Green
}

# Step 4 (optional): MSI repair
if ($Full) {
    Write-Host ""
    Write-Host "Step 4: MSI repair (S&P Capital IQ Pro Office)" -ForegroundColor Cyan
    $msiCache = "C:\ProgramData\Package Cache\${snlMsiId}v1.0.25336.1"
    if (Test-Path $msiCache) {
        Write-Host "  Running msiexec /fa ..." -ForegroundColor Yellow
        Start-Process msiexec -ArgumentList "/fa `"$snlMsiId`" /qn /norestart" -Verb RunAs -Wait
        Write-Host "  MSI repair completed." -ForegroundColor Green
    } else {
        Write-Host "  MSI cache not found at $msiCache" -ForegroundColor Yellow
        Write-Host "  Skipping MSI repair. Reinstall S&P Capital IQ if needed." -ForegroundColor Yellow
    }
}

# --- Done ---
Write-Host ""
Write-Host "======================================" -ForegroundColor Green
Write-Host "  Repair complete!" -ForegroundColor Green
Write-Host "  Open Excel and re-login to S&P." -ForegroundColor Green
Write-Host "======================================" -ForegroundColor Green
