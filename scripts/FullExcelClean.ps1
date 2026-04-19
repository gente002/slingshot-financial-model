# FullExcelClean.ps1 - Kill all Excel processes and clear all recovery/crash artifacts

# --- Kill ALL Excel processes ---
$excelProcs = Get-Process Excel -ErrorAction SilentlyContinue
if ($excelProcs) {
    Write-Host "Killing $($excelProcs.Count) Excel process(es)..." -ForegroundColor Yellow
    $excelProcs | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 3
} else {
    Write-Host "No Excel processes running." -ForegroundColor Green
}

$appdata = [Environment]::GetFolderPath('ApplicationData')
$localappdata = [Environment]::GetFolderPath('LocalApplicationData')

# --- Registry: Resiliency (StartupItems, DocumentRecovery, DisabledItems) ---
foreach ($ver in @("16.0", "15.0")) {
    $resPath = "HKCU:\Software\Microsoft\Office\$ver\Excel\Resiliency"
    if (Test-Path $resPath) {
        # StartupItems
        $siPath = "$resPath\StartupItems"
        if (Test-Path $siPath) {
            $props = (Get-Item $siPath).Property
            if ($props) {
                foreach ($p in @($props)) {
                    Remove-ItemProperty -Path $siPath -Name $p -Force -ErrorAction SilentlyContinue
                }
                Write-Host "Cleared $($props.Count) StartupItems ($ver)" -ForegroundColor Yellow
            }
        }

        # DocumentRecovery
        $drPath = "$resPath\DocumentRecovery"
        if (Test-Path $drPath) {
            $entries = Get-ChildItem $drPath -ErrorAction SilentlyContinue
            if ($entries) {
                foreach ($e in $entries) {
                    Remove-Item $e.PSPath -Recurse -Force -ErrorAction SilentlyContinue
                }
                Write-Host "Cleared $($entries.Count) DocumentRecovery entries ($ver)" -ForegroundColor Yellow
            }
        }

        # DisabledItems
        $diPath = "$resPath\DisabledItems"
        if (Test-Path $diPath) {
            $props = (Get-Item $diPath).Property
            if ($props) {
                foreach ($p in @($props)) {
                    Remove-ItemProperty -Path $diPath -Name $p -Force -ErrorAction SilentlyContinue
                }
                Write-Host "Cleared $($props.Count) DisabledItems ($ver)" -ForegroundColor Yellow
            }
        }
    }
}

# --- Disk: ALL AutoRecover folders (not XLSTART or Personal.xlsb) ---
$excelDir = "$appdata\Microsoft\Excel"
if (Test-Path $excelDir) {
    $recoveryFolders = Get-ChildItem $excelDir -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -ne "XLSTART" }
    foreach ($folder in $recoveryFolders) {
        Remove-Item $folder.FullName -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "Removed recovery folder: $($folder.Name)" -ForegroundColor Yellow
    }

    # Remove (version 1).xlsb recovery files
    $recoveryFiles = Get-ChildItem $excelDir -File -Filter "*(version*).xlsb" -ErrorAction SilentlyContinue
    foreach ($f in $recoveryFiles) {
        Remove-Item $f.FullName -Force -ErrorAction SilentlyContinue
        Write-Host "Removed recovery file: $($f.Name)" -ForegroundColor Yellow
    }
}

# --- Disk: UnsavedFiles ---
$unsaved = "$localappdata\Microsoft\Office\UnsavedFiles"
if (Test-Path $unsaved) {
    $files = Get-ChildItem $unsaved -ErrorAction SilentlyContinue
    if ($files) {
        foreach ($f in $files) {
            Remove-Item $f.FullName -Force -ErrorAction SilentlyContinue
        }
        Write-Host "Cleared $($files.Count) UnsavedFiles" -ForegroundColor Yellow
    }
}

# --- Disk: XLSTART lock files ---
$xlstart = "$appdata\Microsoft\Excel\XLSTART"
if (Test-Path $xlstart) {
    $locks = Get-ChildItem $xlstart -Force -Filter '~$*' -ErrorAction SilentlyContinue
    foreach ($lf in $locks) {
        Remove-Item $lf.FullName -Force -ErrorAction SilentlyContinue
        Write-Host "Removed XLSTART lock: $($lf.Name)" -ForegroundColor Yellow
    }
}

# --- Disk: Temp files from RDK ---
$temp = [System.IO.Path]::GetTempPath()
Get-ChildItem $temp -Filter "rdk_*" -ErrorAction SilentlyContinue | ForEach-Object {
    Remove-Item $_.FullName -Force -ErrorAction SilentlyContinue
    Write-Host "Removed temp: $($_.Name)" -ForegroundColor Yellow
}

# --- Disk: workbook lock file ---
$wbLock = "C:\Users\gente\Downloads\RDK_Phase11B_Consolidated\workbook\~$RDK_Model.xlsm"
if (Test-Path $wbLock) {
    Remove-Item $wbLock -Force -ErrorAction SilentlyContinue
    Write-Host "Removed workbook lock file" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "All clean. Ready for a fresh Setup.bat run." -ForegroundColor Green
