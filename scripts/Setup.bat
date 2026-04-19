@echo off
setlocal EnableDelayedExpansion

:: ============================================================================
:: Setup.bat - RDK Unified Bootstrap (double-click to run)
:: Single menu. Supports config-driven workbook naming.
:: ============================================================================

set "ROOT=%~dp0.."
set "CONFIG=%ROOT%\config"

:menu
echo ========================================
echo  RDK Setup
echo  Kernel Version: 1.3.0
echo ========================================
echo.

echo   [1] Sample model (toy -- 3 entities, Revenue/COGS)  (default)
echo   [2] Insurance FM -- full model (Run Model + actuarial curves)
echo   [3] Fresh start -- blank templates
echo   [4] Rebuild workbook only -- keep config, scenarios, savepoints
echo   [q] Quit
echo.
set "INPUT="
set /P "INPUT=Choose [1-4/q, default=1]: "
if "!INPUT!"=="" set "INPUT=1"
set "CHOICE=!INPUT!"
if /I "!CHOICE!"=="q" goto :cancelled
if "!CHOICE!" NEQ "1" if "!CHOICE!" NEQ "2" if "!CHOICE!" NEQ "3" if "!CHOICE!" NEQ "4" (
    echo Invalid choice. Using default ^(1^).
    set "CHOICE=1"
)

:: -------------------------------------------------------
:: Reset (options 1, 2, and 3)
:: -------------------------------------------------------
if %CHOICE% LEQ 3 (
    :: Always clear transient state
    if exist "%ROOT%\scenarios"   rmdir /S /Q "%ROOT%\scenarios"
    if exist "%ROOT%\savepoints"  rmdir /S /Q "%ROOT%\savepoints"
    if exist "%ROOT%\snapshots"   rmdir /S /Q "%ROOT%\snapshots"
    if exist "%ROOT%\archive"     rmdir /S /Q "%ROOT%\archive"
    if exist "%ROOT%\wal"         rmdir /S /Q "%ROOT%\wal"
    if exist "%ROOT%\output"      rmdir /S /Q "%ROOT%\output"
    if exist "%CONFIG%"           rmdir /S /Q "%CONFIG%"
    del /Q "%ROOT%\workbook\~$*" 2>nul

    :: Protect workspaces -- ask before deleting
    if exist "%ROOT%\workspaces" (
        echo.
        echo  WARNING: Found existing workspaces directory.
        echo  Workspaces contain saved model scenarios and inputs.
        echo.
        set "WS_DEL="
        set /P "WS_DEL=  Delete all workspaces? [y/N, default=N]: "
        if /I "!WS_DEL!"=="y" (
            rmdir /S /Q "%ROOT%\workspaces"
            echo  Workspaces deleted.
        ) else (
            echo  Workspaces preserved.
        )
    )
    echo Cleared: scenarios, snapshots, archive, wal, output, config, workbook
)

:: -------------------------------------------------------
:: Seed config
:: -------------------------------------------------------
if %CHOICE%==1 (
    if not exist "%ROOT%\config_sample" (
        echo ERROR: config_sample\ not found.
        goto :fail
    )
    xcopy "%ROOT%\config_sample" "%CONFIG%\" /E /I /Q >nul
    echo Seeded config\ from config_sample\
)
if %CHOICE%==2 (
    if not exist "%ROOT%\config_insurance" (
        echo ERROR: config_insurance\ not found.
        goto :fail
    )
    xcopy "%ROOT%\config_insurance" "%CONFIG%\" /E /I /Q >nul
    echo Seeded config\ from config_insurance\
)
if %CHOICE%==3 (
    if not exist "%ROOT%\config_blank" (
        echo ERROR: config_blank\ not found.
        goto :fail
    )
    xcopy "%ROOT%\config_blank" "%CONFIG%\" /E /I /Q >nul
    echo Seeded config\ from config_blank\
)
if %CHOICE%==4 (
    if not exist "%CONFIG%" (
        echo ERROR: No config\ directory. Use option 1, 2, or 3 first.
        goto :fail
    )
    echo Keeping existing config, scenarios, and savepoints
)

:: -------------------------------------------------------
:: Read WorkbookName from branding_config (default: RDK_Model)
:: -------------------------------------------------------
set "WB_NAME=RDK_Model"
if exist "%CONFIG%\branding_config.csv" (
    for /F "tokens=2 delims=," %%A in ('findstr /I "WorkbookName" "%CONFIG%\branding_config.csv"') do (
        set "RAW=%%~A"
        if defined RAW set "WB_NAME=!RAW!"
    )
)
echo Workbook name: !WB_NAME!.xlsm

:: Clean up old workbooks if rebuilding
if %CHOICE% LEQ 3 (
    if exist "%ROOT%\workbook\!WB_NAME!.xlsm" del /Q "%ROOT%\workbook\!WB_NAME!.xlsm"
)

:: -------------------------------------------------------
:: Bootstrap
:: -------------------------------------------------------
echo.
powershell -ExecutionPolicy Bypass -File "%~dp0Bootstrap.ps1" -WorkbookName "!WB_NAME!"
if %ERRORLEVEL% NEQ 0 goto :fail

echo.
echo ========================================
echo  Setup complete!
echo  Double-click workbook\!WB_NAME!.xlsm
echo  to open, then run 'RunProjections'.
echo ========================================
goto :done

:cancelled
echo Exiting.
goto :exit

:fail
echo.
echo Setup failed. See errors above.

:done
echo.
echo Press Enter to return to the menu, or type q to quit.
set "AGAIN="
set /P "AGAIN="
if /I "!AGAIN!"=="q" goto :exit
echo.
goto :menu

:exit
