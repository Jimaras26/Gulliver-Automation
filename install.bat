@echo off
setlocal EnableDelayedExpansion

REM ============================================================
REM  Bibecoffee Production Tool - One-Click Installer
REM  Installs to C:\GulliverTool
REM ============================================================

REM Require Administrator privileges
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo  ERROR: Administrator privileges required.
    echo  Right-click install.bat and select "Run as administrator".
    echo.
    pause
    exit /b 1
)

set DEST=C:\GulliverTool

echo.
echo  ============================================================
echo   Bibecoffee Production Tool - Installer
echo  ============================================================
echo.

REM ---- Step 1: Install SEGGER JLink ----
echo  [1/4] Installing SEGGER JLink...
"%~dp0installers\JLink_Windows_V866_x86_64.exe" /S
if !errorlevel! neq 0 (
    echo         WARNING: JLink installer exited with code !errorlevel!.
    echo         It may already be installed. Continuing...
) else (
    echo         Done.
)
echo.

REM ---- Step 2: Install Brother P-touch Editor ----
echo  [2/4] Installing Brother P-touch Editor...
REM puw10029.exe uses /S for unattended install (NSIS-based)
"%~dp0installers\puw10029.exe" /S
if !errorlevel! neq 0 (
    echo         WARNING: P-touch installer exited with code !errorlevel!.
    echo         If a setup window appeared, complete it manually and re-run.
) else (
    echo         Done.
)
echo.

REM ---- Step 3: Copy tool files to C:\GulliverTool ----
echo  [3/4] Copying files to %DEST%...
if not exist "%DEST%" mkdir "%DEST%"
xcopy "%~dp0files\*" "%DEST%\" /E /I /Y /Q
if !errorlevel! neq 0 (
    echo  ERROR: Failed to copy files to %DEST%. Aborting.
    pause
    exit /b 1
)
echo         Done.
echo.

REM ---- Step 4: Create Desktop Shortcut ----
echo  [4/4] Creating Desktop shortcut...
powershell -NoProfile -Command ^
  "$s = (New-Object -ComObject WScript.Shell).CreateShortcut([Environment]::GetFolderPath('Desktop') + '\Bibecoffee Production Tool.lnk');" ^
  "$s.TargetPath = '%DEST%\BibecoffeeProductionTool.exe';" ^
  "$s.WorkingDirectory = '%DEST%';" ^
  "$s.Save()"
echo         Done.
echo.

echo  ============================================================
echo   Installation complete!
echo   Shortcut created on Desktop.
echo   Tool installed at: %DEST%
echo  ============================================================
echo.
pause
