# package.ps1
# Assembles GulliverTool_Setup.zip ready to deploy on a new machine.
#
# USAGE:
#   1. Build the main app:     pyinstaller gui2.spec
#   2. Build standalone esptool: pyinstaller esptool.spec
#   3. Run this script:        .\package.ps1
#
# OUTPUT: GulliverTool_Setup.zip on your Desktop

$ErrorActionPreference = "Stop"

$WORKSPACE  = $PSScriptRoot
$FW_BASE    = "C:\Users\DimitrisOikonomou\Desktop\Gulliver_Testing"
$INSTALLERS = "C:\Users\DimitrisOikonomou\Downloads"
$STAGING    = "$env:TEMP\GulliverTool_Staging"
$OUTPUT_ZIP = "$env:USERPROFILE\Desktop\GulliverTool_Setup.zip"
# Fixed install destination used inside flash.txt and MainConfig.ini
$INSTALL_DEST = "C:\GulliverTool"

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " GulliverTool - Package Builder" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# ---- Step 1: Check build outputs ----
Write-Host "[1/7] Checking build outputs..." -ForegroundColor Yellow

$mainExe   = "$WORKSPACE\dist\BibecoffeeProductionTool.exe"
$esptoolExe = "$WORKSPACE\dist\esptool.exe"

if (-not (Test-Path $mainExe)) {
    Write-Error "BibecoffeeProductionTool.exe not found.`nRun first: pyinstaller gui2.spec"
}
if (-not (Test-Path $esptoolExe)) {
    Write-Error "esptool.exe not found.`nRun first: pyinstaller esptool.spec"
}

Write-Host "      Main EXE  : $mainExe" -ForegroundColor Green
Write-Host "      esptool   : $esptoolExe" -ForegroundColor Green

# ---- Step 2: Create clean staging folder ----
Write-Host "[2/7] Setting up staging folder..." -ForegroundColor Yellow

if (Test-Path $STAGING) { Remove-Item $STAGING -Recurse -Force }
$null = New-Item -ItemType Directory -Path "$STAGING\files\logs"
$null = New-Item -ItemType Directory -Path "$STAGING\installers"

Write-Host "      Staging: $STAGING" -ForegroundColor Green

# ---- Step 3: Copy executables ----
Write-Host "[3/7] Copying executables..." -ForegroundColor Yellow

Copy-Item $mainExe    "$STAGING\files\"
Copy-Item $esptoolExe "$STAGING\files\"

# ---- Step 4: Copy firmware and tool folders ----
Write-Host "[4/7] Copying firmware and tools..." -ForegroundColor Yellow

Copy-Item "$FW_BASE\ESPFW"      "$STAGING\files\ESPFW"      -Recurse
Copy-Item "$FW_BASE\bg95update" "$STAGING\files\bg95update" -Recurse
Copy-Item "$FW_BASE\QFlash_V7.7" "$STAGING\files\QFlash_V7.7" -Recurse

# ---- Step 5: Copy loose files and assets ----
Write-Host "[5/7] Copying assets and firmware binaries..." -ForegroundColor Yellow

$looseFiles = @(
    "coffeeBean.png",
    "Gulliver_Label.lbx",
    "erase.txt",
    "check_blank.txt",
    "Gulliver_Barista_18_jtag.bin",
    "Gulliver_Barista_19_jtag_Ryoma.bin",
    "GULLIVER_V.54.19.260316_JTAG.bin"
)
foreach ($f in $looseFiles) {
    $src = "$FW_BASE\$f"
    if (Test-Path $src) {
        Copy-Item $src "$STAGING\files\"
    } else {
        Write-Warning "File not found, skipping: $src"
    }
}

# Generate flash.txt pointing to the fixed install destination
$flashTxt = @"
device ATSAMD21J18A
if SWD
speed 4000
r
halt
loadfile "$INSTALL_DEST\Gulliver_Barista_19_jtag_Ryoma.bin",0x00000000
r
g
exit
"@
$flashTxt | Set-Content "$STAGING\files\flash.txt" -Encoding UTF8

# Update QFlash MainConfig.ini: replace dev machine path with install destination
$qflashConfig = "$STAGING\files\QFlash_V7.7\MainConfig.ini"
(Get-Content $qflashConfig -Raw) `
    -replace [regex]::Escape("C:\Users\DimitrisOikonomou\Desktop\Gulliver_Testing"), $INSTALL_DEST |
    Set-Content $qflashConfig -Encoding UTF8 -NoNewline

Write-Host "      flash.txt  → target: $INSTALL_DEST" -ForegroundColor Green
Write-Host "      QFlash MainConfig.ini updated" -ForegroundColor Green

# ---- Step 6: Copy installers and install.bat ----
Write-Host "[6/7] Copying installers..." -ForegroundColor Yellow

Copy-Item "$INSTALLERS\JLink_Windows_V866_x86_64.exe" "$STAGING\installers\"
Copy-Item "$INSTALLERS\pew54007us.exe"                "$STAGING\installers\"
Copy-Item "$WORKSPACE\install.bat"                    "$STAGING\"

# ---- Step 7: Create ZIP ----
Write-Host "[7/7] Creating ZIP..." -ForegroundColor Yellow

if (Test-Path $OUTPUT_ZIP) { Remove-Item $OUTPUT_ZIP -Force }
Compress-Archive -Path "$STAGING\*" -DestinationPath $OUTPUT_ZIP -CompressionLevel Optimal

# ---- Done ----
$zipSize = [math]::Round((Get-Item $OUTPUT_ZIP).Length / 1MB, 1)
Write-Host ""
Write-Host "============================================================" -ForegroundColor Green
Write-Host " Package ready!" -ForegroundColor Green
Write-Host " Output : $OUTPUT_ZIP" -ForegroundColor Green
Write-Host " Size   : $zipSize MB" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host ""
Write-Host "To deploy on a new machine:" -ForegroundColor Cyan
Write-Host "  1. Copy GulliverTool_Setup.zip to the new PC" -ForegroundColor Cyan
Write-Host "  2. Extract the ZIP" -ForegroundColor Cyan
Write-Host "  3. Right-click install.bat → Run as Administrator" -ForegroundColor Cyan
Write-Host ""
