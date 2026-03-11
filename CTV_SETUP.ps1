# =============================================================
# CTV_SETUP.ps1 - CTV Trading Platform - Complete Setup
# =============================================================
# Installs Python, pip packages, and TWS if not present,
# then downloads and deploys the latest CTV release from GitHub.
# Run as Administrator for best results.
# =============================================================

$ErrorActionPreference = "SilentlyContinue"

Write-Host ""
Write-Host "============================================================"
Write-Host "   CTV TRADING PLATFORM - SETUP"
Write-Host "============================================================"
Write-Host ""

# Check for admin rights
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")
if (-not $isAdmin) {
    Write-Host "WARNING: Not running as Administrator. Some steps may fail."
    Write-Host ""
}

$downloadDir = "$env:TEMP\ctv_setup"
New-Item -ItemType Directory -Path $downloadDir -Force | Out-Null

# =============================================================
# STEP 1: Python 3.11.9
# =============================================================
Write-Host "[1/4] Checking Python..."

$pythonInstalled = $false
try {
    $ver = & python --version 2>&1
    if ($ver -match "Python") {
        Write-Host "      Already installed: $ver"
        $pythonInstalled = $true
    }
} catch {}

if (-not $pythonInstalled) {
    Write-Host "      Not found - downloading Python 3.11.9..."
    $pyUrl       = "https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe"
    $pyInstaller = "$downloadDir\python-3.11.9-amd64.exe"
    Invoke-WebRequest -Uri $pyUrl -OutFile $pyInstaller -UseBasicParsing | Out-Null

    if (Test-Path $pyInstaller) {
        Write-Host "      Installing Python 3.11.9..."
        Start-Process -FilePath $pyInstaller `
            -ArgumentList "/quiet InstallAllUsers=1 PrependPath=1 Include_test=0" `
            -Wait
        Start-Sleep -Seconds 15
        $env:PATH = [System.Environment]::GetEnvironmentVariable("PATH","Machine") + ";" +
                    [System.Environment]::GetEnvironmentVariable("PATH","User")
        Write-Host "      Python installed."
    } else {
        Write-Host "      ERROR: Failed to download Python."
        Write-Host "      Please install manually from https://www.python.org/downloads/"
        Write-Host "      then re-run this script."
        Read-Host "Press Enter to exit"
        exit 1
    }
}

# =============================================================
# STEP 2: Python libraries
# =============================================================
Write-Host "[2/4] Installing Python libraries..."
foreach ($pkg in @("ibapi","pyyaml","pytz","tzdata","pyinstaller")) {
    & pip install $pkg --quiet 2>$null | Out-Null
}
Write-Host "      Done."

# =============================================================
# STEP 3: Trader Workstation
# =============================================================
Write-Host "[3/4] Checking for Trader Workstation..."

$twsInstalled = (Test-Path "C:\Jts\tws.exe") -or
                (Test-Path "$env:LOCALAPPDATA\Jts\tws.exe")

if ($twsInstalled) {
    $twsPath = if (Test-Path "C:\Jts\tws.exe") { "C:\Jts\tws.exe" } `
               else { "$env:LOCALAPPDATA\Jts\tws.exe" }
    Write-Host "      Already installed: $twsPath"
} else {
    Write-Host "      Not found - downloading TWS..."
    $twsUrl       = "https://download2.interactivebrokers.com/installers/tws/latest-standalone/tws-latest-standalone-windows-x64.exe"
    $twsInstaller = "$downloadDir\tws-installer.exe"
    Invoke-WebRequest -Uri $twsUrl -OutFile $twsInstaller -UseBasicParsing | Out-Null

    if (Test-Path $twsInstaller) {
        Write-Host "      Launching TWS installer - follow the prompts..."
        Start-Process -FilePath $twsInstaller -Wait
        Write-Host "      TWS installed."
    } else {
        Write-Host "      ERROR: Failed to download TWS."
        Write-Host "      Please install manually from https://www.interactivebrokers.com/en/trading/tws.php"
        Read-Host "Press Enter to exit"
        exit 1
    }
}

# =============================================================
# STEP 4: Latest CTV Release
# =============================================================
Write-Host "[4/4] CTV Release..."
Write-Host ""
$answer = Read-Host "      Do you want to get the latest CTV files? (Y/N)"

if ($answer -notmatch "^[Yy]") {
    Write-Host "      Skipped."
} else {
    $zipUrl      = "https://github.com/gah2208/ctv/releases/latest/download/ctv.zip"
    $zipPath     = "$env:USERPROFILE\Downloads\ctv.zip"

    Write-Host "      Downloading latest release from GitHub..."
    Invoke-WebRequest -Uri $zipUrl -OutFile $zipPath -UseBasicParsing | Out-Null

    if (-not (Test-Path $zipPath)) {
        Write-Host "      ERROR: Download failed. Check your internet connection."
        Read-Host "Press Enter to exit"
        exit 1
    }

    Write-Host "      Extracting to C:\CTV..."
    if (-not (Test-Path "C:\CTV")) { New-Item -ItemType Directory -Path "C:\CTV" -Force | Out-Null }
    Expand-Archive -Path $zipPath -DestinationPath "C:\CTV" -Force | Out-Null

    # ---------------------------------------------------------
    # Icons
    # ---------------------------------------------------------
    Write-Host "      Setting up icons and shortcuts..."
    $iconDir = "C:\CTV\icons"
    New-Item -ItemType Directory -Path $iconDir -Force | Out-Null

    function Save-Icon {
        param([string]$B64, [string]$IcoPath)
        # B64 is a pre-built multi-size ICO (16/24/32/48/256px). Write directly.
        [IO.File]::WriteAllBytes($IcoPath, [Convert]::FromBase64String($B64))
    }

    Save-Icon 

    # ---------------------------------------------------------
    # Desktop (handles OneDrive-redirected desktops)
    # ---------------------------------------------------------
    $desktop = (Get-ItemProperty "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders").Desktop
    if (-not $desktop) { $desktop = [Environment]::GetFolderPath("Desktop") }
    $desktop = [System.Environment]::ExpandEnvironmentVariables($desktop)

    function Set-Shortcut {
        param([string]$LnkPath, [string]$Target, [string]$Arguments,
              [string]$WorkDir, [string]$IconLocation)
        $wsh = New-Object -ComObject WScript.Shell
        $sc  = $wsh.CreateShortcut($LnkPath)
        $sc.TargetPath        = $Target
        $sc.Arguments         = $Arguments
        $sc.WorkingDirectory  = $WorkDir
        $sc.IconLocation      = $IconLocation
        $sc.Save()
        [Runtime.InteropServices.Marshal]::ReleaseComObject($wsh) | Out-Null
    }

    $startPs1 = Get-ChildItem "C:\CTV" -Filter "START_TRADING.ps1" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($startPs1) {
        Set-Shortcut "$desktop\Start Trading.lnk" "powershell.exe" `
            "-ExecutionPolicy Bypass -File `"$($startPs1.FullName)`"" $startPs1.DirectoryName "$iconDir\start.ico,0"
    }

    $editor = Get-ChildItem "C:\CTV" -Filter "ctv2_config_editor.py" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($editor) {
        Set-Shortcut "$desktop\Config Editor.lnk" "cmd.exe" `
            "/c start `"`" `"$($editor.FullName)`"" $editor.DirectoryName "$iconDir\config.ico,0"
    }

    $restartPs1 = Get-ChildItem "C:\CTV" -Filter "RESTART.ps1" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($restartPs1) {
        Set-Shortcut "$desktop\Restart Trading.lnk" "powershell.exe" `
            "-ExecutionPolicy Bypass -File `"$($restartPs1.FullName)`"" $restartPs1.DirectoryName "$iconDir\restart.ico,0"
    }

    Write-Host "      Shortcuts updated on desktop."
}

# Cleanup
if (Test-Path $downloadDir) { Remove-Item -Path $downloadDir -Recurse -Force | Out-Null }

Write-Host ""
Write-Host "============================================================"
Write-Host "   CTV SETUP COMPLETE"
Write-Host "============================================================"
Write-Host ""
Read-Host "Press Enter to exit"
