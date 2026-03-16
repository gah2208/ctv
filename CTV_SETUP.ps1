# =============================================================
# CTV_SETUP.ps1 - CTV Trading Platform - Complete Setup  v1.11
# v1.10 - Excluded config/editor from ctv_files; dot-source shortcuts fix
# v1.11 - Removed ctv_files from manifest entirely; only python + tws checked
# =============================================================
# Installs Python, pip packages, and TWS if not present,
# then downloads and deploys the latest CTV release from GitHub.
# Run as Administrator for best results.
# =============================================================

$ErrorActionPreference = "SilentlyContinue"

Write-Output ""
Write-Output "============================================================"
Write-Output "   CTV TRADING PLATFORM - SETUP"
Write-Output "============================================================"
Write-Output ""

# Check for admin rights
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")
if (-not $isAdmin) {
    Write-Output "WARNING: Not running as Administrator. Some steps may fail."
    Write-Output ""
}

$downloadDir = "$env:TEMP\ctv_setup"
New-Item -ItemType Directory -Path $downloadDir -Force | Out-Null

# =============================================================
# STEP 1: Python 3.11.9
# =============================================================
Write-Output "[1/4] Checking Python..."

$pythonInstalled = $false
try {
    $ver = & python --version 2>&1
    if ($ver -match "Python") {
        Write-Output "      Already installed: $ver"
        $pythonInstalled = $true
    }
} catch {}

if (-not $pythonInstalled) {
    Write-Output "      Not found - downloading Python 3.11.9..."
    $pyUrl       = "https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe"
    $pyInstaller = "$downloadDir\python-3.11.9-amd64.exe"
    Invoke-WebRequest -Uri $pyUrl -OutFile $pyInstaller -UseBasicParsing | Out-Null

    if (Test-Path $pyInstaller) {
        Write-Output "      Installing Python 3.11.9..."
        Start-Process -FilePath $pyInstaller `
            -ArgumentList "/quiet InstallAllUsers=1 PrependPath=1 Include_test=0" `
            -Wait
        Start-Sleep -Seconds 15
        $env:PATH = [System.Environment]::GetEnvironmentVariable("PATH","Machine") + ";" +
                    [System.Environment]::GetEnvironmentVariable("PATH","User")
        Write-Output "      Python installed."
    } else {
        Write-Output "      ERROR: Failed to download Python."
        Write-Output "      Please install manually from https://www.python.org/downloads/"
        Write-Output "      then re-run this script."
        Read-Host "Press Enter to exit"
        exit 1
    }
}

# =============================================================
# STEP 2: Python libraries
# =============================================================
Write-Output "[2/4] Checking Python libraries..."
foreach ($pkg in @("ibapi","pyyaml","pytz","tzdata","pyinstaller")) {
    $installed = & pip show $pkg 2>$null
    if (-not $installed) {
        Write-Output "      Installing $pkg..."
        & pip install $pkg --quiet 2>$null | Out-Null
    } else {
        Write-Output "      Already installed: $pkg"
    }
}
Write-Output "      Done."

# =============================================================
# STEP 3: Trader Workstation
# =============================================================
Write-Output "[3/4] Checking for Trader Workstation..."

$twsInstalled = (Test-Path "C:\Jts\tws.exe") -or
                (Test-Path "$env:LOCALAPPDATA\Jts\tws.exe")

if ($twsInstalled) {
    $twsPath = if (Test-Path "C:\Jts\tws.exe") { "C:\Jts\tws.exe" } `
               else { "$env:LOCALAPPDATA\Jts\tws.exe" }
    Write-Output "      Already installed: $twsPath"
} else {
    Write-Output "      Not found - downloading TWS..."
    $twsUrl       = "https://download2.interactivebrokers.com/installers/tws/latest-standalone/tws-latest-standalone-windows-x64.exe"
    $twsInstaller = "$downloadDir\tws-installer.exe"
    Invoke-WebRequest -Uri $twsUrl -OutFile $twsInstaller -UseBasicParsing | Out-Null

    if (Test-Path $twsInstaller) {
        Write-Output "      Launching TWS installer - follow the prompts..."
        Start-Process -FilePath $twsInstaller -Wait
        Write-Output "      TWS installed."
    } else {
        Write-Output "      ERROR: Failed to download TWS."
        Write-Output "      Please install manually from https://www.interactivebrokers.com/en/trading/tws.php"
        Read-Host "Press Enter to exit"
        exit 1
    }
}

# =============================================================
# STEP 4: Latest CTV Release
# =============================================================
Write-Output "[4/5] CTV Release..."

$zipUrl  = "https://github.com/gah2208/ctv/releases/latest/download/ctv.zip"
$zipPath = "$env:USERPROFILE\Downloads\ctv.zip"

Write-Output "      Downloading latest release from GitHub..."
Invoke-WebRequest -Uri $zipUrl -OutFile $zipPath -UseBasicParsing | Out-Null

if (-not (Test-Path $zipPath)) {
    Write-Output "      ERROR: Download failed. Check your internet connection."
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Output "      Extracting to C:\CTV..."
if (-not (Test-Path "C:\CTV")) { New-Item -ItemType Directory -Path "C:\CTV" -Force | Out-Null }
Expand-Archive -Path $zipPath -DestinationPath "C:\CTV" -Force | Out-Null

# ---------------------------------------------------------
# Shortcuts (delegated to CREATE_SHORTCUTS.ps1)
# ---------------------------------------------------------
Write-Output "      Setting up desktop shortcuts..."
$shortcutsScript = "C:\CTV\CREATE_SHORTCUTS.ps1"
if (Test-Path $shortcutsScript) {
    . $shortcutsScript
} else {
    Write-Output "      WARNING: CREATE_SHORTCUTS.ps1 not found in C:\CTV - shortcuts not created."
}

# =============================================================
# STEP 5: Generate install manifest
# =============================================================
Write-Output "[5/5] Generating install manifest..."

$pyExe  = (Get-Command python -ErrorAction SilentlyContinue).Source
$twsExe = if (Test-Path "C:\Jts\tws.exe") { "C:\Jts\tws.exe" } `
          elseif (Test-Path "$env:LOCALAPPDATA\Jts\tws.exe") { "$env:LOCALAPPDATA\Jts\tws.exe" } `
          else { $null }

function Get-SHA256 { param([string]$Path)
    if (-not (Test-Path $Path)) { return $null }
    return (Get-FileHash -Path $Path -Algorithm SHA256).Hash
}

$manifest = @{
    generated = (Get-Date -Format "yyyy-MM-ddTHH:mm:ss")
    python    = @{ path = $pyExe;  sha256 = (Get-SHA256 $pyExe)  }
    tws       = @{ path = $twsExe; sha256 = (Get-SHA256 $twsExe) }
}

$manifestPath = "C:\CTV\install_manifest.json"
$manifest | ConvertTo-Json -Depth 5 | Set-Content $manifestPath -Encoding UTF8
Write-Output "      Manifest written to $manifestPath"

# Cleanup
if (Test-Path $downloadDir) { Remove-Item -Path $downloadDir -Recurse -Force | Out-Null }

Write-Output ""
Write-Output "============================================================"
Write-Output "   CTV SETUP COMPLETE"
Write-Output "============================================================"
Write-Output ""

if (-not $twsInstalled) {
    Write-Output "*** IMPORTANT - Before running Start Trading ***"
    Write-Output "You must configure the TWS API settings in Trader Workstation."
    Write-Output "Opening the TWS API Setup Guide now..."
    Write-Output ""
    $pdfPath = "C:\CTV\TWS_API_SETUP.pdf"
    if (Test-Path $pdfPath) {
        Start-Process $pdfPath
    } else {
        Write-Output "NOTE: TWS_API_SETUP.pdf not found in C:\CTV - please review API settings manually."
    }
}
Read-Host "Press Enter to exit"
