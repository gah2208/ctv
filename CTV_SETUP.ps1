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
        param([string]$B64, [string]$PngPath, [string]$IcoPath)
        [IO.File]::WriteAllBytes($PngPath, [Convert]::FromBase64String($B64))
        Add-Type -AssemblyName System.Drawing
        $bmp    = [System.Drawing.Bitmap]::new($PngPath)
        $handle = $bmp.GetHicon()
        $icon   = [System.Drawing.Icon]::FromHandle($handle)
        $stream = [IO.FileStream]::new($IcoPath, [IO.FileMode]::Create)
        $icon.Save($stream)
        $stream.Close(); $icon.Dispose(); $bmp.Dispose()
        Remove-Item $PngPath -Force
    }

    Save-Icon "/9j/4AAQSkZJRgABAQAAAQABAAD/4gHYSUNDX1BST0ZJTEUAAQEAAAHIAAAAAAQwAABtbnRyUkdCIFhZWiAH4AABAAEAAAAAAABhY3NwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAA9tYAAQAAAADTLQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAlkZXNjAAAA8AAAACRyWFlaAAABFAAAABRnWFlaAAABKAAAABRiWFlaAAABPAAAABR3dHB0AAABUAAAABRyVFJDAAABZAAAAChnVFJDAAABZAAAAChiVFJDAAABZAAAAChjcHJ0AAABjAAAADxtbHVjAAAAAAAAAAEAAAAMZW5VUwAAAAgAAAAcAHMAUgBHAEJYWVogAAAAAAAAb6IAADj1AAADkFhZWiAAAAAAAABimQAAt4UAABjaWFlaIAAAAAAAACSgAAAPhAAAts9YWVogAAAAAAAA9tYAAQAAAADTLXBhcmEAAAAAAAQAAAACZmYAAPKnAAANWQAAE9AAAApbAAAAAAAAAABtbHVjAAAAAAAAAAEAAAAMZW5VUwAAACAAAAAcAEcAbwBvAGcAbABlACAASQBuAGMALgAgADIAMAAxADb/2wBDAAUDBAQEAwUEBAQFBQUGBwwIBwcHBw8LCwkMEQ8SEhEPERETFhwXExQaFRERGCEYGh0dHx8fExciJCIeJBweHx7/2wBDAQUFBQcGBw4ICA4eFBEUHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh7/wAARCABwAFoDASIAAhEBAxEB/8QAHAAAAgMBAQEBAAAAAAAAAAAABgcABAUDCAEC/8QASBAAAQIEAQYICgYIBwAAAAAAAQIDAAQFBhEHEiExQVEIEzY3dHWysxQWIjVTYXOBldMjMjNCYrEVNFJxkaHB0kNUY4KTlML/xAAbAQACAwEBAQAAAAAAAAAAAAAAAQIDBAUGB//EADIRAAIBAgMEBgoDAAAAAAAAAAECAAMRBDFBBRIhcQYTUWHB0RQVMjNCobHh8PEjgZH/2gAMAwEAAhEDEQA/APSdAt2QqNN8Nm5msrfdeeKymsTaB9qrUlLgAHqAwi/4oUf09b+OTnzYsWf5ga9q93q42BphkmRAFoP+KFI9NW/jk582P2LOo+16t/HJz5sECRgI+wXMdhB/xOo3pq38cnPmxPE6jemrfxyc+bBBEhXMLCD/AInUb01b+OTnzYnidRvTVv45OfNggiQXMLCD/idRvTVv45OfNieJ1G9NW/jk582NqempeRlHJqbdSyw2M5a1akjeYDqjlUs2UKkon3ZpQ2MMqIPvOAhFwuZmTE43C4X37heZAmubOo3pq38cnPmxmW3PTvi7TM6adWfBGsVLWVKV5A0knST6zpjvYt8SV3zU41IyUwy1KoSVLdI8oqJwGAx3GM+3OT1N6I12BEkYMLiPD4qjiaYq0TdTryhBZ/mBn2r3erjbSMIxrMGNAaP+q93q42oWs1DKSJEiQRyRIV+V7KEKS25Q6K6FT6hmvvJP2AOwfi/KKeR7KJ4TxVArz/0/1ZaZWftNyVHfuO2KuuXe3ZwW6R4Jcd6EW49ul+zn+o3Y+RIqVioydJpz0/PvpZl2U5ylK/Ibz6otvadx3VFLMbATlcNTp1JpExPVRxCJVCTn52nOx+6BtJ3R5VrT0tOVebmadJmWlnHFLbZBzsxMEl8XRVL6r7ctKtO+DBzMlJVGkkn7x3qP8oauTOzaPb6HJKeXLTdbfl8+ZbOCuKaUcM0DcTiMduB3Ria9drDIT5xjzU6U4rqcOLUk+IjM/mn9mZ3BzlOLtqoTpTgp+azAd4Skf1UY17c5PU3ojXYEFVt0ORoFPVIU9BRL8at1KScc3OOOH7hAtbnJ6m9Ea7AjZRXdW09psvBNgcFSw7ZqOPOEVmcnmfavd6uNiMezOTzPtXu9XGxD1nUGUkLbK3lARQmV0ekOJXU3E4OODSJcH/1uGzXHfKxf7duyyqZTFpcqrqdesMJO0+vcPfCxyb2ZPXjVFz08p1NPQ5nTD6j5TqtZSDv3nZGarVN9xM547bu2qr1fV2z+NRsyPh+/05z5k+tE15x6u118y1FliXZmYdXm8ZhpIzjs3mO2UO0JeSlmbntd5M5QZsBxDjSs7icdWn9n17NRgF4Q+Up2pzj1i2+2qn0GmOmXdQnyTMOIOBx/ACNA2nSdkbfA7rlRmq5U7Qm3RMUZyRXMeDujOCV56EnDcCFHEQDDLu21jXodhfQeob3h472t/Lu8Yz8l+U2XXT1U25ZoNPS7ZLU0v/ESBqP4vz/fARlBu+o3rWm5OTbdEklzNlZZOlTijoClDaT/ACizlIsNykTaKhQ0qnKTNLAaLZzy2onAI0axjoB90H+SmypO3VomamtlddfZLqGSQSw3iAcBvxIBPrwiq1R/4zpOElHbO0GGyq5sqe03aNOOvd88oG1Ofo2Re1v0nUUszt2TzZErK52PFj17kjaduoRk8Eat1O4r4u6sViaXNTkxLsqWtR/ErADcBqAhA35XKpcN21Gp1ebXMzK31JzlakpBICQNgA2Q7eBB5+uXorHaVGxVCiwn0PBYKjgaK0aIsB+XPfPUuyAO3OT1N6I12BB6dsAVucnqb0RrsCLFmhoRWZyeZ9q93q4G8ql+MWzKGRkVIdqryfJTrDIP3lf0EbttpmF2jmyjiGphRfDa1JzglXGrwJG2EtSrBuOu3pNSta45Aacz5ubVpCgdWadpOzdGeszDgus890gxuMo0koYNCXqcL9n3lGxrVnrwqz1QqD625BtRcnJtw69pAJ2/kIMMnuUmQr2VZuy7Uabbt2nSLyi6kfrDiSkYj8IxOnadMLDhC5RxLB7JtabapGlyKixPOJ0KfWNaN+bjrP3j6tebwPOd1XVj3aRDpUhTHfLthbDp7LpXPGo3tHw5RdZSucS5OtZnvVQ0OBjznz/VLneNwr8pXOLcnWsz3qoaHAx5z5/qlzvG4tnenKysstSsG57ipc5LrqlJXPTK2GCvAsu8YojNJ1JJ1j3jTrK+C9c9Xu/KtclbrUyXpl6nDAD6raeMGCEjYBHn+8+WFa6wf7xUObgU8ta51aO8TBCI2s+eJ3pDnaMP/gQefrl6Kx2lQgKz54nekOdow/8AgQefrl6Kx2lQQnqYwBW5yepvRGuwIPjADbnJ6m9Ea7AiSyLZwiszk8z7V7vVxrgRkWZyeZ9q93q42Iic4xlPAOWfnYujrR/tmDfgec7qurHu0iAjLPzsXR1o/wBswbcD3ndV1Y/2kQRxd5S+cW5OtZnvVQ0OBjznz/VLneNwr8pWnKJcnWsz3qoaHAx5z5/qlzvG4IRS3nywrXWD/eKhzcCnlrXOrR3iYTN58sK11g/3ioc3Ap5a1zq0d4mCERtZ88TvSHO0Yf8AwIPP1y9FY7SoQFZ88TvSHO0Yf/Ag8/XL0VjtKghPU5gBtzk9TeiNdgQenUYArc5PU3ojXYESWRbOa9Optz0+W8El5yjrZS44pBcl3M7BSyoA4Lwx0xY4q7v8zQ/+u7/fG/EiN47RO3Rkot96Yna/XabbnGPOKemH1+FDOUo6TgHdZJ1ARRtq3qXZdal6rTLcl6S7MlMo3OLkpoo+kUAEqBdJTic0YqSINMqMwW69ZzL5wkXKsC9j9UrA+jB95x90Gs4iWXKq8LDZYTgpXGfVGBBB9xAMW7oVQSM5o6tUVWbjf9RM3bkrtyXddrFVoVGmH5ya8riWpta3HXCToSl3acTojlZdKolsV6oPWxRJaXqcowG5xLdPnFrbQvBQBCnCNOaD7odrzDLymlOtpWWl57ZP3VYEYj3E/wAYCbMQBlWvlX7Qke6VDQKVJIy8xCkiMjE3uBf5geMC38lFt1GneMXi7RpsTqRNHMRNFxef5WObxuvTjhGhkmtumyco/cVkU+kSrcwhTKnCxM4uBKtICVunURuENuWYZlpduXl20tstpCEISMAlIGAAhaUCZetmtXXbMvgHXplM3S0nSMZghJwH7KV6T6sYEUODbPwjp0lqK1sx9L28oNUbJNa9fS/OM23SAnjVpUt5E2jOWFEKw+m04EHTq3QYWXk9VZr8y/bclQJFyZSlDxSiYXnAHEDynDvg6pUkzTabLyDGPFsNpbTicScBrPrOuLUVki/CUMBc7uUwsy7sP1mh/wDA7/fHOj285J0mTk3phDjrDCG1rQkhKilIBIGnAaIIYkK8jaSJEiQo4PX7JyNSoyadOyhmVPugMJSrNUlwalBWzCKK7erfgUqzP1+ZqDDDiFuMcWhBdCSCAVgYqww9WOGmCOoyxedl30jOUwvOw3jbHdb6QBmhSlEgYYRYKhCgCXrWZVCj88p02QG2zRJuUvitVA1V1xT3E+EILSAl3BBCcMB5OA3a4MHV5gScCQTgcNkUpVIaqM5MFKs13MwOadOAwhIxUEDWRpuVVgNR4zQgTqchIu5RafVls5z0pKqaLmAwSVnBJP8AEj/dBSleLQWoFOIBI3RQkpRLiZt15JC5hRx0aQBqgRt25hTfcue63+zSiRyllqWynP8ArgYK/fHWISqSJEiQQn//2Q=="   "$iconDir\start.png"   "$iconDir\start.ico"
    Save-Icon "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" "$iconDir\restart.png" "$iconDir\restart.ico"
    Save-Icon "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"  "$iconDir\config.png"  "$iconDir\config.ico"

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

    $python = (Get-Command python -ErrorAction SilentlyContinue).Source
    if (-not $python) { $python = "python.exe" }

    $runner = Get-ChildItem "C:\CTV" -Filter "ctv2_runner.py" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($runner) {
        Set-Shortcut "$desktop\Start Trading.lnk" $python `
            "`"$($runner.FullName)`"" $runner.DirectoryName "$iconDir\start.ico,0"
    }

    $editor = Get-ChildItem "C:\CTV" -Filter "ctv2_config_editor.jsx" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($editor) {
        Set-Shortcut "$desktop\Config Editor.lnk" "cmd.exe" `
            "/c start `"`" `"$($editor.FullName)`"" $editor.DirectoryName "$iconDir\config.ico,0"
    }

    $restartBat = Get-ChildItem "C:\CTV" -Filter "restart.bat" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($restartBat) {
        Set-Shortcut "$desktop\Restart Trading.lnk" "cmd.exe" `
            "/c `"$($restartBat.FullName)`"" $restartBat.DirectoryName "$iconDir\restart.ico,0"
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
