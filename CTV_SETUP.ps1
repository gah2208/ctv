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
        $src    = [System.Drawing.Bitmap]::new($PngPath)
        $bmp    = [System.Drawing.Bitmap]::new(256, 256)
        $g      = [System.Drawing.Graphics]::FromImage($bmp)
        $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $g.DrawImage($src, 0, 0, 256, 256)
        $g.Dispose(); $src.Dispose()
        $handle = $bmp.GetHicon()
        $icon   = [System.Drawing.Icon]::FromHandle($handle)
        $stream = [IO.FileStream]::new($IcoPath, [IO.FileMode]::Create)
        $icon.Save($stream)
        $stream.Close(); $icon.Dispose(); $bmp.Dispose()
        Remove-Item $PngPath -Force
    }

    Save-Icon "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"   "$iconDir\start.png"   "$iconDir\start.ico"
    Save-Icon "/9j/4AAQSkZJRgABAQAAAQABAAD/4gHYSUNDX1BST0ZJTEUAAQEAAAHIAAAAAAQwAABtbnRyUkdCIFhZWiAH4AABAAEAAAAAAABhY3NwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAA9tYAAQAAAADTLQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAlkZXNjAAAA8AAAACRyWFlaAAABFAAAABRnWFlaAAABKAAAABRiWFlaAAABPAAAABR3dHB0AAABUAAAABRyVFJDAAABZAAAAChnVFJDAAABZAAAAChiVFJDAAABZAAAAChjcHJ0AAABjAAAADxtbHVjAAAAAAAAAAEAAAAMZW5VUwAAAAgAAAAcAHMAUgBHAEJYWVogAAAAAAAAb6IAADj1AAADkFhZWiAAAAAAAABimQAAt4UAABjaWFlaIAAAAAAAACSgAAAPhAAAts9YWVogAAAAAAAA9tYAAQAAAADTLXBhcmEAAAAAAAQAAAACZmYAAPKnAAANWQAAE9AAAApbAAAAAAAAAABtbHVjAAAAAAAAAAEAAAAMZW5VUwAAACAAAAAcAEcAbwBvAGcAbABlACAASQBuAGMALgAgADIAMAAxADb/2wBDAAUDBAQEAwUEBAQFBQUGBwwIBwcHBw8LCwkMEQ8SEhEPERETFhwXExQaFRERGCEYGh0dHx8fExciJCIeJBweHx7/2wBDAQUFBQcGBw4ICA4eFBEUHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh7/wAARCACPAHMDASIAAhEBAxEB/8QAHAAAAgIDAQEAAAAAAAAAAAAAAAgFBwEDBgIE/8QAQhAAAQMCAgUIBwYEBgMAAAAAAQACAwQFBhEHCBIhMRM1QVFzdJOxFBYiMjdVYRUjOEJxshczNlQlRVaBodFSU5H/xAAbAQEAAgMBAQAAAAAAAAAAAAAAAgMBBQYEB//EADIRAAIBAwIEBAUBCQAAAAAAAAABAgMEEQUhBhIxQRMUUWEVInGx0YEHIzJCkaHB4fD/2gAMAwEAAhEDEQA/AGZstDV3WOerku9bHnM4NY1+QaAvv9X5/ndf4izgrmqXt3+a1YixhZLFcKWhrqprZqh2y0Z8P1WJ1IwWZPCJ0LepXly04tv2Nnq/P86uHiI9X5/nVw8RTcEjJY2yRuDmOGYIWxSyV4x1ID1fn+dV/iI9X5/nVw8RTyEyMED6vz/Orh4iPV+f51cPEU8hMjBA+r8/zq4eIj1fn+dXDxFPITIwQPq/P86uHiLycP1A/wA6uHiLoEEZpkYOe+waj51cPER9g1GXPVw8RTzhkV5PBMjBV1dfbzRVktIy4SvbE4tDnbyUKPxBz1V9oULOCrJZGDQTZ5mg5Ezv3pdNPOGL7a8Rvu1RPLU0srs2Sf8Ar+iY3BPNcvbv81A4mvWH7piSowReWxtkmhD4S/8ANn1fVa/UbNXdJwzh9jpOG9blo90q3LmL2f09ivNBOlDMx2C9zb/dhlceP0V+sc17Q5hBaRmCEnmk3BFfgu9GSIPNG521DK38qtbQZpPFdFHYb1MBUNGUUjj7y1mm6hOnPy1xtJdGdZxPw7RuqPxXTd4S3kl29/yi7ULDXBwzHArK6E+aghCEAIQhACEIQARmtbgQFsWHe6UBTOID/jVX2hQsYi57q+0KFMpLNwTzXL27/NLDrX1lTb9KVLWUkzoZ4oQ5j2neCmewTzXL27/NKzrgfEWLu4UGWx6FkaLcb2fSfhh2G8QiMXJkeyC7827iFVuPcJ3TBF/y+8EW3tQTNVVWW51tnuUNwoJnQ1ELg5rmlOrh2kg0maK6KovMLPSJoQQ8DeHZcVq9T06N1DmjtJdGdbwvxLU0ityVN6Uuq/yiL0IaTIr1Sx2e7ShlbGMmOcffCt8HPgktxbh+8YGxJsO24yx+1BMNwITA6FtJFPiShZbrhIGXCIZbz7/1Xl0zUZN+Xr7SX9zbcVcNU4w+JafvSlu0u3v9PsWivku1wpbZQyVlXK2OKNubiStldVwUVK+pqJGxxsGbnEpX9NWkefEda+12yRzLfG7IkH+YVsL++haU+Z9eyOc4f0CvrNwqcNorq/RfknrhptqG45bJA3O0MPJub1jPir5w/eKK92yKvoZWyRyNz3HglEpNHmIqnC7r/HTHkW7wzL2iOtSmiLSBWYRuraWqe99vkdsvYfyHrWjstUrUqmLlbS6He63wlY3dq5aW050tml3x6+426F8louNLdaCKto5WyRSNBBBX1rqU01lHyWcJQk4yWGgQ73T+iEO90/ossiUxiHnur7QoRiHnur7QoUyks3BPNUvbv80rOuD8RYe7hNNgnmuXt3+aVnXB+IsPdwoFq6FJFOHgm4VNq1cY7jSPLJ4KbbYfqAk8Kbixfhdk7k79qGT6sOXjD+mXBZpaoRx3eFmTh+ZrusfRUperbe8C4n2HcpDPC/OOQbg4Lg8HYkuWFr5DdLZO6OSN2bmg7nDqKaqjnw/pnwUJYjHFdoWe0PzNd/0tPqem+YXiU9po7ThXid6bPy9x81GXVenv+Sssb6VbxiSyQWzI07Q3KZzT76ktDGjl96nbe7wzkrdCdoB+7byXrR7omuFXiaYXuIw0VFJ7RPB+S26e9KVPa6M4NwnI2NrG7E8sZ4DqC19hYVrqp49126I6PiHiGz0q38jpOE5btrtn39fsX3W3ex2vDLa37v7MblGXN90DgqN00aOmMacTYdaJaOYbcjGbwM+kKfxA4nVeDiSSaMEn/ZcHq/aV20zWYTxPKJaKUbEMshz2c+greXllTuqfJL9DgtE1y40m6Vem8p9V6o3aGNI1Rhm4Mttxkc+3yOy9o/yymjt9ZT11JHVU0jZIpG5tIKWjTPo4NqkdfbI3laCb2nBm/Zz6VnQlpLmsdXHZrtM51FIcmPcfcK0tjeVLKr5a46dmd1r+i2+u2vxXTf4v5o/93+4zqHe6f0XO37F1ptNJSzSVDHmqe1kTWne7MqfY4PhDxwc3NdGpxk2kz5hOhUpxU5Rwn0/QpvEPPdX2hQjEPPdX2hQrjyFm4I5rl7d/mlZ1wfiLD3cJpsEc1S9u/wA0rOuD8RYe7hQLV0KSKbix79VyTuTv2pRym4sf4XJO5O/ahkUg8VcGqdPMzSdHCyR4jfC7aaDuO5U+eKtvVR+KkHYu8kBa2srpRkw7G/DVmBjrZ2/eyjdstKVCWSSaZ0sry97zm5xO8lW1rX/FGbsmqowgG3v/AOFsdyHklIYS0hzSQRvBCbe//hbHch5JRxwQDPas2OarE1NNg29x+lsZCSyR+/2eorntNmA/VS6enUbh6FUOza3Pe0qO1QM/4iT7PH0V2SszSto8xrii5uqvS4ZYGn7qEbgAtRrFt41H5Y5l2Oz4J1LyV9+8qqFN9U+/+yjqe9XCSvoXVVXLLHTSNLGuduaM062HpxU2GkmBz2oWn/hKPc9GeMbedp9qkkA6Y96ZzRVLUSYIoW1Ub4pmR7L2vGRBC8GhKtTnONVNfU6D9oE7S4tqNW2lFpNrbHfc4TEPPdX2hQjEPPdX2hQupPkxZuCeape3f5pWdcH4iw93CabBPNUvbv8ANKzrg/EWHu4UC1dCkim4sX4XZO5O/alHKbix79V2TuTv2oZFIPFW3qpfFSDsXeSqQ8Vbeql8VIOxd5ID1rX/ABSm7JqqIK3da/4pTdk1VEEA29//AAtjuQ8ko44JuL/+FsdyHklHHBAXbqffEabuzk4CT/U++I03dnJwEBhzWkZEAj6ry2NjGnYaAPoF7Q73T+iwMlMYh57q+0KEYh57q+0KFYUlm4JIFrl7d/mlx1psMX+9Y9jqLXa6mqibAAXxsJGaYxmFaOOWR8VTVx8o4vLWykDM/RYdhSjcczV1hPbFR2LFlIRc6P8AGX+n63wymds1puMerjJbH0craz0RzORLfazy4ZLsbpbqGCu+zqJ9bV1hG0WNmIDB1k9Cj30ldRXClpbjS1rYKmQRtkjqCQ0/XepKDZaqc2KCcAYxz/p+t8MqztWnCuIbRpJgqrlaammg5Jw23sIGeSYu+WKittqmrRLWymMZhglObj1LmbhO+2voY6q3XFkla8MiAquk9e9IwcughSnPoinNZfCuIbxpImqbbaqmph5MDbYwkZqsP4fYx/0/W+GU49BbJZb22grqetp9the1/pBcDl0cUYyt7LBbn3BsddUwRDOTZnOYH/1FB5wFTk5KPqcnerVcZNXBtqZRyurfRAzkQ32s8uGSWMaPsZZf0/W+GU6tis1Pc7FDcBPVsE0e2xnLk7ujpXy4ctQuss3KR1kMUTyzbM5O0R1b1jle5jklv7FI6rOGL/ZcfS1F0tdTSxOp3APkYQM01GY61BDClGDmKusB7YrPqtTf3td4zv8AtRIZZOZjrWHEbJ3jgoQYWpv72t8ZyPVal/va7xnIMlaYh57q+0KFY4wXZcvbie93S5zySUKWUV8rOkWDwKysFRLSv9GlQ6pxTiKSc5zio2RnxDRwXezxRShvKtadk5jPoK4K9W6os2LTdrC5ss9SAKilJ3P+v0UlU1l7r6ykgqaL0GnMgdI7azLvovRUjzNSXQ9tan4jU4vbB1c8Mc8RjlaHNPQVwuk1g+28OZdFYPIrvQMgAq/0jMudRerU6npA5lPUB4Jd7/0UaG8yu0Wan9fsd6YmOkbIWguaNx6lovFJHX2uopJW7TZYy0j/AGXuglmlpmPni5KQj2mZ55L6FV0Z58uL+hWuBLrLRYVrLNI7aq6Oc07G9O87v+F3tlo20NshpwPaDc3HrPSVx1NYom6T56wSZQujEnJjgZOtd8razWdu+56rqUW8x77ghCFSeQEIQgBCEIAWCsoQEJZKfO6Vc8ozftZAnoCmZYo5GhrwCAcwtXIbMxljOTjxHWsvbK9zc3ANBz3dKk3lk5S5nk28Aoi/R7dVRH/xkBUrM0vjLQcj1r556eSV0ZcW5sOYSLw8iDw8n1LEjgyNzjwAzXhrZDKHPIyA4BYqonzRGNpAB4qJHBB0NG6V01yyPKl+bD9FPwv5SJrusLFPFyUAi3bhkvNPG6MFpIIzzClJ5JSlzG5CEKJAEIQgBCEID//Z" "$iconDir\restart.png" "$iconDir\restart.ico"
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

    $startPs1 = Get-ChildItem "C:\CTV" -Filter "START_TRADING.ps1" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($startPs1) {
        Set-Shortcut "$desktop\Start Trading.lnk" "powershell.exe" `
            "-ExecutionPolicy Bypass -File `"$($startPs1.FullName)`"" $startPs1.DirectoryName "$iconDir\start.ico,0"
    }

    $editor = Get-ChildItem "C:\CTV" -Filter "ctv2_config_editor.jsx" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
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
