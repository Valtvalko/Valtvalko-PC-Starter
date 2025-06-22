# Requires Administrator privileges for winget, pip, and some installations.
# If you encounter an error about execution policy, run this in PowerShell:
# Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force

# --- Self-Elevation Check ---
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit }

# --- Functions ---

function Show-MainMenu {
    Clear-Host
    Write-Host ""
    Write-Host "========================================="
    Write-Host "  WINGET & PROGRAM INSTALLER MENU"
    Write-Host "========================================="
    Write-Host ""
    Write-Host "Please select an option:"
    Write-Host ""

    # --- IMPORTANT NOTE ---
    Write-Host "NOTE: It's highly recommended to install the "Redistributables" [3] first." -ForegroundColor Cyan
    Write-Host "      This ensures essential components for other programs are present." -ForegroundColor Cyan
    Write-Host ""
    # --- END IMPORTANT NOTE ---

    if (-not $personalProgramsInstalled) { Write-Host "[1] Install Personal Programs" } else { Write-Host "[1] Personal Programs (Already Installed)" }
    if (-not $fluentProgramsInstalled) { Write-Host "[2] Install Fluent Programs" } else { Write-Host "[2] Fluent Programs (Already Installed)" }
    if (-not $redistInstalled) { Write-Host "[3] Install Redistributables" } else { Write-Host "[3] Redistributables (Already Installed)" }
    Write-Host "[4] Other Downloads"
    Write-Host "[5] Exit Installer"
    Write-Host ""

    if ($personalProgramsInstalled -and $fluentProgramsInstalled -and $redistInstalled) {
        Write-Host "All core program sets have been installed."
        Write-Host "Please select [5] to exit or [4] for Other Downloads."
    }
}

function Show-OtherMenu {
    param (
        [ref]$backToMainMenu # Flag to signal return to main menu
    )
    $otherExit = $false
    while (-not $otherExit) {
        Clear-Host
        Write-Host ""
        Write-Host "========================================="
        Write-Host "       OTHER DOWNLOADS MENU"
        Write-Host "========================================="
        Write-Host ""
        Write-Host "Select an item to install/download:"
        Write-Host ""

        # --- Tools & Utilities Section ---
        Write-Host "--- Tools & Utilities ---"
        if (-not $iconViewerDownloaded) { Write-Host "[1] Download IconViewer.exe" } else { Write-Host "[1] IconViewer.exe (Already Downloaded)" }
        if (-not $fileCRTxtDownloaded) { Write-Host "[2] Download FileCR.txt" } else { Write-Host "[2] FileCR.txt (Already Downloaded)" }
        if (-not $spotdlInstalled) { Write-Host "[3] Install SpotDL" } else { Write-Host "[3] SpotDL (Already Installed)" }
        if (-not $launcherXDownloaded) { Write-Host "[4] Download LauncherX_2.1.2_x64_Setup.exe" } else { Write-Host "[4] LauncherX_2.1.2_x64_Setup.exe (Already Downloaded)" }
        # --- NEW OPTION ---
        if (-not $winUtilDownloaded) { Write-Host "[5] Download WinUtil.ps1" } else { Write-Host "[5] WinUtil.ps1 (Already Downloaded)" }
        
        # --- Separator with empty lines ---
        Write-Host ""
        Write-Host "---------------------------------"
        Write-Host ""
        
        # --- Content & Customization Section ---
        Write-Host "--- Content & Customization ---"
        if (-not $personalRainmeterDownloaded) { Write-Host "[6] Download Personal.Rainmeter.zip" } else { Write-Host "[6] Personal.Rainmeter.zip (Already Downloaded)" } # Option number changed
        if (-not $crossWallpapersDownloaded) { Write-Host "[7] Download Cross Wallpapers.zip" } else { Write-Host "[7] Cross Wallpapers.zip (Already Downloaded)" } # Option number changed
        if (-not $spicetifyInstalled) { Write-Host "[8] Install Spicetify" } else { Write-Host "[8] Spicetify (Already Installed)" } # Option number changed
        if (-not $rectify11Downloaded) { Write-Host "[9] Download Rectify11Installer (x64).zip" } else { Write-Host "[9] Rectify11Installer (x64).zip (Already Downloaded)" } # Option number changed
        if (-not $systemAppIconsDownloaded) { Write-Host "[10] Download System.App.Icons.zip" } else { Write-Host "[10] System.App.Icons.zip (Already Downloaded)" } # Option number changed


        Write-Host ""
        Write-Host "[11] Back to Main Menu" # Option number changed
        Write-Host ""

        $otherChoice = Read-Host "Enter your choice (1-11)" # Range changed

        switch ($otherChoice) {
            "1" { # IconViewer
                if (-not $iconViewerDownloaded) {
                    Download-IconViewer
                } else {
                    Write-Host "IconViewer.exe already seems to be downloaded." -ForegroundColor Yellow
                    Read-Host "Press Enter to continue..."
                }
            }
            "2" { # FileCR.txt
                if (-not $fileCRTxtDownloaded) {
                    Download-FileCRTxt
                } else {
                    Write-Host "FileCR.txt already seems to be downloaded." -ForegroundColor Yellow
                    Read-Host "Press Enter to continue..."
                }
            }
            "3" { # SpotDL
                if (-not $spotdlInstalled) {
                    Install-SpotDL
                } else {
                    Write-Host "SpotDL already installed." -ForegroundColor Yellow
                    Read-Host "Press Enter to continue..."
                }
            }
            "4" { # LauncherX
                if (-not $launcherXDownloaded) {
                    Download-LauncherX
                } else {
                    Write-Host "LauncherX_2.1.2_x64_Setup.exe already seems to be downloaded." -ForegroundColor Yellow
                    Read-Host "Press Enter to continue..."
                }
            }
            "5" { # NEW: WinUtil.ps1
                if (-not $winUtilDownloaded) {
                    Download-WinUtil
                } else {
                    Write-Host "WinUtil.ps1 already seems to be downloaded." -ForegroundColor Yellow
                    Read-Host "Press Enter to continue..."
                }
            }
            "6" { # Personal Rainmeter (Option number changed)
                if (-not $personalRainmeterDownloaded) {
                    Download-PersonalRainmeter
                } else {
                    Write-Host "Personal.Rainmeter.zip already seems to be downloaded." -ForegroundColor Yellow
                    Read-Host "Press Enter to continue..."
                }
            }
            "7" { # Cross Wallpapers (Option number changed)
                if (-not $crossWallpapersDownloaded) {
                    Download-CrossWallpapers
                } else {
                    Write-Host "Cross Wallpapers.zip already seems to be downloaded." -ForegroundColor Yellow
                    Read-Host "Press Enter to continue..."
                }
            }
            "8" { # Spicetify (Option number changed)
                if (-not $spicetifyInstalled) {
                    Install-Spicetify
                } else {
                    Write-Host "Spicetify already installed." -ForegroundColor Yellow
                    Read-Host "Press Enter to continue..."
                }
            }
            "9" { # Rectify11 (Option number changed)
                if (-not $rectify11Downloaded) {
                    Download-Rectify11
                } else {
                    Write-Host "Rectify11Installer (x64).zip already seems to be downloaded." -ForegroundColor Yellow
                    Read-Host "Press Enter to continue..."
                }
            }
            "10" { # System.App.Icons.zip (Option number changed)
                if (-not $systemAppIconsDownloaded) {
                    Download-SystemAppIcons
                } else {
                    Write-Host "System.App.Icons.zip already seems to be downloaded." -ForegroundColor Yellow
                    Read-Host "Press Enter to continue..."
                }
            }
            "11" { # Back to Main Menu (Option number changed)
                $otherExit = $true
                $backToMainMenu.Value = $true
            }
            default {
                Write-Host "Invalid choice. Please enter 1-11." -ForegroundColor Red
                Read-Host "Press Enter to continue..."
            }
        }
    }
}

# --- Specific Download/Installation Functions ---

function Install-Spicetify {
    Clear-Host
    Write-Host ""
    Write-Host "========================================="
    Write-Host "       INSTALLING: Spicetify"
    Write-Host "========================================="
    Write-Host ""
    Write-Host "Running Spicetify installation scripts... This may take a moment."
    try {
        Invoke-Expression "iwr -useb https://raw.githubusercontent.com/spicetify/cli/main/install.ps1 | iex ; iwr -useb https://raw.githubusercontent.com/spicetify/marketplace/main/resources/install.ps1 | iex"
        Write-Host ""
        Write-Host "Spicetify installation commands executed. Manual steps might be needed." -ForegroundColor Yellow
        Write-Host "(You might need to run 'spicetify backup' and 'spicetify apply' manually.)" -ForegroundColor Yellow
        Write-Host ""
        $script:spicetifyInstalled = $true
    } catch {
        Write-Host ""
        Write-Host "Failed to install Spicetify: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "(Please check PowerShell execution policy.)" -ForegroundColor Red
        Write-Host "(You might need to run: Set-ExecutionPolicy RemoteSigned -Scope CurrentUser)" -ForegroundColor Red
        Write-Host ""
    }
    Read-Host "Press Enter to return to the Other Downloads menu..."
}

function Install-SpotDL {
    Clear-Host
    Write-Host ""
    Write-Host "========================================="
    Write-Host "        INSTALLING: SpotDL"
    Write-Host "========================================="
    Write-Host ""
    Write-Host "Installing SpotDL via pip... Requires Python to be installed."
    Write-Host "Executing command: py -m pip install spotdl" -ForegroundColor Cyan
    try {
        py -m pip install spotdl
        if ($LASTEXITCODE -ne 0) {
             Write-Host "Pip command returned a non-zero exit code. Check Python installation and PATH." -ForegroundColor Red
        } else {
            $script:spotdlInstalled = $true
        }
    } catch {
        Write-Host ""
        Write-Host "Failed to execute pip command: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "(Ensure Python is installed and added to PATH.)" -ForegroundColor Red
        Write-Host ""
    }
    Read-Host "Press Enter to return to the Other Downloads menu..."
}

function Download-CrossWallpapers {
    Clear-Host
    Write-Host ""
    Write-Host "========================================="
    Write-Host "   DOWNLOADING: Cross Wallpapers.zip"
    Write-Host "========================================="
    Write-Host ""

    $url = "https://github.com/Valtvalko/Valt-s-Starter-PC/raw/refs/heads/main/Cross%20Wallpapers.zip"
    $desktopPath = [Environment]::GetFolderPath([Environment+SpecialFolder]::Desktop)
    $outputFileName = "Cross Wallpapers.zip"
    $outputPath = Join-Path -Path $desktopPath -ChildPath $outputFileName

    Write-Host "Downloading from: $url"
    Write-Host "Saving to: $outputPath"
    Write-Host ""

    try {
        Invoke-WebRequest -Uri $url -OutFile $outputPath -ErrorAction Stop
        Write-Host "Download completed successfully!" -ForegroundColor Green
        $script:crossWallpapersDownloaded = $true # Set global flag for this download
    }
    catch {
        Write-Host "Error downloading the file:" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        Write-Host "Please ensure you have an active internet connection and write permissions to your Desktop." -ForegroundColor Yellow
    }

    Read-Host "Press Enter to return to the Other Downloads menu..."
}

function Download-IconViewer {
    Clear-Host
    Write-Host ""
    Write-Host "========================================="
    Write-Host "    DOWNLOADING: IconViewer.exe"
    Write-Host "========================================="
    Write-Host ""

    $url = "https://www.botproductions.com/iconview/download/IconViewer3.02-Setup-x64.exe"
    $desktopPath = [Environment]::GetFolderPath([Environment+SpecialFolder]::Desktop)
    $outputFileName = "IconViewer3.02-Setup-x64.exe"
    $outputPath = Join-Path -Path $desktopPath -ChildPath $outputFileName

    Write-Host "Downloading from: $url"
    Write-Host "Saving to: $outputPath"
    Write-Host ""

    try {
        Invoke-WebRequest -Uri $url -OutFile $outputPath -ErrorAction Stop
        Write-Host "Download completed successfully!" -ForegroundColor Green
        $script:iconViewerDownloaded = $true
    }
    catch {
        Write-Host "Error downloading the file:" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        Write-Host "Please ensure you have an active internet connection and write permissions to your Desktop." -ForegroundColor Yellow
    }

    Read-Host "Press Enter to return to the Other Downloads menu..."
}

function Download-PersonalRainmeter {
    Clear-Host
    Write-Host ""
    Write-Host "========================================="
    Write-Host " DOWNLOADING: Personal.Rainmeter.zip"
    Write-Host "========================================="
    Write-Host ""

    $url = "https://github.com/Valtvalko/Valt-s-Starter-PC/releases/download/Rainmeter/Personal.Rainmeter.zip"
    $desktopPath = [Environment]::GetFolderPath([Environment+SpecialFolder]::Desktop)
    $outputFileName = "Personal.Rainmeter.zip"
    $outputPath = Join-Path -Path $desktopPath -ChildPath $outputFileName

    Write-Host "Downloading from: $url"
    Write-Host "Saving to: $outputPath"
    Write-Host ""

    try {
        Invoke-WebRequest -Uri $url -OutFile $outputPath -ErrorAction Stop
        Write-Host "Download completed successfully!" -ForegroundColor Green
        $script:personalRainmeterDownloaded = $true
    }
    catch {
        Write-Host "Error downloading the file:" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        Write-Host "Please ensure you have an active internet connection and write permissions to your Desktop." -ForegroundColor Yellow
    }

    Read-Host "Press Enter to return to the Other Downloads menu..."
}

function Download-FileCRTxt {
    Clear-Host
    Write-Host ""
    Write-Host "========================================="
    Write-Host "     DOWNLOADING: FileCR.txt"
    Write-Host "========================================="
    Write-Host ""

    $url = "https://raw.githubusercontent.com/Valtvalko/Valt-s-Starter-PC/refs/heads/main/FileCR.txt"
    $desktopPath = [Environment]::GetFolderPath([Environment+SpecialFolder]::Desktop)
    $outputFileName = "FileCR.txt"
    $outputPath = Join-Path -Path $desktopPath -ChildPath $outputFileName

    Write-Host "Downloading from: $url"
    Write-Host "Saving to: $outputPath"
    Write-Host ""

    try {
        Invoke-WebRequest -Uri $url -OutFile $outputPath -ErrorAction Stop
        Write-Host "Download completed successfully!" -ForegroundColor Green
        $script:fileCRTxtDownloaded = $true
    }
    catch {
        Write-Host "Error downloading the file:" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        Write-Host "Please ensure you have an active internet connection and write permissions to your Desktop." -ForegroundColor Yellow
    }

    Read-Host "Press Enter to return to the Other Downloads menu..."
}

function Download-Rectify11 {
    Clear-Host
    Write-Host ""
    Write-Host "========================================="
    Write-Host "  DOWNLOADING: Rectify11Installer (x64).zip"
    Write-Host "========================================="
    Write-Host ""

    $url = "https://nightly.link/Rectify11/Installer/workflows/build/legacy-installer/Rectify11Installer%20%28x64%29.zip"
    $desktopPath = [Environment]::GetFolderPath([Environment+SpecialFolder]::Desktop)
    $outputFileName = "Rectify11Installer (x64).zip"
    $outputPath = Join-Path -Path $desktopPath -ChildPath $outputFileName

    Write-Host "Downloading from: $url"
    Write-Host "Saving to: $outputPath"
    Write-Host ""

    try {
        Invoke-WebRequest -Uri $url -OutFile $outputPath -ErrorAction Stop
        Write-Host "Download completed successfully!" -ForegroundColor Green
        $script:rectify11Downloaded = $true
    }
    catch {
        Write-Host "Error downloading the file:" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        Write-Host "Please ensure you have an active internet connection and write permissions to your Desktop." -ForegroundColor Yellow
    }

    Read-Host "Press Enter to return to the Other Downloads menu..."
}

function Download-SystemAppIcons {
    Clear-Host
    Write-Host ""
    Write-Host "========================================="
    Write-Host "  DOWNLOADING: System.App.Icons.zip"
    Write-Host "========================================="
    Write-Host ""

    $url = "https://github.com/Valtvalko/Valtvalko-PC-Starter/releases/download/IconPack/System.App.Icons.zip"
    $desktopPath = [Environment]::GetFolderPath([Environment+SpecialFolder]::Desktop)
    $outputFileName = "System.App.Icons.zip"
    $outputPath = Join-Path -Path $desktopPath -ChildPath $outputFileName

    Write-Host "Downloading from: $url"
    Write-Host "Saving to: $outputPath"
    Write-Host ""

    try {
        Invoke-WebRequest -Uri $url -OutFile $outputPath -ErrorAction Stop
        Write-Host "Download completed successfully!" -ForegroundColor Green
        $script:systemAppIconsDownloaded = $true
    }
    catch {
        Write-Host "Error downloading the file:" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        Write-Host "Please ensure you have an active internet connection and write permissions to your Desktop." -ForegroundColor Yellow
    }

    Read-Host "Press Enter to return to the Other Downloads menu..."
}

function Download-LauncherX {
    Clear-Host
    Write-Host ""
    Write-Host "========================================="
    Write-Host "  DOWNLOADING: LauncherX_2.1.2_x64_Setup.exe"
    Write-Host "========================================="
    Write-Host ""

    $url = "https://github.com/Apollo199999999/LauncherX/releases/download/v2.1.2/LauncherX_2.1.2_x64_Setup.exe"
    $desktopPath = [Environment]::GetFolderPath([Environment+SpecialFolder]::Desktop)
    $outputFileName = "LauncherX_2.1.2_x64_Setup.exe"
    $outputPath = Join-Path -Path $desktopPath -ChildPath $outputFileName

    Write-Host "Downloading from: $url"
    Write-Host "Saving to: $outputPath"
    Write-Host ""

    try {
        Invoke-WebRequest -Uri $url -OutFile $outputPath -ErrorAction Stop
        Write-Host "Download completed successfully!" -ForegroundColor Green
        $script:launcherXDownloaded = $true
    }
    catch {
        Write-Host "Error downloading the file:" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        Write-Host "Please ensure you have an active internet connection and write permissions to your Desktop." -ForegroundColor Yellow
    }

    Read-Host "Press Enter to return to the Other Downloads menu..."
}

function Download-WinUtil {
    Clear-Host
    Write-Host ""
    Write-Host "========================================="
    Write-Host "  DOWNLOADING: WinUtil.ps1"
    Write-Host "========================================="
    Write-Host ""

    $url = "https://github.com/ChrisTitusTech/winutil/releases/download/25.05.23/winutil.ps1"
    $desktopPath = [Environment]::GetFolderPath([Environment+SpecialFolder]::Desktop)
    $outputFileName = "winutil.ps1"
    $outputPath = Join-Path -Path $desktopPath -ChildPath $outputFileName

    Write-Host "Downloading from: $url"
    Write-Host "Saving to: $outputPath"
    Write-Host ""

    try {
        Invoke-WebRequest -Uri $url -OutFile $outputPath -ErrorAction Stop
        Write-Host "Download completed successfully!" -ForegroundColor Green
        $script:winUtilDownloaded = $true
        Write-Host "To run WinUtil, open PowerShell as Administrator and navigate to your Desktop."
        Write-Host "Then execute: '. $outputPath'" -ForegroundColor Cyan
    }
    catch {
        Write-Host "Error downloading the file:" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        Write-Host "Please ensure you have an active internet connection and write permissions to your Desktop." -ForegroundColor Yellow
    }

    Read-Host "Press Enter to return to the Other Downloads menu..."
}


function Install-ProgramSet {
    param (
        [string]$SetName,
        [scriptblock]$InstallCommands,
        [ref]$InstalledFlag # Use [ref] to modify the variable in the parent scope
    )

    Clear-Host
    Write-Host ""
    Write-Host "========================================="
    Write-Host "  INSTALLING: $SetName"
    Write-Host "========================================="
    Write-Host ""

    # Execute the installation commands
    Invoke-Command -ScriptBlock $InstallCommands

    Write-Host ""
    Write-Host "$SetName installation finished!"
    $InstalledFlag.Value = $true # Set the flag in the parent scope
    Read-Host "Press Enter to continue to the next step..."
}

# --- Main Script Logic ---

# Initialize installation/download flags
$personalProgramsInstalled = $false
$fluentProgramsInstalled = $false
$redistInstalled = $false
$spicetifyInstalled = $false 
$spotdlInstalled = $false     
$crossWallpapersDownloaded = $false
$iconViewerDownloaded = $false    
$personalRainmeterDownloaded = $false 
$fileCRTxtDownloaded = $false     
$rectify11Downloaded = $false 
$systemAppIconsDownloaded = $false
$launcherXDownloaded = $false 
$winUtilDownloaded = $false # NEW: Flag for WinUtil
$exitInstaller = $false

# Check if files already exist on desktop/relevant paths at startup to pre-set flags
$desktopPath = [Environment]::GetFolderPath([Environment+SpecialFolder]::Desktop)

# Check for downloads on Desktop
if (Test-Path (Join-Path -Path $desktopPath -ChildPath "Cross Wallpapers.zip")) { $crossWallpapersDownloaded = $true }
if (Test-Path (Join-Path -Path $desktopPath -ChildPath "IconViewer3.02-Setup-x64.exe")) { $iconViewerDownloaded = $true }
if (Test-Path (Join-Path -Path $desktopPath -ChildPath "Personal.Rainmeter.zip")) { $personalRainmeterDownloaded = $true }
if (Test-Path (Join-Path -Path $desktopPath -ChildPath "FileCR.txt")) { $fileCRTxtDownloaded = $true }
if (Test-Path (Join-Path -Path $desktopPath -ChildPath "Rectify11Installer (x64).zip")) { $rectify11Downloaded = $true } 
if (Test-Path (Join-Path -Path $desktopPath -ChildPath "System.App.Icons.zip")) { $systemAppIconsDownloaded = $true }
if (Test-Path (Join-Path -Path $desktopPath -ChildPath "LauncherX_2.1.2_x64_Setup.exe")) { $launcherXDownloaded = $true }
# NEW: Check for WinUtil
if (Test-Path (Join-Path -Path $desktopPath -ChildPath "winutil.ps1")) { $winUtilDownloaded = $true }


while (-not $exitInstaller) {
    Show-MainMenu
    $choice = Read-Host "Enter your choice (1-5): " 

    switch ($choice) {
        "1" {
            if (-not $personalProgramsInstalled) {
                Install-ProgramSet "Personal Programs" {
                    Write-Host "Installing BrianApps Sizer..."
                    winget install --id=BrianApps.Sizer -e --accept-package-agreements --accept-source-agreements
                    if ($LASTEXITCODE -ne 0) { Write-Host "Failed to install BrianApps Sizer." -ForegroundColor Red }

                    Write-Host "Installing Proton Pass..."
                    winget install --id=Proton.ProtonPass -e --accept-package-agreements --accept-source-agreements
                    if ($LASTEXITCODE -ne 0) { Write-Host "Failed to install Proton Pass." -ForegroundColor Red }

                    Write-Host "Installing OBS Studio..."
                    winget install --id=OBSProject.OBSStudio -e --accept-package-agreements --accept-source-agreements
                    if ($LASTEXITCODE -ne 0) { Write-Host "Failed to install OBS Studio." -ForegroundColor Red }

                    Write-Host "Installing Obsidian..."
                    winget install --id=Obsidian.Obsidian -e --accept-package-agreements --accept-source-agreements
                    if ($LASTEXITCODE -ne 0) { Write-Host "Failed to install Obsidian." -ForegroundColor Red }

                    Write-Host "Installing Spotify..."
                    winget install --id=Spotify.Spotify -e --accept-package-agreements --accept-source-agreements
                    if ($LASTEXITCODE -ne 0) { Write-Host "Failed to install Spotify." -ForegroundColor Red }

                    Write-Host "Installing Bulk Crap Uninstaller..."
                    winget install --id=Klocman.BulkCrapUninstaller -e --accept-package-agreements --accept-source-agreements
                    if ($LASTEXITCODE -ne 0) { Write-Host "Failed to install Bulk Crap Uninstaller." -ForegroundColor Red }

                    Write-Host "Installing Google Chrome..."
                    winget install --id=Google.Chrome -e --accept-package-agreements --accept-source-agreements
                    if ($LASTEXITCODE -ne 0) { Write-Host "Failed to install Google Chrome." -ForegroundColor Red }
                } ([ref]$personalProgramsInstalled)
            } else {
                Write-Host "Personal Programs already installed."
                Read-Host "Press Enter to continue..."
            }
        }
        "2" {
            if (-not $fluentProgramsInstalled) {
                Install-ProgramSet "Fluent Programs" {
                    Write-Host "Installing PowerToys..."
                    winget install --id Microsoft.PowerToys --exact --accept-package-agreements --accept-source-agreements
                    if ($LASTEXITCODE -ne 0) { Write-Host "Failed to install PowerToys." -ForegroundColor Red }

                    Write-Host "Installing Flow Launcher..."
                    winget install --id=Flow-Launcher.Flow-Launcher --exact --accept-package-agreements --accept-source-agreements
                    if ($LASTEXITCODE -ne 0) { Write-Host "Failed to install Flow Launcher." -ForegroundColor Red }

                    Write-Host "Installing Nora Music Player..."
                    winget install --id=Sandakan.Nora --exact --accept-package-agreements --accept-source-agreements
                    if ($LASTEXITCODE -ne 0) { Write-Host "Failed to install Nora Music Player." -ForegroundColor Red }

                    Write-Host "Installing Files..."
                    winget install --id FilesCommunity.Files --exact --accept-package-agreements --accept-source-agreements
                    if ($LASTEXITCODE -ne 0) { Write-Host "Failed to install Files." -ForegroundColor Red }

                    Write-Host "Installing FxSound..."
                    winget install --id=FxSound.FxSound --exact --accept-package-agreements --accept-source-agreements
                    if ($LASTEXITCODE -ne 0) { Write-Host "Failed to install FxSound." -ForegroundColor Red }

                    Write-Host "Installing Nilesoft Shell..."
                    winget install --id Nilesoft.Shell --exact --accept-package-agreements --accept-source-agreements
                    if ($LASTEXITCODE -ne 0) { Write-Host "Failed to install Nilesoft Shell." -ForegroundColor Red }

                    Write-Host "Installing Twinkle Tray..."
                    winget install --id=xanderfrangos.twinkletray --exact --accept-package-agreements --accept-source-agreements
                    if ($LASTEXITCODE -ne 0) { Write-Host "Failed to install Twinkle Tray." -ForegroundColor Red }

                    Write-Host "Installing Lively Wallpaper..."
                    winget install --id rocksdanister.LivelyWallpaper --exact --accept-package-agreements --accept-source-agreements
                    if ($LASTEXITCODE -ne 0) { Write-Host "Failed to install Lively Wallpaper." -ForegroundColor Red }

                    Write-Host "Installing Quick Look..."
                    winget install --id=QL-Win.QuickLook --exact --accept-package-agreements --accept-source-agreements
                    if ($LASTEXITCODE -ne 0) { Write-Host "Failed to install Quick Look." -ForegroundColor Red }

                    Write-Host "Installing Everything..."
                    winget install --id=voidtools.Everything --exact --accept-package-agreements --accept-source-agreements
                    if ($LASTEXITCODE -ne 0) { Write-Host "Failed to install Everything." -ForegroundColor Red }

                    Write-Host "Installing FluentWeather..."
                    winget install 9pfd136m8457 --exact --accept-package-agreements --accept-source-agreements
                    if ($LASTEXITCODE -ne 0) { Write-Host "Failed to install FluentWeather." -ForegroundColor Red }

                    Write-Host "Installing Mica TM..."
                    winget install --id MicaForEveryone.MicaForEveryone --exact --accept-package-agreements --accept-source-agreements
                    if ($LASTEXITCODE -ne 0) { Write-Host "Failed to install Mica TM." -ForegroundColor Red }
                    
                    Write-Host "Installing ShareX..."
                    winget.exe install --id "9NCRCVJC50WL" -e --accept-package-agreements --accept-source-agreements
                    if ($LASTEXITCODE -ne 0) { Write-Host "Failed to install ShareX." -ForegroundColor Red }

                    # LocalSend via Chocolatey
                    Write-Host "Installing LocalSend via Chocolatey..."
                    choco.exe install localsend -y --no-progress
                    if ($LASTEXITCODE -ne 0) { Write-Host "Failed to install LocalSend. Ensure Chocolatey is installed and added to PATH." -ForegroundColor Red }
                } ([ref]$fluentProgramsInstalled)
            } else {
                Write-Host "Fluent Programs already installed."
                Read-Host "Press Enter to continue..."
            }
        }
        "3" {
            if (-not $redistInstalled) {
                Install-ProgramSet "Redistributables" {
                    # Chocolatey via Winget
                    Write-Host "Installing Chocolatey via Winget..."
                    winget.exe install --id "Chocolatey.Chocolatey" -e --accept-package-agreements --accept-source-agreements
                    if ($LASTEXITCODE -ne 0) { Write-Host "Failed to install Chocolatey. Manual installation may be required: 'Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))'." -ForegroundColor Red }


                    Write-Host "Installing VC Redist 2005 x86..."
                    winget install --id=Microsoft.VCRedist.2005.x86 -e --accept-package-agreements --accept-source-agreements
                    if ($LASTEXITCODE -ne 0) { Write-Host "Failed to install VC Redist 2005 x86." -ForegroundColor Red }

                    Write-Host "Installing VC Redist 2005 x64..."
                    winget install --id=Microsoft.VCRedist.2005.x64 -e --accept-package-agreements --accept-source-agreements
                    if ($LASTEXITCODE -ne 0) { Write-Host "Failed to install VC Redist 2005 x64." -ForegroundColor Red }

                    Write-Host "Installing VC Redist 2008 x86..."
                    winget install --id=Microsoft.VCRedist.2008.x86 -e --accept-package-agreements --accept-source-agreements
                    if ($LASTEXITCODE -ne 0) { Write-Host "Failed to install VC Redist 2008 x86." -ForegroundColor Red }

                    Write-Host "Installing VC Redist 2008 x64..."
                    winget install --id=Microsoft.VCRedist.2008.x64 -e --accept-package-agreements --accept-source-agreements
                    if ($LASTEXITCODE -ne 0) { Write-Host "Failed to install VC Redist 2008 x64." -ForegroundColor Red }

                    Write-Host "Installing VC Redist 2010 x86..."
                    winget install --id=Microsoft.VCRedist.2010.x86 -e --accept-package-agreements --accept-source-agreements
                    if ($LASTEXITCODE -ne 0) { Write-Host "Failed to install VC Redist 2010 x86." -ForegroundColor Red }

                    Write-Host "Installing VC Redist 2010 x64..."
                    winget install --id=Microsoft.VCRedist.2010.x64 -e --accept-package-agreements --accept-source-agreements
                    if ($LASTEXITCODE -ne 0) { Write-Host "Failed to install VC Redist 2010 x64." -ForegroundColor Red }

                    Write-Host "Installing VC Redist 2012 x86..."
                    winget install --id=Microsoft.VCRedist.2012.x86 -e --accept-package-agreements --accept-source-agreements
                    if ($LASTEXITCODE -ne 0) { Write-Host "Failed to install VC Redist 2012 x86." -ForegroundColor Red }

                    Write-Host "Installing VC Redist 2012 x64..."
                    winget install --id=Microsoft.VCRedist.2012.x64 -e --accept-package-agreements --accept-source-agreements
                    if ($LASTEXITCODE -ne 0) { Write-Host "Failed to install VC Redist 2012 x64." -ForegroundColor Red }

                    Write-Host "Installing VC Redist 2013 x86..."
                    winget install --id=Microsoft.VCRedist.2013.x86 -e --accept-package-agreements --accept-source-agreements
                    if ($LASTEXITCODE -ne 0) { Write-Host "Failed to install VC Redist 2013 x86." -ForegroundColor Red }

                    Write-Host "Installing VC Redist 2013 x64..."
                    winget install --id=Microsoft.VCRedist.2013.x64 -e --accept-package-agreements --accept-source-agreements
                    if ($LASTEXITCODE -ne 0) { Write-Host "Failed to install VC Redist 2013 x64." -ForegroundColor Red }

                    Write-Host "Installing VC Redist 2015+ x86..."
                    winget install --id=Microsoft.VCRedist.2015+.x86 -e --accept-package-agreements --accept-source-agreements
                    if ($LASTEXITCODE -ne 0) { Write-Host "Failed to install VC Redist 2015+ x86." -ForegroundColor Red }

                    Write-Host "Installing VC Redist 2015+ x64..."
                    winget install --id=Microsoft.VCRedist.2015+.x64 -e --accept-package-agreements --accept-source-agreements
                    if ($LASTEXITCODE -ne 0) { Write-Host "Failed to install VC Redist 2015+ x64." -ForegroundColor Red }

                    Write-Host "Installing Microsoft XNA Redist..."
                    winget install --id=Microsoft.XNARedist -e --accept-package-agreements --accept-source-agreements
                    if ($LASTEXITCODE -ne 0) { Write-Host "Failed to install Microsoft XNA Redist." -ForegroundColor Red }

                    Write-Host "Installing Microsoft DirectX..."
                    winget install --id=Microsoft.DirectX -e --accept-package-agreements --accept-source-agreements
                    if ($LASTEXITCODE -ne 0) { Write-Host "Failed to install Microsoft DirectX." -ForegroundColor Red }

                    Write-Host "Installing OpenAL..."
                    winget install --id=OpenAL.OpenAL -e --accept-package-agreements --accept-source-agreements
                    if ($LASTEXITCODE -ne 0) { Write-Host "Failed to install OpenAL." -ForegroundColor Red }

                    Write-Host "Installing Oracle Java Runtime Environment..."
                    winget install --id=Oracle.JavaRuntimeEnvironment -e --accept-package-agreements --accept-source-agreements
                    if ($LASTEXITCODE -ne 0) { Write-Host "Failed to install Oracle Java Runtime Environment." -ForegroundColor Red }

                    Write-Host "Installing Python 3.12..."
                    winget install --id=Python.Python.3.12 -e --accept-package-agreements --accept-source-agreements
                    if ($LASTEXITCODE -ne 0) { Write-Host "Failed to install Python 3.12." -ForegroundColor Red }

                    Write-Host "Installing FFmpeg..."
                    winget install --id=Gyan.FFmpeg -e --accept-package-agreements --accept-source-agreements
                    if ($LASTEXITCODE -ne 0) { Write-Host "Failed to install FFmpeg." -ForegroundColor Red }
                } ([ref]$redistInstalled)
            } else {
                Write-Host "Redistributables already installed."
                Read-Host "Press Enter to continue..."
            }
        }
        "4" { # Other Downloads
            $dummyVar = $false
            Show-OtherMenu -backToMainMenu ([ref]$dummyVar)
        }
        "5" { # Exit
            $exitInstaller = $true
        }
        default {
            Write-Host "Invalid choice. Please enter 1-5." -ForegroundColor Red
            Read-Host "Press Enter to continue..."
        }
    }
}

# --- Final Message ---
Clear-Host
Write-Host ""
Write-Host "========================================="
Write-Host "  Installer Finished."
Write-Host "========================================="
Write-Host ""
Write-Host "All selected installations are complete. You can now close this window."
Read-Host "Press Enter to exit."