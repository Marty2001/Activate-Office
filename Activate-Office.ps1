<#
.SYNOPSIS
    Activates Microsoft Office using the Ohook activation method. This script is a PowerShell conversion of the original batch file.
    Homepage: massgrave.dev
    Email: mas.help@outlook.com
.DESCRIPTION
    This script provides functionalities to install or uninstall Ohook activation.
    It performs necessary system checks, handles administrative elevation via a UAC prompt, and interacts with Windows licensing services.
.PARAMETER Ohook
    A switch to activate Office with Ohook in unattended (silent) mode.
.PARAMETER Ohook_Uninstall
    A switch to remove Ohook activation in unattended (silent) mode.
#>
[CmdletBinding()]
param(
    [switch]$Ohook,
    [switch]$Ohook_Uninstall
)

#============================================================================
# UAC ELEVATION PROMPT
# This is the section that handles the User Account Control (UAC) prompt.
#============================================================================

# Check if the script is running with Administrator privileges
$identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object System.Security.Principal.WindowsPrincipal($identity)
$isAdmin = $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    # If not running as admin, re-launch the script with elevated privileges
    Write-Warning "This script requires administrator rights. Requesting elevation..."
    try {
        # Get all parameters passed to the script to forward them
        $params = $PSBoundParameters.GetEnumerator() | ForEach-Object {
            if ($_.Value -is [bool] -and $_.Value) { "-$($_.Key)" }
            else { "-$($_.Key) `"$($_.Value)`"" }
        }
        $allArgs = $params -join " "

        # Check if running from irm|iex (no file path)
        if ([string]::IsNullOrEmpty($PSCommandPath)) {
            # Running from web (irm | iex), save to temp file first
            Write-Host "Detected web execution. Creating temporary script file..." -ForegroundColor Yellow
            
            # Get the current script content
            $scriptContent = $MyInvocation.MyCommand.ScriptBlock.ToString()
            
            # Create temp file path
            $tempScript = Join-Path $env:TEMP "Ohook-Activation-$(Get-Random).ps1"
            
            # Save script to temp file
            $scriptContent | Out-File -FilePath $tempScript -Encoding UTF8 -Force
            
            # Launch elevated with cmd /k to keep window open
            $psCommand = "powershell.exe -ExecutionPolicy Bypass -NoProfile -File `"$tempScript`" $allArgs"
            Start-Process cmd.exe -ArgumentList "/k", $psCommand -Verb RunAs
        }
        else {
            # Running from a file, use cmd /k to keep window open
            $psCommand = "powershell.exe -ExecutionPolicy Bypass -NoProfile -File `"$PSCommandPath`" $allArgs"
            Start-Process cmd.exe -ArgumentList "/k", $psCommand -Verb RunAs
        }
    }
    catch {
        Write-Error "Failed to elevate privileges. Error: $_"
        Read-Host "Press Enter to exit"
    }
    # Exit the current (non-elevated) script
    exit
}

# Banner to show script is running with admin rights
Write-Host "============================================================" -ForegroundColor Green
Write-Host "  Running with Administrator privileges" -ForegroundColor White
Write-Host "============================================================" -ForegroundColor Green
Write-Host ""

$Host.UI.RawUI.BackgroundColor = "White"
$Host.UI.RawUI.ForegroundColor = "Black"
Clear-Host

#============================================================================
# Initial Setup and Variables
#============================================================================

$masver = "3.7"
$mas = "https://massgrave.dev/"
$github = "https://github.com/massgravel/Microsoft-Activation-Scripts"

# Determine if running in unattended mode based on parameters
$unattended = $false
if ($Ohook.IsPresent -or $Ohook_Uninstall.IsPresent) {
    $unattended = $true
}

#============================================================================
# Helper Functions
#============================================================================

# Function to write colored text to the console
function Write-ColorText {
    param (
        [string]$Message,
        [string]$ForegroundColor = "Black",
        [string]$BackgroundColor = "White"
    )
    Write-Host $Message -ForegroundColor $ForegroundColor -BackgroundColor $BackgroundColor
}

# Function to check for running Office applications before proceeding
function Test-RunningOfficeApps {
    $officeProcesses = @(
        "msaccess", "excel", "groove", "lync", "onenote",
        "outlook", "powerpnt", "winproj", "mspub", "visio", "winword"
    )
    
    # Get running processes whose names are in the list
    $runningApps = Get-Process -ErrorAction SilentlyContinue | Where-Object { $officeProcesses -contains $_.ProcessName }
    
    if ($runningApps) {
        $appNames = ($runningApps | ForEach-Object { $_.ProcessName }) -join ", "
        Write-ColorText "Warning: Please close the following Office applications before proceeding: $appNames" "Yellow"
        return $false
    }
    return $true
}

# Function to detect Office installations
function Get-OfficeInstallation {
    $officeInstalls = @()
    
    # Check for Office 16.0 C2R
    $o16c2r_paths = @(
        "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun",
        "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Office\ClickToRun"
    )
    
    foreach ($path in $o16c2r_paths) {
        if (Test-Path $path) {
            $installPath = (Get-ItemProperty -Path $path -Name InstallPath -ErrorAction SilentlyContinue).InstallPath
            if ($installPath -and (Test-Path "$installPath\root\Licenses16\ProPlus*.xrm-ms")) {
                $version = (Get-ItemProperty -Path "$path\Configuration" -Name VersionToReport -ErrorAction SilentlyContinue).VersionToReport
                $platform = (Get-ItemProperty -Path "$path\Configuration" -Name Platform -ErrorAction SilentlyContinue).Platform
                
                $officeInstalls += [PSCustomObject]@{
                    Version = "16.0"
                    Type = "C2R"
                    Path = "$installPath\root"
                    VersionNumber = $version
                    Architecture = $platform
                    Registry = $path
                }
            }
        }
    }
    
    # Check for Office 15.0 C2R
    $o15c2r_paths = @(
        "HKLM:\SOFTWARE\Microsoft\Office\15.0\ClickToRun",
        "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Office\15.0\ClickToRun"
    )
    
    foreach ($path in $o15c2r_paths) {
        if (Test-Path $path) {
            $installPath = (Get-ItemProperty -Path $path -Name InstallPath -ErrorAction SilentlyContinue).InstallPath
            if ($installPath -and (Test-Path "$installPath\root\Licenses\ProPlus*.xrm-ms")) {
                $version = (Get-ItemProperty -Path "$path\Configuration" -Name VersionToReport -ErrorAction SilentlyContinue).VersionToReport
                $platform = (Get-ItemProperty -Path "$path\Configuration" -Name Platform -ErrorAction SilentlyContinue).Platform
                
                $officeInstalls += [PSCustomObject]@{
                    Version = "15.0"
                    Type = "C2R"
                    Path = "$installPath\root"
                    VersionNumber = $version
                    Architecture = $platform
                    Registry = $path
                }
            }
        }
    }
    
    # Check for Office 16.0 MSI
    $o16msi_paths = @(
        "HKLM:\SOFTWARE\Microsoft\Office\16.0\Common\InstallRoot",
        "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Office\16.0\Common\InstallRoot"
    )
    
    foreach ($path in $o16msi_paths) {
        if (Test-Path $path) {
            $installPath = (Get-ItemProperty -Path $path -Name Path -ErrorAction SilentlyContinue).Path
            if ($installPath -and (Test-Path "$installPath\*Picker.dll")) {
                $officeInstalls += [PSCustomObject]@{
                    Version = "16.0"
                    Type = "MSI"
                    Path = $installPath
                    VersionNumber = "16.0"
                    Architecture = if ($path -like "*Wow6432Node*") { "x86" } else { "x64" }
                    Registry = $path -replace "\\Common\\InstallRoot", ""
                }
            }
        }
    }
    
    return $officeInstalls
}

# Function to download Ohook files from GitHub
function Get-OhookFiles {
    param([string]$Architecture)
    
    $tempPath = Join-Path $env:TEMP "Ohook"
    if (-not (Test-Path $tempPath)) {
        New-Item -ItemType Directory -Path $tempPath -Force | Out-Null
    }
    
    $dllName = if ($Architecture -eq "x64") { "sppc64.dll" } else { "sppc32.dll" }
    $dllPath = Join-Path $tempPath $dllName
    
    # Download from GitHub (you'll need to host these files)
    $githubUrl = "https://raw.githubusercontent.com/massgravel/Microsoft-Activation-Scripts/master/BIN/$dllName"
    
    try {
        Write-Host "Downloading $dllName..." -ForegroundColor DarkCyan
        Invoke-WebRequest -Uri $githubUrl -OutFile $dllPath -UseBasicParsing -ErrorAction Stop
        return $dllPath
    }
    catch {
        Write-ColorText "Failed to download $dllName from GitHub: $_" "DarkRed"
        return $null
    }
}

# Function to install Ohook DLL
function Install-OhookDLL {
    param(
        [string]$HookPath,
        [string]$SppcPath,
        [string]$Architecture
    )
    
    try {
        # Create symlink for sppcs.dll pointing to system sppc.dll
        $sppcsPath = Join-Path $HookPath "sppcs.dll"
        $sppcDllPath = Join-Path $HookPath "sppc.dll"
        
        # Remove old files if they exist
        if (Test-Path $sppcsPath) { Remove-Item $sppcsPath -Force }
        if (Test-Path $sppcDllPath) { Remove-Item $sppcDllPath -Force }
        
        # Create symbolic link
        cmd /c mklink "$sppcsPath" "$SppcPath" 2>&1 | Out-Null
        
        if (-not (Test-Path $sppcsPath)) {
            throw "Failed to create symbolic link"
        }
        
        # Download and copy custom sppc.dll
        $customDll = Get-OhookFiles -Architecture $Architecture
        if (-not $customDll) {
            throw "Failed to download custom DLL"
        }
        
        Copy-Item $customDll $sppcDllPath -Force
        
        if (-not (Test-Path $sppcDllPath)) {
            throw "Failed to copy custom DLL"
        }
        
        Write-Host "Symlinking System's sppc.dll            [$sppcsPath] [Successful]" -ForegroundColor DarkGreen
        Write-Host "Copying Custom sppc.dll                 [$sppcDllPath] [Successful]" -ForegroundColor DarkGreen
        
        return $true
    }
    catch {
        Write-ColorText "Installing Ohook DLL Failed: $_" "DarkRed"
        return $false
    }
}

#============================================================================
# Core Logic Functions
#============================================================================

function Install-OhookActivation {
    Write-Host "Initializing Ohook installation..." -ForegroundColor DarkCyan
    Write-Host ""
    
    # Detect Office installations
    $officeInstalls = Get-OfficeInstallation
    
    if ($officeInstalls.Count -eq 0) {
        Write-ColorText "No supported Office installation found." "DarkRed"
        Write-ColorText "Supported versions: Office 2013, 2016, 2019, 2021, 2024 (C2R or MSI)" "DarkYellow"
        Write-Host ""
        Read-Host "Press Enter to continue"
        return
    }
    
    $activated = $false
    
    foreach ($office in $officeInstalls) {
        Write-Host "Found Office $($office.Version) $($office.Type) [$($office.VersionNumber) | $($office.Architecture)]" -ForegroundColor DarkCyan
        
        # Determine hook path based on architecture and type
        if ($office.Type -eq "C2R") {
            if ($office.Architecture -eq "x64") {
                $hookPath = Join-Path $office.Path "vfs\System"
                $sppcPath = "$env:SystemRoot\System32\sppc.dll"
            } else {
                $hookPath = Join-Path $office.Path "vfs\SystemX86"
                $sppcPath = "$env:SystemRoot\SysWOW64\sppc.dll"
            }
        } else {
            # MSI installation
            $hookPath = $office.Path
            if ($office.Architecture -eq "x64") {
                $sppcPath = "$env:SystemRoot\System32\sppc.dll"
            } else {
                $sppcPath = "$env:SystemRoot\SysWOW64\sppc.dll"
            }
        }
        
        # Install Ohook DLL
        Write-Host "Installing Ohook to $hookPath..." -ForegroundColor DarkCyan
        $result = Install-OhookDLL -HookPath $hookPath -SppcPath $sppcPath -Architecture $office.Architecture
        
        if ($result) {
            $activated = $true
            Write-ColorText "Office $($office.Version) activated successfully!" "DarkGreen"
        } else {
            Write-ColorText "Failed to activate Office $($office.Version)" "DarkRed"
        }
        
        Write-Host ""
    }
    
    if ($activated) {
        Write-ColorText "Office is permanently activated." "DarkGreen"
        Write-ColorText "For help, visit: $($mas)troubleshoot" "Black"
    } else {
        Write-ColorText "Activation failed. Please check the errors above." "DarkRed"
    }
    
    Write-Host ""
    Read-Host "Press Enter to continue"
}

function Uninstall-OhookActivation {
    Write-Host "Uninstalling Ohook activation..." -ForegroundColor DarkCyan
    Write-Host ""
    
    $removed = $false
    
    # Detect Office installations
    $officeInstalls = Get-OfficeInstallation
    
    foreach ($office in $officeInstalls) {
        # Determine hook path
        if ($office.Type -eq "C2R") {
            $hookPaths = @(
                (Join-Path $office.Path "vfs\System"),
                (Join-Path $office.Path "vfs\SystemX86")
            )
        } else {
            $hookPaths = @($office.Path)
        }
        
        foreach ($hookPath in $hookPaths) {
            # Remove sppc*.dll files
            $files = Get-ChildItem -Path $hookPath -Filter "sppc*.dll" -ErrorAction SilentlyContinue
            
            foreach ($file in $files) {
                try {
                    Remove-Item $file.FullName -Force
                    Write-Host "Removed: $($file.FullName)" -ForegroundColor DarkGreen
                    $removed = $true
                }
                catch {
                    Write-ColorText "Failed to remove: $($file.FullName)" "DarkRed"
                }
            }
        }
    }
    
    # Remove registry keys
    $kmskey = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\0ff1ce15-a989-479d-af46-f275c6370663"
    if (Test-Path $kmskey) {
        Remove-Item $kmskey -Force -ErrorAction SilentlyContinue
        Write-Host "Removed registry key for preventing non-genuine banner" -ForegroundColor DarkGreen
        $removed = $true
    }
    
    Write-Host ""
    
    if ($removed) {
        Write-ColorText "Successfully uninstalled Ohook activation." "DarkGreen"
    } else {
        Write-ColorText "Ohook activation is not installed." "DarkYellow"
    }
    
    Write-Host ""
    Read-Host "Press Enter to continue"
}

#============================================================================
# User Interface and Script Execution
#============================================================================

function Show-Menu {
    Clear-Host
    Write-Host "============================================================" -ForegroundColor DarkGreen
    Write-Host "  Ohook Activation $masver (PowerShell Version)" -ForegroundColor Black
    Write-Host "============================================================" -ForegroundColor DarkGreen
    Write-Host
    Write-Host "         [1] Install Ohook Office Activation"
    Write-Host "         [2] Uninstall Ohook"
    Write-Host "         [3] Download Office"
    Write-Host
    Write-Host "         [0] Exit"
    Write-Host
    Write-Host "------------------------------------------------------------"
}

# --- Main Script Body ---

try {
    # First, handle unattended execution if parameters were passed
    if ($unattended) {
        if (-not (Test-RunningOfficeApps)) { 
            Write-Host "Office applications are running. Please close them first." -ForegroundColor Red
            Read-Host "Press Enter to exit"
            exit 1 
        }
        
        if ($Ohook.IsPresent) {
            Install-OhookActivation
        }
        elseif ($Ohook_Uninstall.IsPresent) {
            Uninstall-OhookActivation
        }
        
        Write-Host ""
        Write-Host "Operation completed." -ForegroundColor Green
        Read-Host "Press Enter to exit"
        exit 0
    }

    # If not unattended, show the interactive menu
    do {
        Show-Menu
        
        # Check for running apps and display a warning if necessary
        $appsRunning = -not (Test-RunningOfficeApps)
        if ($appsRunning) {
            Write-Host
        }
        
        $choice = Read-Host "Choose a menu option [1,2,3,0]"
        
        switch ($choice) {
            '1' {
                if (-not $appsRunning) {
                    Install-OhookActivation
                }
                else {
                    Read-Host "Press Enter to return to the menu..."
                }
            }
            '2' {
                if (-not $appsRunning) {
                    Uninstall-OhookActivation
                }
                else {
                    Read-Host "Press Enter to return to the menu..."
                }
            }
            '3' {
                # Opens the link in the default web browser
                Start-Process "$($mas)genuine-installation-media"
            }
            '0' {
                Write-Host "Exiting."
                Start-Sleep -Seconds 1
            }
            default {
                Write-ColorText "Invalid option. Please try again." "DarkRed"
                Start-Sleep -Seconds 1
            }
        }
    } while ($choice -ne '0')

    Write-Host "`nScript completed. You can now close this window." -ForegroundColor DarkGreen
}
finally {
    Write-Host ""
    Write-Host "Press any key to close this window..." -ForegroundColor DarkCyan
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
}
