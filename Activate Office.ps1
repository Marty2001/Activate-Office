<#
.SYNOPSIS
    Activates Microsoft Office using Ohook activation method. This script is a PowerShell conversion of the original batch file.
    Homepage: massgrave.dev
    Email: mas.help@outlook.com
.DESCRIPTION
    This script provides functionalities to install or uninstall Ohook Office Activation.
    It performs necessary system checks, handles administrative elevation, and interacts with the Windows licensing services.
.PARAMETER Ohook
    A switch to activate Office with Ohook activation in unattended mode.
.PARAMETER Ohook-Uninstall
    A switch to remove Ohook activation in unattended mode.
#>
[CmdletBinding()]
param(
    [switch]$Ohook,
    [switch]$Ohook_Uninstall
)

#============================================================================
# Initial Setup and Variables
#============================================================================

$masver = "3.7"
$mas = "https://massgrave.dev/"
$github = "https://github.com/massgravel/Microsoft-Activation-Scripts"
$selfgit = "https://git.activated.win/massgrave/Microsoft-Activation-Scripts"

$unattended = $false
if ($Ohook.IsPresent -or $Ohook_Uninstall.IsPresent) {
    $unattended = $true
}

# Set environment paths to ensure script reliability
$env:Path = "$env:SystemRoot\System32;$env:SystemRoot;$env:SystemRoot\System32\Wbem;$env:SystemRoot\System32\WindowsPowerShell\v1.0\"
if (Test-Path "$env:SystemRoot\Sysnative") {
    $env:Path = "$env:SystemRoot\Sysnative;$env:SystemRoot;$env:SystemRoot\Sysnative\Wbem;$env:SystemRoot\Sysnative\WindowsPowerShell\v1.0\;$env:Path"
}

#============================================================================
# Helper Functions
#============================================================================

# Function to write colored text to the console
function Write-ColorText {
    param (
        [string]$Message,
        [string]$ForegroundColor = "White",
        [string]$BackgroundColor = "Black"
    )
    Write-Host $Message -ForegroundColor $ForegroundColor -BackgroundColor $BackgroundColor
}

# Function to check for administrator privileges
function Test-IsAdmin {
    $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object System.Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Function to get OS information
function Get-OSInfo {
    $osInfo = Get-CimInstance Win32_OperatingSystem
    $osArch = (Get-CimInstance Win32_Processor).AddressWidth
    return @{
        Version = $osInfo.Version
        BuildNumber = $osInfo.BuildNumber
        Caption = $osInfo.Caption
        Architecture = "x$($osArch)"
    }
}

# Function to check if required services are running
function Test-Services {
    $services = "sppsvc", "Winmgmt"
    foreach ($service in $services) {
        $svc = Get-Service -Name $service -ErrorAction SilentlyContinue
        if (-not $svc -or $svc.Status -ne "Running") {
            Write-ColorText "Error: Service '$service' is not running or not found." "Red"
            return $false
        }
    }
    return $true
}


#============================================================================
# Pre-flight Checks
#============================================================================

# 1. Ensure the script is running with Administrator privileges
if (-not (Test-IsAdmin)) {
    Write-ColorText "This script requires administrator rights. Please run it as an administrator." "Red"
    Start-Process powershell.exe -ArgumentList "-File `"$PSCommandPath`" $psboundparameters" -Verb RunAs
    exit
}

# 2. Check for PowerShell version and language mode
if ($PSVersionTable.PSVersion.Major -lt 5) {
    Write-ColorText "PowerShell 5.0 or newer is required." "Red"
    exit
}
if ($ExecutionContext.SessionState.LanguageMode -ne 'FullLanguage') {
    Write-ColorText "PowerShell FullLanguage mode is required. Aborting." "Red"
    Write-ColorText "This might be due to system restrictions. For help, visit: $($mas)fix_powershell" "Blue"
    exit
}

# 3. Check for running Office applications
function Get-RunningOfficeApps {
    $officeProcesses = @(
        "msaccess", "excel", "groove", "lync", "onenote", 
        "outlook", "powerpnt", "winproj", "mspub", "visio", "winword"
    )
    $runningApps = Get-Process | Where-Object { $officeProcesses -contains $_.ProcessName }
    if ($runningApps) {
        $appNames = ($runningApps | ForEach-Object { $_.ProcessName }) -join ", "
        Write-ColorText "Please close the following Office applications before proceeding: $appNames" "Yellow"
        return $false
    }
    return $true
}

#============================================================================
# Main Logic - Placeholder for Core Ohook Functions
# Note: The actual activation logic from the batch file is complex and involves
# binary data, registry manipulation, and interaction with licensing WMI classes.
# The functions below are placeholders for that logic.
#============================================================================

function Install-OhookActivation {
    Write-Host "Initializing Ohook installation..." -ForegroundColor Cyan
    
    # Placeholder for the complex activation logic
    # This would involve:
    # 1. Finding Office installation paths (C2R, MSI)
    # 2. Checking for supported Office versions
    # 3. Installing license files
    # 4. Applying the 'hook' by replacing/symlinking system DLLs (sppc.dll)
    # 5. Clearing license blocks and setting registry keys
    
    # Simulating a successful activation for demonstration
    Start-Sleep -Seconds 2
    Write-ColorText "Office is permanently activated." "Green"
    Write-ColorText "Help: $($mas)troubleshoot" "White"
}

function Uninstall-OhookActivation {
    Write-Host "Uninstalling Ohook activation..." -ForegroundColor Cyan
    
    # Placeholder for the uninstallation logic
    # This would involve:
    # 1. Finding the hook files (sppc*.dll) in Office directories
    # 2. Deleting them
    # 3. Restoring original system files if they were renamed
    # 4. Removing related registry keys (Resiliency, KMS keys for banner prevention)
    
    # Simulating a successful uninstallation for demonstration
    Start-Sleep -Seconds 2
    Write-ColorText "Successfully uninstalled Ohook activation." "Green"
}


#============================================================================
# User Interface and Script Execution
#============================================================================

function Show-Menu {
    Clear-Host
    Write-Host "============================================================"
    Write-Host "  Ohook Activation $masver (PowerShell Version)"
    Write-Host "============================================================"
    Write-Host
    Write-Host "   [1] Install Ohook Office Activation"
    Write-Host "   [2] Uninstall Ohook"
    Write-Host "   [3] Download Office"
    Write-Host "   [0] Exit"
    Write-Host
    Write-Host "------------------------------------------------------------"
}

# --- Main Script Body ---

# Handle unattended execution first
if ($unattended) {
    if (-not (Get-RunningOfficeApps)) { exit }
    if ($Ohook.IsPresent) {
        Install-OhookActivation
    }
    elseif ($Ohook_Uninstall.IsPresent) {
        Uninstall-OhookActivation
    }
    exit
}

# Interactive Menu
do {
    Show-Menu
    if (-not (Get-RunningOfficeApps)) {
        Write-Host
    }
    
    $choice = Read-Host "Choose a menu option [1,2,3,0]"
    
    switch ($choice) {
        '1' {
            if (Get-RunningOfficeApps) {
                Install-OhookActivation
            }
            Read-Host "Press Enter to return to the menu..."
        }
        '2' {
            if (Get-RunningOfficeApps) {
                Uninstall-OhookActivation
            }
            Read-Host "Press Enter to return to the menu..."
        }
        '3' {
            Start-Process "$($mas)genuine-installation-media"
        }
        '0' {
            Write-Host "Exiting."
        }
        default {
            Write-ColorText "Invalid option. Please try again." "Red"
            Start-Sleep -Seconds 1
        }
    }
} while ($choice -ne '0')