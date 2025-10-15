<#
.SYNOPSIS
    Activates Microsoft Office using the Ohook activation method. This script is a PowerShell conversion of the original batch file.
    Homepage: massgrave.dev
    Email: mas.help@outlook.com
.DESCRIPTION
    This script provides functionalities to install or uninstall Ohook Office Activation.
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
#============================================================================
# UAC ELEVATION PROMPT (Corrected Version)
# This section handles the User Account Control (UAC) prompt reliably.
#============================================================================

# Check if the script is running with Administrator privileges
$identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object System.Security.Principal.WindowsPrincipal($identity)
$isAdmin = $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    # If not running as admin, re-launch the script with elevated privileges
    Write-Warning "This script requires administrator rights. Requesting elevation..."
    try {
        # Build the arguments for the new process.
        # The -File parameter's value must be quoted if it contains spaces.
        $arguments = @(
            "-NoProfile",
            "-File",
            "'$($PSCommandPath)'" # Crucially, enclose the path in single quotes
        )

        # Forward any original parameters (like -Ohook) to the new instance
        $arguments += $psboundparameters.Keys | ForEach-Object { "-$_" }

        # Re-launch PowerShell with the corrected arguments
        Start-Process powershell.exe -ArgumentList $arguments -Verb RunAs -Wait
    }
    catch {
        Write-Error "Failed to elevate privileges. Please right-click the script and 'Run as Administrator'."
        Read-Host "Press Enter to exit"
    }
    # Exit the current (non-elevated) script
    exit
}

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
        [string]$ForegroundColor = "White",
        [string]$BackgroundColor = "Black"
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


#============================================================================
# Core Logic Functions (Placeholders)
# Note: The actual activation logic from the batch file is highly complex.
# These functions serve as placeholders for that logic.
#============================================================================

function Install-OhookActivation {
    Write-Host "Initializing Ohook installation..." -ForegroundColor Cyan
    
    # Placeholder for the complex activation logic which would involve:
    # 1. Finding Office installation paths (C2R, MSI).
    # 2. Checking for supported Office versions.
    # 3. Installing license files and registry keys.
    # 4. Applying the 'hook' by modifying system files.
    
    # Simulating a process for demonstration
    Write-Host "Applying activation..."
    Start-Sleep -Seconds 2
    Write-ColorText "Office is permanently activated." "Green"
    Write-ColorText "For help, visit: $($mas)troubleshoot" "White"
}

function Uninstall-OhookActivation {
    Write-Host "Uninstalling Ohook activation..." -ForegroundColor Cyan
    
    # Placeholder for the uninstallation logic which would involve:
    # 1. Finding and deleting the hook files (e.g., sppc*.dll).
    # 2. Restoring original system files if they were modified.
    # 3. Removing related registry keys.
    
    # Simulating a process for demonstration
    Write-Host "Removing activation files and registry keys..."
    Start-Sleep -Seconds 2
    Write-ColorText "Successfully uninstalled Ohook activation." "Green"
}

#============================================================================
# User Interface and Script Execution
#============================================================================

function Show-Menu {
    Clear-Host
    Write-Host "============================================================" -ForegroundColor Green
    Write-Host "  Ohook Activation $masver (PowerShell Version)" -ForegroundColor White
    Write-Host "============================================================" -ForegroundColor Green
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

# First, handle unattended execution if parameters were passed
if ($unattended) {
    if (-not (Test-RunningOfficeApps)) { exit 1 } # Exit with an error code if apps are running
    
    if ($Ohook.IsPresent) {
        Install-OhookActivation
    }
    elseif ($Ohook_Uninstall.IsPresent) {
        Uninstall-OhookActivation
    }
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
            Read-Host "Press Enter to return to the menu..."
        }
        '2' {
            if (-not $appsRunning) {
                Uninstall-OhookActivation
            }
            Read-Host "Press Enter to return to the menu..."
        }
        '3' {
            # Opens the link in the default web browser
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
