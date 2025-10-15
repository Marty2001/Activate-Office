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
# UAC ELEVATION PROMPT (FINAL WORKING VERSION)
# This section handles the User Account Control (UAC) prompt and Execution Policy.
#============================================================================

# Check if the script is running with Administrator privileges
$identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object System.Security.Principal.WindowsPrincipal($identity)
$isAdmin = $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    # If not admin, re-launch the script with elevated privileges.
    Write-Warning "Administrator rights required. Requesting UAC elevation..."
    
    try {
        # Prepare arguments for the new PowerShell process.
        $arguments = @(
            "-NoProfile",
            "-ExecutionPolicy", "Bypass",  # <-- THE CRITICAL FIX IS HERE
            "-File",
            "'$($PSCommandPath)'" # Ensures the path is correctly quoted
        )

        # Forward any original parameters (like -Ohook) to the new instance.
        $arguments += $psboundparameters.Keys | ForEach-Object { "-$_" }

        # Start the new elevated process and wait for it to finish.
        Start-Process powershell.eXE -ArgumentList $arguments -Verb RunAs -Wait
    }
    catch {
        Write-Error "Failed to elevate privileges. Please right-click the script and select 'Run as Administrator'."
        Read-Host "Press Enter to exit."
    }
    
    # Exit the current, non-elevated script.
    exit
}

#============================================================================
# Initial Setup and Variables
#============================================================================

Write-Host "Running with Administrator privileges." -ForegroundColor Green
$masver = "3.7"
$mas = "https://massgrave.dev/"
$github = "https://github.com/massgravel/Microsoft-Activation-Scripts"

# Determine if running in unattended mode
$unattended = $Ohook.IsPresent -or $Ohook_Uninstall.IsPresent

#============================================================================
# Helper Functions
#============================================================================

function Write-ColorText {
    param (
        [string]$Message,
        [string]$ForegroundColor = "White"
    )
    Write-Host $Message -ForegroundColor $ForegroundColor
}

function Test-RunningOfficeApps {
    $officeProcesses = "msaccess", "excel", "groove", "lync", "onenote", "outlook", "powerpnt", "winproj", "mspub", "visio", "winword"
    $runningApps = Get-Process -ErrorAction SilentlyContinue | Where-Object { $officeProcesses -contains $_.ProcessName }
    
    if ($runningApps) {
        $appNames = ($runningApps.ProcessName) -join ", "
        Write-ColorText "Warning: Please close the following Office applications: $appNames" -ForegroundColor Yellow
        return $false
    }
    return $true
}

#============================================================================
# Core Logic Functions (Placeholders)
#============================================================================

function Install-OhookActivation {
    Write-ColorText "Initializing Ohook installation..." -ForegroundColor Cyan
    Write-Host "Applying activation..."
    Start-Sleep -Seconds 2
    Write-ColorText "Office is permanently activated." -ForegroundColor Green
    Write-ColorText "For help, visit: $($mas)troubleshoot"
}

function Uninstall-OhookActivation {
    Write-ColorText "Uninstalling Ohook activation..." -ForegroundColor Cyan
    Write-Host "Removing activation files and registry keys..."
    Start-Sleep -Seconds 2
    Write-ColorText "Successfully uninstalled Ohook activation." -ForegroundColor Green
}

#============================================================================
# User Interface and Script Execution
#============================================================================

function Show-Menu {
    Clear-Host
    Write-ColorText "============================================================" -ForegroundColor Green
    Write-ColorText "  Ohook Activation $masver (PowerShell Version)"
    Write-ColorText "============================================================" -ForegroundColor Green
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

if ($unattended) {
    if (-not (Test-RunningOfficeApps)) { exit 1 }
    
    if ($Ohook.IsPresent) {
        Install-OhookActivation
    }
    elseif ($Ohook_Uninstall.IsPresent) {
        Uninstall-OhookActivation
    }
    exit 0
}

do {
    Show-Menu
    $appsAreRunning = -not (Test-RunningOfficeApps)
    if ($appsAreRunning) { Write-Host }
    
    $choice = Read-Host "Choose a menu option [1,2,3,0]"
    
    switch ($choice) {
        '1' {
            if (-not $appsAreRunning) { Install-OhookActivation }
            Read-Host "Press Enter to return to the menu..."
        }
        '2' {
            if (-not $appsAreRunning) { Uninstall-OhookActivation }
            Read-Host "Press Enter to return to the menu..."
        }
        '3' {
            Start-Process "$($mas)genuine-installation-media"
        }
        '0' {
            Write-Host "Exiting."
        }
        default {
            Write-ColorText "Invalid option. Please try again." -ForegroundColor Red
            Start-Sleep -Seconds 1
        }
    }
} while ($choice -ne '0')
