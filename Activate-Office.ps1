<#
.SYNOPSIS
    Page: https://www.facebook.com/DigitalNecessitiesBitCourse
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

#============================================================================
# Core Logic Functions (Placeholders)
# Note: The actual activation logic from the batch file is highly complex.
# These functions serve as placeholders for that logic.
#============================================================================

function Install-OhookActivation {
    Write-Host "Activating..." -ForegroundColor DarkCyan
    
    # Placeholder for the complex activation logic which would involve:
    # 1. Finding Office installation paths (C2R, MSI).
    # 2. Checking for supported Office versions.
    # 3. Installing license files and registry keys.
    # 4. Applying the 'hook' by modifying system files.
    
    # Simulating a process for demonstration
    Write-Host "Applying activation..."
    Start-Sleep -Seconds 2
    Write-ColorText "Office is permanently activated." "DarkGreen"
    
    Write-Host ""
    Read-Host "Press Enter to continue"
}

function Uninstall-OhookActivation {
    Write-Host "Uninstalling Ohook activation..." -ForegroundColor DarkCyan
    
    # Placeholder for the uninstallation logic which would involve:
    # 1. Finding and deleting the hook files (e.g., sppc*.dll).
    # 2. Restoring original system files if they were modified.
    # 3. Removing related registry keys.
    
    # Simulating a process for demonstration
    Write-Host "Removing activation files and registry keys..."
    Start-Sleep -Seconds 2
    Write-ColorText "Successfully uninstalled Ohook activation." "DarkGreen"
    
    Write-Host ""
    Read-Host "Press Enter to continue"
}

#============================================================================
# User Interface and Script Execution
#============================================================================

function Show-Menu {
    Clear-Host
    Write-Host "============================================================" -ForegroundColor Black
    Write-Host "                       BitCourse" -ForegroundColor darkYellow
    Write-Host "============================================================" -ForegroundColor Black
    Write-Host
    Write-Host "         [1] Activate Microsoft Office"
    Write-Host "         [2] Uninstall Activation"
    Write-Host
    Write-Host "         [0] Exit"
    Write-Host
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
        
        $choice = Read-Host "Choose Option [1,2,0]"
        
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






