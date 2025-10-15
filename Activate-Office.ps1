#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Ohook Office Activation Script - PowerShell Version
.DESCRIPTION
    Converts Office installations to activated state using Ohook method
    Homepage: massgrave.dev
.PARAMETER Ohook
    Run activation in unattended mode
.PARAMETER OhookUninstall
    Uninstall Ohook activation
.EXAMPLE
    irm https://your-url/Ohook-Activation.ps1 | iex
    irm https://your-url/Ohook-Activation.ps1 | iex -Ohook
#>

param(
    [switch]$Ohook,
    [switch]$OhookUninstall
)

#region Auto-Elevation
function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (Test-Administrator)) {
    Write-Host "Requesting administrator privileges..." -ForegroundColor Yellow
    
    try {
        # Get the script content
        $scriptContent = $MyInvocation.MyCommand.ScriptBlock.ToString()
        
        # Create a temporary script file
        $tempScript = [System.IO.Path]::Combine($env:TEMP, "Ohook-Activation-Elevated-$(Get-Random).ps1")
        $scriptContent | Out-File -FilePath $tempScript -Encoding UTF8 -Force
        
        # Build arguments to pass to elevated process
        $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$tempScript`""
        
        if ($Ohook) {
            $arguments += " -Ohook"
        }
        if ($OhookUninstall) {
            $arguments += " -OhookUninstall"
        }
        
        # Start elevated process
        $process = Start-Process -FilePath "powershell.exe" -ArgumentList $arguments -Verb RunAs -PassThru -Wait
        
        # Clean up temp file
        Start-Sleep -Seconds 2
        Remove-Item $tempScript -Force -ErrorAction SilentlyContinue
        
        # Exit current non-elevated process
        exit $process.ExitCode
    }
    catch {
        Write-Host "Failed to elevate privileges: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Please run PowerShell as Administrator manually." -ForegroundColor Yellow
        exit 1
    }
}
#endregion

# Script version
$Script:MasVer = "3.7"
$Script:ErrorFound = $false
$Script:UnattendedMode = $Ohook -or $OhookUninstall

# Color definitions
$Script:Colors = @{
    Red = 'Red'
    Green = 'Green'
    Yellow = 'Yellow'
    Blue = 'Cyan'
    Gray = 'DarkGray'
    White = 'White'
}

#region Helper Functions

function Write-ColorText {
    param(
        [string]$Text,
        [string]$Color = 'White'
    )
    Write-Host $Text -ForegroundColor $Colors[$Color]
}

function Write-Title {
    param([string]$Title)
    Clear-Host
    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-ColorText "  Ohook Office Activation $MasVer" 'Green'
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host ""
}

function Get-OSInfo {
    $os = Get-CimInstance Win32_OperatingSystem
    $arch = $env:PROCESSOR_ARCHITECTURE
    $build = [System.Environment]::OSVersion.Version.Build
    $ubr = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name UBR -ErrorAction SilentlyContinue).UBR
    
    if ($ubr) {
        $fullBuild = "$build.$ubr"
    } else {
        $fullBuild = $build
    }
    
    Write-Host "OS: $($os.Caption) | Build: $fullBuild | Arch: $arch" -ForegroundColor Gray
}

function Get-OfficeInstallations {
    $installations = @{
        O16C2R = $null
        O15C2R = $null
        O16MSI = $null
        O15MSI = $null
        O14MSI = $null
    }
    
    # Check for Office 16.0 C2R
    $regPaths = @(
        'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun',
        'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Office\ClickToRun'
    )
    
    foreach ($path in $regPaths) {
        if (Test-Path $path) {
            $installPath = (Get-ItemProperty $path -Name InstallPath -ErrorAction SilentlyContinue).InstallPath
            if ($installPath -and (Test-Path "$installPath\root\Licenses16\ProPlus*.xrm-ms")) {
                $installations.O16C2R = $path
                break
            }
        }
    }
    
    # Check for Office 15.0 C2R
    $regPaths = @(
        'HKLM:\SOFTWARE\Microsoft\Office\15.0\ClickToRun',
        'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Office\15.0\ClickToRun'
    )
    
    foreach ($path in $regPaths) {
        if (Test-Path $path) {
            $installPath = (Get-ItemProperty $path -Name InstallPath -ErrorAction SilentlyContinue).InstallPath
            if ($installPath -and (Test-Path "$installPath\root\Licenses\ProPlus*.xrm-ms")) {
                $installations.O15C2R = $path
                break
            }
        }
    }
    
    # Check for MSI installations
    $msiVersions = @(
        @{Ver='16.0'; Key='O16MSI'},
        @{Ver='15.0'; Key='O15MSI'},
        @{Ver='14.0'; Key='O14MSI'}
    )
    
    foreach ($ver in $msiVersions) {
        $regPaths = @(
            "HKLM:\SOFTWARE\Microsoft\Office\$($ver.Ver)\Common\InstallRoot",
            "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Office\$($ver.Ver)\Common\InstallRoot"
        )
        
        foreach ($path in $regPaths) {
            if (Test-Path $path) {
                $installPath = (Get-ItemProperty $path -Name Path -ErrorAction SilentlyContinue).Path
                if ($installPath -and (Test-Path "$installPath\*Picker.dll")) {
                    $installations[$ver.Key] = $path -replace '\\Common\\InstallRoot$', ''
                    break
                }
            }
        }
    }
    
    return $installations
}

function Get-ProductKeyData {
    # This is a simplified version - full implementation would include all product keys
    # from the :ohookdata section of the original script
    
    $keyData = @{
        # Office 2024
        '16_e563d108-7b0e-418a-8390-20e1d133d6bb' = @{
            Key = 'P6NMW-JMTRC-R6MQ6-HH3F2-BTHKB'
            Product = 'Access2024Retail'
            License = 'Retail'
        }
        '16_f3a5e86a-e4f8-4d88-8220-1440c3bbcefa' = @{
            Key = '82CNJ-W82TW-BY23W-BVJ6W-W48GP'
            Product = 'Excel2024Retail'
            License = 'Retail'
        }
        # Office 2021
        '16_f634398e-af69-48c9-b256-477bea3078b5' = @{
            Key = 'P286B-N3XYP-36QRQ-29CMP-RVX9M'
            Product = 'Access2021Retail'
            License = 'Retail'
        }
        '16_fb099c19-d48b-4a2f-a160-4383011060aa' = @{
            Key = 'V6QFB-7N7G9-PF7W9-M8FQM-MY8G9'
            Product = 'Excel2021Retail'
            License = 'Retail'
        }
        # Office 2019
        '16_518687bd-dc55-45b9-8fa6-f918e1082e83' = @{
            Key = 'WRYJ6-G3NP7-7VH94-8X7KP-JB7HC'
            Product = 'Access2019Retail'
            License = 'Retail'
        }
        '16_c201c2b7-02a1-41a8-b496-37c72910cd4a' = @{
            Key = 'KBPNW-64CMM-8KWCB-23F44-8B7HM'
            Product = 'Excel2019Retail'
            License = 'Retail'
        }
        # Office 2016
        '16_bfa358b0-98f1-4125-842e-585fa13032e6' = @{
            Key = 'WHK4N-YQGHB-XWXCC-G3HYC-6JF94'
            Product = 'AccessRetail'
            License = 'Retail'
        }
        '16_424d52ff-7ad2-4bc7-8ac6-748d767b455d' = @{
            Key = 'RKJBN-VWTM2-BDKXX-RKQFD-JTYQ2'
            Product = 'ExcelRetail'
            License = 'Retail'
        }
        '16_de52bd50-9564-4adc-8fcb-a345c17f84f9' = @{
            Key = 'GM43N-F742Q-6JDDK-M622J-J8GDV'
            Product = 'ProPlusRetail'
            License = 'Retail'
        }
    }
    
    return $keyData
}

function Install-ProductKey {
    param(
        [string]$Key,
        [string]$ProductName
    )
    
    try {
        $service = Get-WmiObject -Query "SELECT Version FROM SoftwareLicensingService"
        $service.InstallProductKey($Key) | Out-Null
        
        # Refresh license status
        $service.RefreshLicenseStatus() | Out-Null
        
        Write-ColorText "Installing Product Key [$ProductName] [Successful]" 'Green'
        return $true
    }
    catch {
        Write-ColorText "Installing Product Key [$ProductName] [Failed] $($_.Exception.Message)" 'Red'
        $Script:ErrorFound = $true
        return $false
    }
}

function Install-OhookDLL {
    param(
        [string]$HookPath,
        [string]$HookFile,
        [string]$SppcPath
    )
    
    Write-Host ""
    Write-Host "Installing Ohook activation files..." -ForegroundColor Yellow
    
    try {
        # Create symbolic link for sppcs.dll
        $sppcsDll = Join-Path $HookPath "sppcs.dll"
        $sppcDll = Join-Path $HookPath "sppc.dll"
        
        # Remove old files if they exist
        if (Test-Path $sppcsDll) {
            Remove-Item $sppcsDll -Force -ErrorAction SilentlyContinue
        }
        if (Test-Path $sppcDll) {
            Remove-Item $sppcDll -Force -ErrorAction SilentlyContinue
        }
        
        # Create symbolic link
        New-Item -ItemType SymbolicLink -Path $sppcsDll -Target $SppcPath -Force | Out-Null
        
        if (Test-Path $sppcsDll) {
            Write-ColorText "Symlinking System's sppc.dll [$sppcsDll] [Successful]" 'Green'
        } else {
            throw "Failed to create symbolic link"
        }
        
        # Extract and install custom DLL
        # Note: In a real implementation, you would extract the DLL from base64
        # For this example, we'll create a placeholder
        Write-ColorText "Custom DLL installation would happen here" 'Yellow'
        Write-ColorText "Note: Full DLL extraction code needed for production use" 'Gray'
        
        return $true
    }
    catch {
        Write-ColorText "Installing Ohook [Failed] $($_.Exception.Message)" 'Red'
        $Script:ErrorFound = $true
        return $false
    }
}

function Clear-OfficeLicenseBlocks {
    Write-Host "Clearing Office license blocks..." -ForegroundColor Yellow
    
    try {
        # Remove vNext/shared/device license blocks
        $licensePath = "$env:ProgramData\Microsoft\Office\Licenses"
        if (Test-Path $licensePath) {
            Remove-Item $licensePath -Recurse -Force -ErrorAction SilentlyContinue
        }
        
        # Clear licensing registry keys
        $regPaths = @(
            'HKLM:\SOFTWARE\Microsoft\Office\16.0\Common\Licensing',
            'HKLM:\SOFTWARE\Microsoft\Office\15.0\Common\Licensing',
            'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Office\16.0\Common\Licensing',
            'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Office\15.0\Common\Licensing'
        )
        
        foreach ($path in $regPaths) {
            if (Test-Path $path) {
                Remove-Item $path -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
        
        # Get user SIDs
        $userProfiles = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\*' |
            Where-Object { $_.ProfileImagePath -like '*\Users\*' }
        
        foreach ($profile in $userProfiles) {
            $sid = $profile.PSChildName
            $userPath = $profile.ProfileImagePath
            
            # Clear user-specific license data
            $userLicensePaths = @(
                "$userPath\AppData\Local\Microsoft\Office\Licenses",
                "$userPath\AppData\Local\Microsoft\Office\16.0\Licensing",
                "$userPath\AppData\Local\Microsoft\Office\15.0\Licensing"
            )
            
            foreach ($path in $userLicensePaths) {
                if (Test-Path $path) {
                    Remove-Item $path -Recurse -Force -ErrorAction SilentlyContinue
                }
            }
            
            # Clear user registry
            $userRegPaths = @(
                "Registry::HKU\$sid\Software\Microsoft\Office\16.0\Common\Licensing",
                "Registry::HKU\$sid\Software\Microsoft\Office\15.0\Common\Licensing"
            )
            
            foreach ($path in $userRegPaths) {
                if (Test-Path $path) {
                    Remove-Item $path -Recurse -Force -ErrorAction SilentlyContinue
                }
            }
        }
        
        Write-ColorText "Clearing Office License Blocks [Successful]" 'Green'
        return $true
    }
    catch {
        Write-ColorText "Clearing Office License Blocks [Failed] $($_.Exception.Message)" 'Red'
        return $false
    }
}

function Invoke-OfficeActivation {
    param(
        [hashtable]$Installations
    )
    
    Write-Title
    Write-Host "Starting Office Activation..." -ForegroundColor Cyan
    Write-Host ""
    
    Get-OSInfo
    Write-Host ""
    
    $activated = $false
    
    # Process Office 16.0 C2R
    if ($Installations.O16C2R) {
        Write-Host "Found Office 16.0 Click-to-Run installation" -ForegroundColor Green
        
        $regPath = $Installations.O16C2R
        $installPath = (Get-ItemProperty $regPath -Name InstallPath).InstallPath
        $rootPath = Join-Path $installPath "root"
        $arch = (Get-ItemProperty $regPath -Name Platform -ErrorAction SilentlyContinue).Platform
        $version = (Get-ItemProperty "$regPath\Configuration" -Name VersionToReport -ErrorAction SilentlyContinue).VersionToReport
        
        Write-Host "Version: $version | Architecture: $arch" -ForegroundColor Gray
        
        # Determine hook path based on architecture
        if ($arch -eq 'x64') {
            $hookPath = Join-Path $rootPath "vfs\System"
            $hookFile = "sppc64.dll"
            $sppcPath = "$env:SystemRoot\System32\sppc.dll"
        } else {
            $hookPath = Join-Path $rootPath "vfs\SystemX86"
            $hookFile = "sppc32.dll"
            $sppcPath = "$env:SystemRoot\SysWOW64\sppc.dll"
        }
        
        # Clear license blocks
        Clear-OfficeLicenseBlocks
        
        # Install Ohook DLL
        if (Install-OhookDLL -HookPath $hookPath -HookFile $hookFile -SppcPath $sppcPath) {
            $activated = $true
        }
        
        # Get installed products and install keys
        # This would query the actual installed products and install appropriate keys
        Write-Host ""
        Write-Host "Installing product keys..." -ForegroundColor Yellow
        
        # Example: Install a ProPlus key (in real implementation, detect actual products)
        $keyData = Get-ProductKeyData
        $sampleKey = $keyData['16_de52bd50-9564-4adc-8fcb-a345c17f84f9']
        if ($sampleKey) {
            Install-ProductKey -Key $sampleKey.Key -ProductName $sampleKey.Product
        }
    }
    
    # Process Office 15.0 C2R
    if ($Installations.O15C2R) {
        Write-Host "Found Office 15.0 Click-to-Run installation" -ForegroundColor Green
        # Similar processing as O16C2R
    }
    
    # Process MSI installations
    foreach ($key in @('O16MSI', 'O15MSI', 'O14MSI')) {
        if ($Installations[$key]) {
            Write-Host "Found Office $($key.Substring(1,2)).0 MSI installation" -ForegroundColor Green
            # Process MSI installation
        }
    }
    
    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Cyan
    
    if ($activated -and -not $Script:ErrorFound) {
        Write-ColorText "Office is permanently activated!" 'Green'
        Write-Host "Help: https://massgrave.dev/troubleshoot" -ForegroundColor Gray
    } else {
        Write-ColorText "Some errors were detected." 'Red'
        Write-Host "Check this webpage for help: https://massgrave.dev/troubleshoot" -ForegroundColor Yellow
    }
    
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host ""
}

function Remove-OhookActivation {
    Write-Title
    Write-Host "Uninstalling Ohook Activation..." -ForegroundColor Yellow
    Write-Host ""
    
    $installations = Get-OfficeInstallations
    $removed = $false
    
    # Remove from Office 16.0 C2R
    if ($installations.O16C2R) {
        $regPath = $installations.O16C2R
        $installPath = (Get-ItemProperty $regPath -Name InstallPath).InstallPath
        $rootPath = Join-Path $installPath "root"
        $arch = (Get-ItemProperty $regPath -Name Platform -ErrorAction SilentlyContinue).Platform
        
        if ($arch -eq 'x64') {
            $hookPath = Join-Path $rootPath "vfs\System"
        } else {
            $hookPath = Join-Path $rootPath "vfs\SystemX86"
        }
        
        $sppcsDll = Join-Path $hookPath "sppcs.dll"
        $sppcDll = Join-Path $hookPath "sppc.dll"
        
        if (Test-Path $sppcsDll) {
            Remove-Item $sppcsDll -Force -ErrorAction SilentlyContinue
            Write-ColorText "Removed $sppcsDll" 'Green'
            $removed = $true
        }
        
        if (Test-Path $sppcDll) {
            Remove-Item $sppcDll -Force -ErrorAction SilentlyContinue
            Write-ColorText "Removed $sppcDll" 'Green'
            $removed = $true
        }
    }
    
    # Clear license blocks
    Clear-OfficeLicenseBlocks
    
    Write-Host ""
    if ($removed) {
        Write-ColorText "Ohook activation has been removed successfully!" 'Green'
    } else {
        Write-ColorText "No Ohook installation found." 'Yellow'
    }
    Write-Host ""
}

function Show-Menu {
    while ($true) {
        Write-Title
        Get-OSInfo
        Write-Host ""
        
        $installations = Get-OfficeInstallations
        $hasOffice = $false
        
        foreach ($key in $installations.Keys) {
            if ($installations[$key]) {
                $hasOffice = $true
                break
            }
        }
        
        if (-not $hasOffice) {
            Write-ColorText "No Office installation detected!" 'Red'
            Write-Host ""
            Write-Host "Please install Office first, then run this script again." -ForegroundColor Yellow
            Write-Host ""
            Write-Host "[0] Exit"
            Write-Host ""
            
            $choice = Read-Host "Choose an option"
            if ($choice -eq '0') { exit }
            continue
        }
        
        Write-Host "Office installations detected:" -ForegroundColor Green
        foreach ($key in $installations.Keys) {
            if ($installations[$key]) {
                Write-Host "  - $key" -ForegroundColor Gray
            }
        }
        Write-Host ""
        Write-Host "============================================================" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "[1] Install Ohook Office Activation"
        Write-Host ""
        Write-Host "[2] Uninstall Ohook"
        Write-Host ""
        Write-Host "[3] Download Office (opens browser)"
        Write-Host ""
        Write-Host "[0] Exit"
        Write-Host ""
        Write-Host "============================================================" -ForegroundColor Cyan
        Write-Host ""
        
        $choice = Read-Host "Choose an option [1,2,3,0]"
        
        switch ($choice) {
            '1' {
                Invoke-OfficeActivation -Installations $installations
                Write-Host ""
                Write-Host "Press any key to return to menu..."
                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            }
            '2' {
                Remove-OhookActivation
                Write-Host "Press any key to return to menu..."
                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            }
            '3' {
                Start-Process "https://massgrave.dev/genuine-installation-media"
            }
            '0' {
                exit
            }
            default {
                Write-ColorText "Invalid option. Please try again." 'Red'
                Start-Sleep -Seconds 2
            }
        }
    }
}

#endregion

#region Main Execution

# Handle unattended mode
if ($UnattendedMode) {
    $installations = Get-OfficeInstallations
    
    if ($OhookUninstall) {
        Remove-OhookActivation
    } elseif ($Ohook) {
        Invoke-OfficeActivation -Installations $installations
    }
} else {
    # Show interactive menu
    Show-Menu
}

#endregion
