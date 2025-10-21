# This script is hosted on GitHub for Ohook Activation AIO
# Method based on check3rd-av-ts pattern

if (-not $args) {
    Write-Host ''
    Write-Host 'Ohook Activation AIO - PowerShell Wrapper' -NoNewline
    Write-Host ''
    Write-Host ''
}

& {
    $psv = (Get-Host).Version.Major
    $troubleshoot = 'https://mass-grave.dev/troubleshoot'

    if ($ExecutionContext.SessionState.LanguageMode.value__ -ne 0) {
        $ExecutionContext.SessionState.LanguageMode
        Write-Host "PowerShell is not running in Full Language Mode."
        Write-Host "Help - $troubleshoot" -ForegroundColor White -BackgroundColor Blue
        return
    }

    try {
        [void][System.AppDomain]::CurrentDomain.GetAssemblies(); [void][System.Math]::Sqrt(144)
    }
    catch {
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Powershell failed to load .NET command."
        Write-Host "Help - $troubleshoot" -ForegroundColor White -BackgroundColor Blue
        return
    }

    function Check3rdAV {
        $cmd = if ($psv -ge 3) { 'Get-CimInstance' } else { 'Get-WmiObject' }
        $avList = & $cmd -Namespace root\SecurityCenter2 -Class AntiVirusProduct | Where-Object { $_.displayName -notlike '*windows*' } | Select-Object -ExpandProperty displayName

        if ($avList) {
            Write-Host '3rd party Antivirus might be blocking the script - ' -ForegroundColor White -BackgroundColor Blue -NoNewline
            Write-Host " $($avList -join ', ')" -ForegroundColor DarkRed -BackgroundColor White
        }
    }

    function CheckFile {
        param ([string]$FilePath)
        if (-not (Test-Path $FilePath)) {
            Check3rdAV
            Write-Host "Failed to find or create Ohook Activation file, aborting!"
            Write-Host "Help - $troubleshoot" -ForegroundColor White -BackgroundColor Blue
            throw
        }
    }

    try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    $cmdFilePath = Join-Path $scriptDir "Ohook_Activation_AIO.cmd"
    
    if (-not (Test-Path $cmdFilePath)) {
        Write-Host "Error: Ohook_Activation_AIO.cmd not found in script directory!"
        Write-Host "Please ensure the CMD file is in the same folder as this PowerShell script."
        Write-Host "Help - $troubleshoot" -ForegroundColor White -BackgroundColor Blue
        return
    }

    $rand = [Guid]::NewGuid().Guid
    $isAdmin = [bool]([Security.Principal.WindowsIdentity]::GetCurrent().Groups -match 'S-1-5-32-544')
    $tempPath = if ($isAdmin) { "$env:SystemRoot\Temp\Ohook_$rand.cmd" } else { "$env:USERPROFILE\AppData\Local\Temp\Ohook_$rand.cmd" }
    
    Copy-Item -Path $cmdFilePath -Destination $tempPath -Force
    CheckFile $tempPath

    $env:ComSpec = "$env:SystemRoot\system32\cmd.exe"
    $chkcmd = & $env:ComSpec /c "echo CMD is working"
    if ($chkcmd -notcontains "CMD is working") {
        Write-Warning "cmd.exe is not working.`nReport this issue at $troubleshoot"
    }

    if ($psv -lt 3) {
        if (Test-Path "$env:SystemRoot\Sysnative") {
            Write-Warning "Command is running with x86 Powershell, run it with x64 Powershell instead..."
            return
        }
        $p = & saps -FilePath $env:ComSpec -ArgumentList "/c `"$tempPath`"" -Verb RunAs -PassThru -WindowStyle Normal
        $p.WaitForExit()
    }
    else {
        & saps -FilePath $env:ComSpec -ArgumentList "/c `"$tempPath`"" -Verb RunAs -Wait -WindowStyle Normal
    }	
	
    CheckFile $tempPath

    $tempPatterns = @("$env:SystemRoot\Temp\Ohook*.cmd", "$env:USERPROFILE\AppData\Local\Temp\Ohook*.cmd")
    foreach ($pattern in $tempPatterns) { 
        Get-Item $pattern -ErrorAction SilentlyContinue | Remove-Item -ErrorAction SilentlyContinue
    }
} @args
