# Modified by: BitCourse https://web.facebook.com/DigitalNecessitiesBitCourse
# Usage: irm https://raw.githubusercontent.com/Marty2001/Activate-Office/refs/heads/main/BitCourse-Office-Activation.ps1 | iex

$host.UI.RawUI.BackgroundColor = "White"
$host.UI.RawUI.ForegroundColor = "Black"
Clear-Host

Write-Host ''
Write-Host 'BitCourse-MS Office Activation https://web.facebook.com/DigitalNecessitiesBitCourse' -ForegroundColor DarkYellow
Write-Host ''

& {
    $psv = (Get-Host).Version.Major
    $troubleshoot = 'https://github.com/[user]/[repo]/issues'

  

    if ($ExecutionContext.SessionState.LanguageMode.value__ -ne 0) {
        $ExecutionContext.SessionState.LanguageMode
        Write-Host "PowerShell is not running in Full Language Mode." -ForegroundColor Red
        return
    }

    try {
        [void][System.AppDomain]::CurrentDomain.GetAssemblies(); [void][System.Math]::Sqrt(144)
    }
    catch {
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "PowerShell failed to load .NET command." -ForegroundColor Red
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
            Write-Host "Failed to create BitCourse file in temp folder, aborting!"
            throw
        }
    }

    try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

    $URLs = @(
        'https://raw.githubusercontent.com/Marty2001/Activate-Office/refs/heads/main/BitCourse_Office_Activation.cmd'
    )
    
    Write-Progress -Activity "Downloading..." -Status "Please wait"
    $errors = @()
    foreach ($URL in $URLs | Sort-Object { Get-Random }) {
        try {
            if ($psv -ge 3) {
                $response = Invoke-RestMethod $URL
            }
            else {
                $w = New-Object Net.WebClient
                $response = $w.DownloadString($URL)
            }
            break
        }
        catch {
            $errors += $_
        }
    }
    Write-Progress -Activity "Downloading..." -Status "Done" -Completed

    if (-not $response) {
        Check3rdAV
        foreach ($err in $errors) {
            Write-Host "Error: $($err.Exception.Message)" -ForegroundColor Red
        }
        Write-Host "Failed to retrieve BitCourse script from any of the available repositories, aborting!"
        Write-Host "Check if antivirus or firewall is blocking the connection."
        return
    }

    # Check for AutoRun registry which may create issues with CMD
    $paths = "HKCU:\SOFTWARE\Microsoft\Command Processor", "HKLM:\SOFTWARE\Microsoft\Command Processor"
    foreach ($path in $paths) { 
        if (Get-ItemProperty -Path $path -Name "Autorun" -ErrorAction SilentlyContinue) { 
            Write-Warning "Autorun registry found, CMD may crash! `nManually copy-paste the below command to fix...`nRemove-ItemProperty -Path '$path' -Name 'Autorun'"
        } 
    }

    $rand = [Guid]::NewGuid().Guid
    $isAdmin = [bool]([Security.Principal.WindowsIdentity]::GetCurrent().Groups -match 'S-1-5-32-544')
    $FilePath = if ($isAdmin) { "$env:SystemRoot\Temp\BitCourse_$rand.cmd" } else { "$env:USERPROFILE\AppData\Local\Temp\BitCourse_$rand.cmd" }
    Set-Content -Path $FilePath -Value "@::: $rand `r`n$response"
    CheckFile $FilePath

    $env:ComSpec = "$env:SystemRoot\system32\cmd.exe"
    $chkcmd = & $env:ComSpec /c "echo CMD is working"
    if ($chkcmd -notcontains "CMD is working") {
        Write-Warning "cmd.exe is not working."
    }

    if ($psv -lt 3) {
        if (Test-Path "$env:SystemRoot\Sysnative") {
            Write-Warning "Command is running with x86 PowerShell, run it with x64 PowerShell instead..."
            return
        }
        $p = saps -FilePath $env:ComSpec -ArgumentList "/c """"$FilePath"" $args""" -Verb RunAs -PassThru
        $p.WaitForExit()
    }
    else {
        saps -FilePath $env:ComSpec -ArgumentList "/c """"$FilePath"" $args""" -Wait -Verb RunAs
    }	
    CheckFile $FilePath

    $FilePaths = @("$env:SystemRoot\Temp\BitCourse*.cmd", "$env:USERPROFILE\AppData\Local\Temp\BitCourse*.cmd")
    foreach ($FilePath in $FilePaths) { Get-Item $FilePath -ErrorAction SilentlyContinue | Remove-Item }
} @args
