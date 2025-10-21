@set masver=3.7
@echo off



::============================================================================
::
::   Homepage: mass()grave(dot)dev
::      Email: mas.help@outlook.com
::
::============================================================================



::  To activate Office with Ohook activation, run the script with "/Ohook" parameter or change 0 to 1 in below line
set _act=0

::  To remove Ohook activation, run the script with /Ohook-Uninstall parameter or change 0 to 1 in below line
set _rem=0

::  To run the script in debug mode, change 0 to "/Ohook" in below line
set "_debug=0"

::  If value is changed in above lines or parameter is used then script will run in unattended mode



::========================================================================================================================================

::  Set environment variables, it helps if they are misconfigured in the system

setlocal EnableExtensions
setlocal DisableDelayedExpansion

set "PathExt=.COM;.EXE;.BAT;.CMD;.VBS;.VBE;.JS;.JSE;.WSF;.WSH;.MSC"

set "SysPath=%SystemRoot%\System32"
set "Path=%SystemRoot%\System32;%SystemRoot%;%SystemRoot%\System32\Wbem;%SystemRoot%\System32\WindowsPowerShell\v1.0\"
if exist "%SystemRoot%\Sysnative\reg.exe" (
set "SysPath=%SystemRoot%\Sysnative"
set "Path=%SystemRoot%\Sysnative;%SystemRoot%;%SystemRoot%\Sysnative\Wbem;%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\;%Path%"
)

set "ComSpec=%SysPath%\cmd.exe"
set "PSModulePath=%ProgramFiles%\WindowsPowerShell\Modules;%SysPath%\WindowsPowerShell\v1.0\Modules"

set re1=
set re2=
set "_cmdf=%~f0"
for %%# in (%*) do (
if /i "%%#"=="re1" set re1=1
if /i "%%#"=="re2" set re2=1
if /i "%%#"=="-qedit" (set re1=1&set re2=1)
)

:: Re-launch the script with x64 process if it was initiated by x86 process on x64 bit Windows
:: or with ARM64 process if it was initiated by x86/ARM32 process on ARM64 Windows

if exist %SystemRoot%\Sysnative\cmd.exe if not defined re1 (
setlocal EnableDelayedExpansion
start %SystemRoot%\Sysnative\cmd.exe /c ""!_cmdf!" %* re1"
exit /b
)

:: Re-launch the script with ARM32 process if it was initiated by x64 process on ARM64 Windows

if exist %SystemRoot%\SysArm32\cmd.exe if %PROCESSOR_ARCHITECTURE%==AMD64 if not defined re2 (
setlocal EnableDelayedExpansion
start %SystemRoot%\SysArm32\cmd.exe /c ""!_cmdf!" %* re2"
exit /b
)

::========================================================================================================================================

::  Debug code

if "%_debug%" EQU "0" (
set "nul1=1>nul"
set "nul2=2>nul"
set "nul6=2^>nul"
set "nul=>nul 2>&1"
goto :_debug
)

set "nul1="
set "nul2="
set "nul6="
set "nul="

@echo on
@prompt $G
@call :_debug "%_debug%" >"%~dp0_tmp.log" 2>&1
@cmd /u /c type "%~dp0_tmp.log">"%~dp0_Debug.log"
@del "%~dp0_tmp.log"
@echo off
@exit /b

:_debug

::========================================================================================================================================

set "blank="
set "mas=ht%blank%tps%blank%://mass%blank%grave.dev/"
set "github=ht%blank%tps%blank%://github.com/massgra%blank%vel/Micro%blank%soft-Acti%blank%vation-Scripts"
set "selfgit=ht%blank%tps%blank%://git.acti%blank%vated.win/massg%blank%rave/Micr%blank%osoft-Act%blank%ivation-Scripts"

::  Check if Null service is working, it's important for the batch script

sc query Null | find /i "RUNNING"
if %errorlevel% NEQ 0 (
echo:
echo Null service is not running, script may crash...
echo:
echo:
echo Check this webpage for help - %mas%fix_service
echo:
echo:
ping 127.0.0.1 -n 20
)
cls

::  Check LF line ending

pushd "%~dp0"
>nul findstr /v "$" "%~nx0" && (
echo:
echo Error - Script either has LF line ending issue or an empty line at the end of the script is missing.
echo:
echo:
echo Check this webpage for help - %mas%troubleshoot
echo:
echo:
ping 127.0.0.1 -n 20 >nul
popd
exit /b
)
popd

::========================================================================================================================================

cls
color 07
title  Ohook Activation %masver%

set _args=
set _elev=
set _unattended=0

set _args=%*
if defined _args set _args=%_args:"=%
if defined _args set _args=%_args:re1=%
if defined _args set _args=%_args:re2=%
if defined _args (
for %%A in (%_args%) do (
if /i "%%A"=="/Ohook"                  set _act=1
if /i "%%A"=="/Ohook-Uninstall"        set _rem=1
if /i "%%A"=="-el"                     set _elev=1
)
)

for %%A in (%_act% %_rem%) do (if "%%A"=="1" set _unattended=1)

::========================================================================================================================================

call :dk_setvar

if %winbuild% EQU 1 (
%eline%
echo Failed to detect Windows build number.
echo:
setlocal EnableDelayedExpansion
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Check this webpage for help - " %_Yellow% " %mas%troubleshoot"
goto dk_done
)

if exist "%Systemdrive%\Users\WDAGUtilityAccount" (
sc query gcs | find /i "RUNNING" %nul% && (
%eline%
echo Windows Sandbox detected; activation is not supported.
echo The script cannot run due to missing licensing components. Aborting...
echo:
goto dk_done
)
)

if %winbuild% LSS 6001 (
%nceline%
echo Unsupported OS version detected [%winbuild%].
echo MAS only supports Windows Vista/7/8/8.1/10/11 and their Server equivalents.
if %winbuild% EQU 6000 (
echo:
echo Windows Vista RTM is not supported because Powershell cannot be installed.
echo Upgrade to Windows Vista SP1 or SP2.
)
goto dk_done
)

if %winbuild% LSS 7600 if not exist "%SysPath%\WindowsPowerShell\v1.0\Modules" (
%nceline%
if not exist %ps% (
echo PowerShell is not installed in your system.
)
echo Install PowerShell 2.0 using the following URL.
echo:
echo https://www.catalog.update.microsoft.com/Search.aspx?q=KB968930
if %_unattended%==0 start https://www.catalog.update.microsoft.com/Search.aspx?q=KB968930
goto dk_done
)

::========================================================================================================================================

::  Fix special character limitations in path name

set "_work=%~dp0"
if "%_work:~-1%"=="\" set "_work=%_work:~0,-1%"

set "_batf=%~f0"
set "_batp=%_batf:'=''%"

set _PSarg="""%~f0""" -el %_args%
set _PSarg=%_PSarg:'=''%

set "_ttemp=%userprofile%\AppData\Local\Temp"

setlocal EnableDelayedExpansion

::========================================================================================================================================

echo "!_batf!" | find /i "!_ttemp!" %nul1% && (
if /i not "!_work!"=="!_ttemp!" (
%eline%
echo The script was launched from the temp folder.
echo You are most likely running the script directly from the archive file.
echo:
echo Extract the archive file and launch the script from the extracted folder.
goto dk_done
)
)

::========================================================================================================================================

::  Elevate script as admin and pass arguments and preventing loop

%nul1% fltmc || (
if not defined _elev %psc% "start cmd.exe -arg '/c \"!_PSarg!\"' -verb runas" && exit /b
%eline%
echo This script needs admin rights.
echo Right click on this script and select 'Run as administrator'.
goto dk_done
)

::========================================================================================================================================

::  Check PowerShell

::pstst $ExecutionContext.SessionState.LanguageMode :pstst

for /f "delims=" %%a in ('%psc% "if ($PSVersionTable.PSEdition -ne 'Core') {$f=[System.IO.File]::ReadAllText('!_batp!') -split ':pstst';. ([scriptblock]::Create($f[1]))}" %nul6%') do (set tstresult=%%a)

if /i not "%tstresult%"=="FullLanguage" (
%eline%
for /f "delims=" %%a in ('%psc% "$ExecutionContext.SessionState.LanguageMode" %nul6%') do (set tstresult2=%%a)
echo Test 1 - %tstresult%
echo Test 2 - !tstresult2!
echo:

REM check LanguageMode

echo: !tstresult2! | findstr /i "ConstrainedLanguage RestrictedLanguage NoLanguage" %nul1% && (
echo FullLanguage mode not found in PowerShell. Aborting...
echo If you have applied restrictions on Powershell then undo those changes.
echo:
set fixes=%fixes% %mas%fix_powershell
call :dk_color2 %Blue% "Check this webpage for help - " %_Yellow% " %mas%fix_powershell"
goto dk_done
)

REM check Powershell core version

cmd /c "%psc% "$PSVersionTable.PSEdition"" | find /i "Core" %nul1% && (
echo Windows Powershell is needed for MAS but it seems to be replaced with Powershell core. Aborting...
echo:
set fixes=%fixes% %mas%in-place_repair_upgrade
call :dk_color2 %Blue% "Check this webpage for help - " %_Yellow% " %mas%in-place_repair_upgrade"
goto dk_done
)

REM check for Mal-ware that may cause issues with Powershell

for /r "%ProgramFiles%\" %%f in (secureboot.exe) do if exist "%%f" (
echo "%%f"
echo Mal%blank%ware found, PowerShell is not working properly.
echo:
set fixes=%fixes% %mas%remove_mal%w%ware
call :dk_color2 %Blue% "Check this webpage for help - " %_Yellow% " %mas%remove_mal%w%ware"
goto dk_done
)

REM check if .NET is working properly

if /i "!tstresult2!"=="FullLanguage" (
cmd /c "%psc% ""try {[System.AppDomain]::CurrentDomain.GetAssemblies(); [System.Math]::Sqrt(144)} catch {Exit 3}""" %nul%
if !errorlevel!==3 (
echo Windows Powershell failed to load .NET command. Aborting...
echo:
set fixes=%fixes% %mas%in-place_repair_upgrade
call :dk_color2 %Blue% "Check this webpage for help - " %_Yellow% " %mas%in-place_repair_upgrade"
goto dk_done
)
)

REM check antivirus and other errors

echo PowerShell is not working properly. Aborting...

if /i "!tstresult2!"=="FullLanguage" (
echo:
echo Your antivirus software might be blocking the script.
echo:
sc query sense | find /i "RUNNING" %nul% && (
echo Installed Antivirus - Microsoft Defender for Endpoint
)
cmd /c "%psc% ""$av = Get-WmiObject -Namespace root\SecurityCenter2 -Class AntiVirusProduct; $n = @(); foreach ($i in $av) { $n += $i.displayName }; if ($n) { Write-Host ('Installed Antivirus - ' + ($n -join ', '))}"""
)

echo:
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Check this webpage for help - " %_Yellow% " %mas%troubleshoot"
goto dk_done
)

::========================================================================================================================================

::  Disable QuickEdit and launch from conhost.exe to avoid Terminal app

if %winbuild% GEQ 17763 (
set terminal=1
) else (
set terminal=
)

::  Check if script is running in Terminal app

if defined terminal (
set lines=0
for /f "skip=3 tokens=* delims=" %%A in ('mode con') do if "!lines!"=="0" (
for %%B in (%%A) do set lines=%%B
)
if !lines! GEQ 100 set terminal=
)

if %_unattended%==1 goto :skipQE
for %%# in (%_args%) do (if /i "%%#"=="-qedit" goto :skipQE)

::  Relaunch to disable QuickEdit in the current session and use conhost.exe instead of the Terminal app
::  This code disables QuickEdit for the current cmd.exe session without making permanent registry changes
::  It is included because clicking on the script window can pause execution, causing confusion that the script has stopped due to an error

set resetQE=1
reg query HKCU\Console /v QuickEdit %nul2% | find /i "0x0" %nul1% && set resetQE=0
reg add HKCU\Console /v QuickEdit /t REG_DWORD /d 0 /f %nul1%

if defined terminal (
start conhost.exe "!_batf!" %_args% -qedit
start reg add HKCU\Console /v QuickEdit /t REG_DWORD /d %resetQE% /f %nul1%
exit /b
) else if %resetQE% EQU 1 (
start cmd.exe /c ""!_batf!" %_args% -qedit"
start reg add HKCU\Console /v QuickEdit /t REG_DWORD /d %resetQE% /f %nul1%
exit /b
)

:skipQE

::========================================================================================================================================

::  Check for updates

set -=
set old=
set pingp=
set upver=%masver:.=%

for %%A in (
activ%-%ated.win
mass%-%grave.dev
) do if not defined pingp (
for /f "delims=[] tokens=2" %%B in ('ping -n 1 %%A') do (
if not "%%B"=="" (set old=1& set pingp=1)
for /f "delims=[] tokens=2" %%C in ('ping -n 1 updatecheck%upver%.%%A') do (
if not "%%C"=="" set old=
)
)
)

if defined old (
echo ________________________________________________
%eline%
echo Your version of MAS [%masver%] is outdated.
echo ________________________________________________
echo:
if not %_unattended%==1 (
echo [1] Get Latest MAS
echo [0] Continue Anyway
echo:
call :dk_color %_Green% "Choose a menu option using your keyboard [1,0] :"
choice /C:10 /N
if !errorlevel!==2 rem
if !errorlevel!==1 (start %selfgit% & start %github% & start %mas% & exit /b)
)
)
cls

::========================================================================================================================================

if %_rem%==1 goto :oh_uninstall

:oh_menu

if %_unattended%==0 (
cls
if not defined terminal mode 76, 25
title  Ohook Activation %masver%
call :oh_checkapps
echo:
echo:
echo:
echo:
if defined checknames (call :dk_color %_Yellow% "                Close [!checknames!] before proceeding...")
echo         ____________________________________________________________
echo:
echo                 [1] Install Ohook Office Activation
echo:
echo                 [2] Uninstall Ohook
echo                 ____________________________________________
echo:
echo                 [3] Download Office
echo:
echo                 [0] %_exitmsg%
echo         ____________________________________________________________
echo:
call :dk_color2 %_White% "             " %_Green% "Choose a menu option using your keyboard [1,2,3,0]"
choice /C:1230 /N
set _el=!errorlevel!
if !_el!==4  exit /b
if !_el!==3  start %mas%genuine-installation-media &goto :oh_menu
if !_el!==2  goto :oh_uninstall
if !_el!==1  goto :oh_menu2
goto :oh_menu
)

::========================================================================================================================================

:oh_menu2

cls
if not defined terminal (
mode 140, 32
if exist "%SysPath%\spp\store_test\" mode 140, 32
%psc% "&{$W=$Host.UI.RawUI.WindowSize;$B=$Host.UI.RawUI.BufferSize;$W.Height=32;$B.Height=300;$Host.UI.RawUI.WindowSize=$W;$Host.UI.RawUI.BufferSize=$B;}" %nul%
)
title  Ohook Activation %masver%

echo:
echo Initializing...
call :dk_chkmal

if not exist %SysPath%\%_slexe% (
%eline%
echo [%SysPath%\%_slexe%] file is missing, aborting...
echo:
if not defined results (
call :dk_color %Blue% "Go back to Main Menu, select Troubleshoot and run DISM Restore and SFC Scan options."
call :dk_color %Blue% "After that, restart system and try activation again."
echo:
set fixes=%fixes% %mas%troubleshoot
call :dk_color2 %Blue% "Check this webpage for help - " %_Yellow% " %mas%troubleshoot"
)
goto dk_done
)

::========================================================================================================================================

set spp=SoftwareLicensingProduct
set sps=SoftwareLicensingService

call :dk_reflection
call :dk_ckeckwmic
call :dk_product
call :dk_sppissue

::========================================================================================================================================

set error=

cls
echo:
call :dk_showosinfo

::========================================================================================================================================

echo Initiating Diagnostic Tests...

set "_serv=%_slser% Winmgmt"

::  Software Protection
::  Windows Management Instrumentation

set notwinact=1
set ohookact=1
call :dk_errorcheck

call :oh_setspp

::  Check unsupported office versions

set o14c2r=
set o16uwp=

set _68=HKLM\SOFTWARE\Microsoft\Office
set _86=HKLM\SOFTWARE\Wow6432Node\Microsoft\Office
%nul% reg query %_68%\14.0\CVH /f Click2run /k         && set o14c2r=Office 2010 C2R
%nul% reg query %_86%\14.0\CVH /f Click2run /k         && set o14c2r=Office 2010 C2R

if %winbuild% GEQ 10240 (
for /f "delims=" %%a in ('%psc% "(Get-AppxPackage -name 'Microsoft.Office.Desktop' | Select-Object -ExpandProperty InstallLocation)" %nul6%') do (if exist "%%a\Integration\Integrator.exe" set o16uwp=Office UWP )
)

if not "%o14c2r%%o16uwp%"=="" (
echo:
call :dk_color %Red% "Checking Unsupported Office Install     [ %o14c2r%%o16uwp%]"
if not "%o16uwp%"=="" call :dk_color %Blue% "Use TSforge option to activate it."
)

if %winbuild% GEQ 10240 %psc% "Get-AppxPackage -name "Microsoft.MicrosoftOfficeHub"" | find /i "Office" %nul1% && (
set ohub=1
)

::========================================================================================================================================

::  Check supported office versions

call :oh_getpath

sc query ClickToRunSvc %nul%
set error1=%errorlevel%

if defined o16c2r if %error1% EQU 1060 (
call :dk_color %Red% "Checking ClickToRun Service             [Not found, Office 16.0 files found]"
set o16c2r=
set error=1
)

sc query OfficeSvc %nul%
set error2=%errorlevel%

if defined o15c2r if %error1% EQU 1060 if %error2% EQU 1060 (
call :dk_color %Red% "Checking ClickToRun Service             [Not found, Office 15.0 files found]"
set o15c2r=
set error=1
)

if "%o16c2r%%o15c2r%%o16msi%%o15msi%%o14msi%"=="" (
set error=1
echo:
if not "%o14c2r%%o16uwp%"=="" (
call :dk_color %Red% "Checking Supported Office Install       [Not Found]"
) else (
call :dk_color %Red% "Checking Installed Office               [Not Found]"
)

if defined ohub (
echo:
echo You only have the Office Dashboard app installed. You need to install the full version of Office.
)
echo:
call :dk_color %Blue% "Download and install Office from the below URL and then try again."
echo:
set fixes=%fixes% %mas%genuine-installation-media
call :dk_color %_Yellow% "%mas%genuine-installation-media"
goto dk_done
)

set multioffice=
if not "%o16c2r%%o15c2r%%o16msi%%o15msi%%o14msi%"=="1" set multioffice=1
if not "%o14c2r%%o16uwp%"=="" set multioffice=1

if defined multioffice (
call :dk_color %Gray% "Checking Multiple Office Install        [Found, its recommended to install only one version]"
)

::========================================================================================================================================

::  Check Windows Server

set winserver=
reg query "HKLM\SYSTEM\CurrentControlSet\Control\ProductOptions" /v ProductType %nul2% | find /i "WinNT" %nul1% || set winserver=1
if not defined winserver (
reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID %nul2% | find /i "Server" %nul1% && set winserver=1
)

::========================================================================================================================================

::  Process Office 15.0 C2R

if not defined o15c2r goto :starto16c2r

call :oh_reset
call :dk_actids 0ff1ce15-a989-479d-af46-f275c6370663

set oVer=15
for /f "skip=2 tokens=2*" %%a in ('"reg query %o15c2r_reg% /v InstallPath" %nul6%') do (set "_oRoot=%%b\root")
for /f "skip=2 tokens=2*" %%a in ('"reg query %o15c2r_reg%\Configuration /v Platform" %nul6%') do (set "_oArch=%%b")
for /f "skip=2 tokens=2*" %%a in ('"reg query %o15c2r_reg%\Configuration /v VersionToReport" %nul6%') do (set "_version=%%b")
for /f "skip=2 tokens=2*" %%a in ('"reg query %o15c2r_reg%\Configuration /v ProductReleaseIds" %nul6%') do (set "_prids=%o15c2r_reg%\Configuration /v ProductReleaseIds" & set "_config=%o15c2r_reg%\Configuration")
if not defined _oArch   for /f "skip=2 tokens=2*" %%a in ('"reg query %o15c2r_reg%\propertyBag /v Platform" %nul6%') do (set "_oArch=%%b")
if not defined _version for /f "skip=2 tokens=2*" %%a in ('"reg query %o15c2r_reg%\propertyBag /v version" %nul6%') do (set "_version=%%b")
if not defined _prids   for /f "skip=2 tokens=2*" %%a in ('"reg query %o15c2r_reg%\propertyBag /v ProductReleaseId" %nul6%') do (
