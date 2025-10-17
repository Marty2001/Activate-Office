# Ohook Activation Wrapper

A secure PowerShell wrapper for executing Ohook_Activation_AIO.cmd with integrity verification and automatic cleanup.

## Features

- ✅ SHA256 hash verification for security
- ✅ Multiple mirror URLs for reliability
- ✅ Automatic admin privilege elevation
- ✅ Antivirus detection warnings
- ✅ Automatic cleanup after execution
- ✅ Support for both PowerShell v2 and v3+

## Usage

### Method 1: Direct Execution (Recommended)

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -Command "& { $(Invoke-RestMethod 'https://raw.githubusercontent.com/Marty2001/ohook-activation/main/Ohook-Activation-Wrapper.ps1') }"
