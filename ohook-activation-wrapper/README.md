# Ohook Activation AIO

A PowerShell wrapper script that downloads and executes the Ohook Activation AIO CMD script with enhanced security checks and white background console styling.

## Features

- **Secure Download**: Downloads from GitHub with multiple fallback URLs
- **Integrity Verification**: SHA256 hash verification of downloaded scripts
- **3rd Party AV Detection**: Checks for third-party antivirus that might block execution
- **AutoRun Registry Check**: Detects problematic AutoRun registry entries
- **White Console Theme**: All console popups use white background with black text
- **Admin Privilege Check**: Automatically requests admin privileges when needed

## Usage

### Method 1: Direct PowerShell Execution
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -Command "& { iwr -useb 'https://raw.githubusercontent.com/yourusername/ohook-activation-aio/main/Ohook_Activation_AIO.ps1' | iex }"
