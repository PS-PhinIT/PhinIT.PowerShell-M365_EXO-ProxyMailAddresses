# Installation Guide for easyEXCH ProxyAddress Manager

## Prerequisites Installation

### 1. Install RSAT Tools (Active Directory PowerShell Module)

#### Windows 10/11 (Version 1809 or later)

Open PowerShell as Administrator and run:

```powershell
# Check if already installed
Get-WindowsCapability -Name RSAT.ActiveDirectory* -Online

# Install RSAT-AD-PowerShell
Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0

# Verify installation
Import-Module ActiveDirectory
Get-Module ActiveDirectory
```

#### Windows 10 (Older versions)

1. Download RSAT from Microsoft Download Center
2. Install RSAT package
3. Enable Active Directory module in Windows Features

#### Windows Server 2016/2019/2022

```powershell
# Check if already installed
Get-WindowsFeature -Name RSAT-AD-PowerShell

# Install RSAT-AD-PowerShell
Install-WindowsFeature -Name RSAT-AD-PowerShell -IncludeAllSubFeature

# Verify installation
Import-Module ActiveDirectory
Get-Module ActiveDirectory
```

### 2. Check PowerShell Version

```powershell
# Check PowerShell version (requires 5.1 or higher)
$PSVersionTable.PSVersion

# Output should show Major version 5 or higher
```

### 3. Check .NET Framework Version

```powershell
# Check .NET Framework version (requires 4.5 or higher)
Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" | Select-Object Release, Version

# Release numbers:
# 378389 = .NET 4.5
# 378675/378758 = .NET 4.5.1
# 379893 = .NET 4.5.2
# 393295/393297 = .NET 4.6
# 394254/394271 = .NET 4.6.1
# 394802/394806 = .NET 4.6.2
# 460798/460805 = .NET 4.7
# 461308/461310 = .NET 4.7.1
# 461808/461814 = .NET 4.7.2
# 528040/528049 = .NET 4.8
```

## Application Installation

### Method 1: Git Clone

```powershell
# Clone repository
git clone https://github.com/PS-easyIT/easyEXCH-ProxyMailAddresses.git

# Navigate to folder
cd easyEXCH-ProxyMailAddresses

# Run application
.\easyEXCH-ProxyMailAddresses.ps1
```

### Method 2: Direct Download

1. Download ZIP file from GitHub
2. Extract to desired location
3. Right-click on file → Properties → Unblock (if necessary)
4. Open PowerShell in extracted folder
5. Run: `.\easyEXCH-ProxyMailAddresses.ps1`

## PowerShell Execution Policy

If you encounter execution policy errors:

```powershell
# Check current policy
Get-ExecutionPolicy -List

# Set policy for current user (recommended)
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

# Alternative: Set policy for current process only (temporary)
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process

# Verify new policy
Get-ExecutionPolicy
```

## Active Directory Permissions

The user running the script needs:

### Minimum Required Permissions:
- Read permissions on User objects in AD
- Write permissions on ProxyAddresses attribute

### How to Verify Permissions:

```powershell
# Test AD connectivity
Import-Module ActiveDirectory

# Test user search
Get-ADUser -Filter {SamAccountName -eq "testuser"} -Properties ProxyAddresses

# Test if you can read ProxyAddresses
$user = Get-ADUser -Filter {SamAccountName -eq "testuser"} -Properties ProxyAddresses
$user.ProxyAddresses
```

### Granting Permissions (Domain Admin):

If you need to grant specific permissions:

1. Open "Active Directory Users and Computers"
2. View → Advanced Features
3. Right-click on OU containing users → Properties
4. Security tab → Advanced
5. Add → Select principal
6. Select user/group that needs access
7. Permissions:
   - Read all properties
   - Write proxyAddresses

## Troubleshooting Installation

### Issue: "Cannot load ActiveDirectory module"

**Solution:**
```powershell
# Verify RSAT is installed
Get-WindowsCapability -Name RSAT.ActiveDirectory* -Online | Where-Object State -eq "Installed"

# Try manual import
Import-Module ActiveDirectory -Verbose

# Check module location
Get-Module -ListAvailable ActiveDirectory
```

### Issue: "Script is not digitally signed"

**Solution:**
```powershell
# Option 1: Change execution policy
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

# Option 2: Unblock file
Unblock-File -Path ".\easyEXCH-ProxyMailAddresses.ps1"

# Option 3: Run with bypass (one-time)
PowerShell.exe -ExecutionPolicy Bypass -File ".\easyEXCH-ProxyMailAddresses.ps1"
```

### Issue: "WPF window not showing"

**Solution:**
```powershell
# Verify .NET Framework
[System.Reflection.Assembly]::GetEntryAssembly().ImageRuntimeVersion

# Check WPF assemblies
[System.Reflection.Assembly]::LoadWithPartialName("PresentationFramework")
[System.Reflection.Assembly]::LoadWithPartialName("PresentationCore")

# Try running as different user with admin rights
```

### Issue: "Access Denied" when modifying ProxyAddresses

**Solution:**
- Verify you have write permissions on ProxyAddresses attribute
- Contact Domain Administrator for necessary permissions
- Test with a user account you own

## Post-Installation Verification

```powershell
# Test script
.\easyEXCH-ProxyMailAddresses.ps1

# Application should:
# 1. Load without errors
# 2. Display WPF window
# 3. Show "Ready" in status bar
# 4. Allow user search (test with known user)
```

## Creating Desktop Shortcut (Optional)

1. Right-click on Desktop → New → Shortcut
2. Location: `powershell.exe -ExecutionPolicy Bypass -File "C:\Path\To\easyEXCH-ProxyMailAddresses.ps1"`
3. Name: "easyEXCH ProxyAddress Manager"
4. Click Finish
5. Right-click shortcut → Properties
6. Change icon (optional)
7. Run as administrator (if needed)

## Network Requirements

- Access to Domain Controllers
- DNS resolution working
- Firewall allows LDAP (389/tcp, 636/tcp)
- Kerberos authentication (88/tcp, 88/udp)

## Support

If you continue to have installation issues:
1. Check prerequisites are met
2. Review error messages
3. Open issue on GitHub with details
4. Include PowerShell version, Windows version, and error messages
