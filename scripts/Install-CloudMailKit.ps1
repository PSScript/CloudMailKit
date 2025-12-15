<#
.SYNOPSIS
    Installs CloudMailKit for legacy applications (VB6, Classic ASP, old .NET)

.DESCRIPTION
    Installs and registers CloudMailKit DLL for COM interop on Windows servers.
    Designed for rescuing stone age applications when SMTP auth is disabled.

.PARAMETER InstallPath
    Where to install CloudMailKit (default: C:\CloudMailKit)

.PARAMETER RegisterCOM
    Register for COM interop (required for VB6/Classic ASP)

.PARAMETER InstallToGAC
    Install to Global Assembly Cache (for enterprise deployments)

.EXAMPLE
    .\Install-CloudMailKit.ps1
    Simple installation with COM registration

.EXAMPLE
    .\Install-CloudMailKit.ps1 -InstallPath "D:\Apps\CloudMailKit" -InstallToGAC
    Custom location + GAC installation
#>

[CmdletBinding()]
param(
    [string]$InstallPath = "C:\CloudMailKit",
    [switch]$RegisterCOM = $true,
    [switch]$InstallToGAC = $false
)

# Ensure running as Administrator
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Error "This script must be run as Administrator!"
    exit 1
}

Write-Host "CloudMailKit Installation Script" -ForegroundColor Cyan
Write-Host "=================================" -ForegroundColor Cyan
Write-Host ""

# Check .NET Framework version
Write-Host "Checking .NET Framework..." -ForegroundColor Yellow
$dotNetVersion = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" -ErrorAction SilentlyContinue).Release

if ($dotNetVersion -ge 528040) {
    Write-Host "✓ .NET Framework 4.8+ detected" -ForegroundColor Green
} elseif ($dotNetVersion -ge 394802) {
    Write-Host "✓ .NET Framework 4.6.2+ detected (minimum supported)" -ForegroundColor Green
} else {
    Write-Warning ".NET Framework 4.8 recommended. Download from: https://dotnet.microsoft.com/download/dotnet-framework/net48"
    $continue = Read-Host "Continue anyway? (y/n)"
    if ($continue -ne 'y') { exit 0 }
}

# Create installation directory
Write-Host ""
Write-Host "Creating installation directory..." -ForegroundColor Yellow
if (-not (Test-Path $InstallPath)) {
    New-Item -ItemType Directory -Path $InstallPath -Force | Out-Null
    Write-Host "✓ Created: $InstallPath" -ForegroundColor Green
} else {
    Write-Host "✓ Directory exists: $InstallPath" -ForegroundColor Green
}

# Find CloudMailKit DLL
Write-Host ""
Write-Host "Locating CloudMailKit.dll..." -ForegroundColor Yellow

$dllPath = $null
$searchPaths = @(
    ".\bin\Release\net48\CloudMailKit.dll",
    ".\bin\Debug\net48\CloudMailKit.dll",
    "..\bin\Release\net48\CloudMailKit.dll",
    ".\src\CloudMailKit\bin\Release\net48\CloudMailKit.dll",
    ".\src\CloudMailKit\bin\Debug\net48\CloudMailKit.dll"
)

foreach ($path in $searchPaths) {
    if (Test-Path $path) {
        $dllPath = Resolve-Path $path
        break
    }
}

if (-not $dllPath) {
    Write-Error "CloudMailKit.dll not found! Build the project first."
    Write-Host "Run: dotnet build -c Release" -ForegroundColor Yellow
    exit 1
}

Write-Host "✓ Found: $dllPath" -ForegroundColor Green

# Copy DLL to installation directory
Write-Host ""
Write-Host "Copying files..." -ForegroundColor Yellow
$destDll = Join-Path $InstallPath "CloudMailKit.dll"
Copy-Item $dllPath $destDll -Force
Write-Host "✓ Copied to: $destDll" -ForegroundColor Green

# Copy dependencies
$sourcePath = Split-Path $dllPath -Parent
$dependencies = @("System.Configuration.ConfigurationManager.dll")
foreach ($dep in $dependencies) {
    $depPath = Join-Path $sourcePath $dep
    if (Test-Path $depPath) {
        Copy-Item $depPath $InstallPath -Force
        Write-Host "✓ Copied dependency: $dep" -ForegroundColor Green
    }
}

# Register for COM
if ($RegisterCOM) {
    Write-Host ""
    Write-Host "Registering COM interop..." -ForegroundColor Yellow

    $regasm = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"
    if (-not (Test-Path $regasm)) {
        $regasm = "C:\Windows\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe"
    }

    if (Test-Path $regasm) {
        & $regasm $destDll /tlb /codebase /nologo
        if ($LASTEXITCODE -eq 0) {
            Write-Host "✓ COM registration successful" -ForegroundColor Green
        } else {
            Write-Error "COM registration failed!"
            exit 1
        }
    } else {
        Write-Error "RegAsm.exe not found! Ensure .NET Framework SDK is installed."
        exit 1
    }
}

# Install to GAC (optional)
if ($InstallToGAC) {
    Write-Host ""
    Write-Host "Installing to Global Assembly Cache..." -ForegroundColor Yellow

    $gacutil = "C:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\gacutil.exe"
    if (-not (Test-Path $gacutil)) {
        $gacutil = "C:\Program Files (x86)\Microsoft SDKs\Windows\v8.1A\bin\NETFX 4.5.1 Tools\gacutil.exe"
    }

    if (Test-Path $gacutil) {
        & $gacutil /i $destDll /nologo
        if ($LASTEXITCODE -eq 0) {
            Write-Host "✓ GAC installation successful" -ForegroundColor Green
        } else {
            Write-Warning "GAC installation failed (this is optional)"
        }
    } else {
        Write-Warning "gacutil.exe not found. GAC installation skipped (this is optional)."
    }
}

# Create sample config file
Write-Host ""
Write-Host "Creating sample configuration..." -ForegroundColor Yellow

$configContent = @"
<?xml version="1.0" encoding="utf-8"?>
<configuration>
  <appSettings>
    <!-- CloudMailKit Configuration -->
    <!-- Get these values from Azure AD App Registration -->
    <add key="CloudMailKit.TenantId" value="your-tenant-id-here" />
    <add key="CloudMailKit.ClientId" value="your-client-id-here" />
    <add key="CloudMailKit.ClientSecret" value="your-client-secret-here" />
    <add key="CloudMailKit.MailboxAddress" value="noreply@yourcompany.com" />
  </appSettings>
</configuration>
"@

$configPath = Join-Path $InstallPath "CloudMailKit.config.sample"
$configContent | Out-File -FilePath $configPath -Encoding UTF8
Write-Host "✓ Created sample config: $configPath" -ForegroundColor Green

# Create VB6 sample
$vb6Sample = @"
' CloudMailKit VB6 Sample
' Copy this to your VB6 project

Private Sub SendEmail()
    On Error GoTo ErrorHandler

    ' Create CloudMailKit client
    Dim client As Object
    Set client = CreateObject("CloudMailKit.UnifiedMailClient")

    ' Initialize with your Azure AD credentials
    client.Initialize _
        "your-tenant-id", _
        "your-client-id", _
        "your-client-secret", _
        "sender@yourcompany.com"

    ' Send email
    client.SendSimple _
        "sender@yourcompany.com", _
        "recipient@example.com", _
        "Test Email from VB6", _
        "This email was sent using CloudMailKit!", _
        False

    MsgBox "Email sent successfully!", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
"@

$vb6Path = Join-Path $InstallPath "VB6-Sample.txt"
$vb6Sample | Out-File -FilePath $vb6Path -Encoding ASCII
Write-Host "✓ Created VB6 sample: $vb6Path" -ForegroundColor Green

# Create Classic ASP sample
$aspSample = @"
<%@ Language=VBScript %>
<%
' CloudMailKit Classic ASP Sample
On Error Resume Next

Set client = Server.CreateObject("CloudMailKit.UnifiedMailClient")
client.Initialize _
    "your-tenant-id", _
    "your-client-id", _
    "your-client-secret", _
    "noreply@yourcompany.com"

client.SendSimple _
    "noreply@yourcompany.com", _
    "customer@example.com", _
    "Welcome to Our Site", _
    "<h1>Welcome!</h1><p>Thank you for signing up.</p>", _
    True

If Err.Number = 0 Then
    Response.Write "Email sent successfully!"
Else
    Response.Write "Error: " & Err.Description
End If

Set client = Nothing
%>
"@

$aspPath = Join-Path $InstallPath "ClassicASP-Sample.asp"
$aspSample | Out-File -FilePath $aspPath -Encoding ASCII
Write-Host "✓ Created Classic ASP sample: $aspPath" -ForegroundColor Green

# Create .NET Framework sample
$dotnetSample = @"
using CloudMailKit;
using System;

namespace LegacyEmailApp
{
    class Program
    {
        static void Main(string[] args)
        {
            // Initialize CloudMailKit
            var client = new UnifiedMailClient();
            client.Initialize(
                tenantId: "your-tenant-id",
                clientId: "your-client-id",
                clientSecret: "your-client-secret",
                mailboxAddress: "sender@yourcompany.com"
            );

            // Send email
            client.SendSimple(
                from: "sender@yourcompany.com",
                to: "recipient@example.com",
                subject: "Test Email",
                body: "Hello from .NET Framework!",
                isHtml: false
            );

            Console.WriteLine("Email sent successfully!");
        }
    }
}
"@

$dotnetPath = Join-Path $InstallPath "DotNet-Sample.cs"
$dotnetSample | Out-File -FilePath $dotnetPath -Encoding UTF8
Write-Host "✓ Created .NET sample: $dotnetPath" -ForegroundColor Green

# Display next steps
Write-Host ""
Write-Host "============================================" -ForegroundColor Green
Write-Host "Installation Complete!" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Green
Write-Host ""
Write-Host "Installation Path: $InstallPath" -ForegroundColor Cyan
Write-Host ""
Write-Host "Next Steps:" -ForegroundColor Yellow
Write-Host "1. Register an Azure AD app (see LEGACY-RESCUE-GUIDE.md)"
Write-Host "2. Edit the sample config with your credentials"
Write-Host "3. Test with one of the sample files"
Write-Host ""
Write-Host "For VB6/Classic ASP:" -ForegroundColor Yellow
Write-Host "  - COM is registered and ready to use"
Write-Host "  - See: $vb6Path"
Write-Host "  - See: $aspPath"
Write-Host ""
Write-Host "For .NET Framework:" -ForegroundColor Yellow
Write-Host "  - Add reference to: $destDll"
Write-Host "  - See: $dotnetPath"
Write-Host ""
Write-Host "Documentation:" -ForegroundColor Yellow
Write-Host "  - README.md"
Write-Host "  - LEGACY-RESCUE-GUIDE.md"
Write-Host "  - MIGRATION-GUIDE.md"
Write-Host ""
Write-Host "Need help? https://github.com/PSScript/CloudMailKit/issues" -ForegroundColor Cyan
Write-Host ""
