<#
.SYNOPSIS
    Build CloudMailKit solution
#>

param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Release"
)

$ErrorActionPreference = "Stop"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "CloudMailKit Build Script" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Find MSBuild
$msbuild = & "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe" `
    -latest -requires Microsoft.Component.MSBuild -find MSBuild\**\Bin\MSBuild.exe `
    | Select-Object -First 1

if (-not $msbuild) {
    Write-Error "MSBuild not found. Please install Visual Studio 2019 or later."
    exit 1
}

Write-Host "Using MSBuild: $msbuild" -ForegroundColor Gray
Write-Host ""

# Restore NuGet packages
Write-Host "[1/3] Restoring NuGet packages..." -ForegroundColor Yellow
& $msbuild CloudMailKit.sln /t:Restore /p:Configuration=$Configuration
if ($LASTEXITCODE -ne 0) {
    Write-Error "NuGet restore failed"
    exit 1
}
Write-Host "      ✓ Success" -ForegroundColor Green
Write-Host ""

# Build solution
Write-Host "[2/3] Building solution ($Configuration)..." -ForegroundColor Yellow
& $msbuild CloudMailKit.sln /p:Configuration=$Configuration /verbosity:minimal
if ($LASTEXITCODE -ne 0) {
    Write-Error "Build failed"
    exit 1
}
Write-Host "      ✓ Success" -ForegroundColor Green
Write-Host ""

# Generate strong name key if missing
Write-Host "[3/3] Checking strong name key..." -ForegroundColor Yellow
$snkPath = "src\CloudMailKit\CloudMailKit.snk"
if (-not (Test-Path $snkPath)) {
    Write-Host "      Generating strong name key..." -ForegroundColor Gray
    & sn -k $snkPath
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Failed to generate strong name key"
        exit 1
    }
}
Write-Host "      ✓ Success" -ForegroundColor Green
Write-Host ""

Write-Host "========================================" -ForegroundColor Green
Write-Host "BUILD COMPLETED" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "Output: src\CloudMailKit\bin\$Configuration\" -ForegroundColor Cyan
```

---

# 6. Visual Studio Compilation Steps

## Step 1: Open Solution
```
1. Launch Visual Studio 2019/2022
2. File → Open → Project/Solution
3. Select CloudMailKit.sln
4. Wait for NuGet restore (automatic)
```

## Step 2: Set Configuration
```
1. Top toolbar: Debug/Release dropdown → Select "Release"
2. Platform: "Any CPU"
```

## Step 3: Build
```
Option A (GUI):
  Build → Build Solution (Ctrl+Shift+B)

Option B (Command Line):
  Open Developer Command Prompt
  cd <solution-directory>
  msbuild CloudMailKit.sln /p:Configuration=Release
```

## Step 4: Verify Output
```
Navigate to: src\CloudMailKit\bin\Release\net48\

Files should include:
  ✓ CloudMailKit.dll
  ✓ CloudMailKit.pdb
  ✓ CloudMailKit.xml (documentation)
  ✓ MimeKit.dll (dependency)
  ✓ System.Text.Json.dll (dependency)
```

## Step 5: Register COM
```
Open Developer Command Prompt as Administrator:

cd src\CloudMailKit\bin\Release\net48
regasm CloudMailKit.dll /tlb:CloudMailKit.tlb /codebase
```

Expected output:
```
Microsoft .NET Framework Assembly Registration Utility version 4.8.xxxx
Types registered successfully