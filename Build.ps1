#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Build, register, and launch the Hello VSTO Excel add-in.

.DESCRIPTION
    This script performs the following steps:
    1. Locates MSBuild.exe (required for VSTO projects)
    2. Restores NuGet packages
    3. Builds the VSTO project
    4. Registers the add-in in Windows Registry
    5. Launches Excel with the add-in loaded

.EXAMPLE
    .\Build.ps1
    Builds and launches the Hello VSTO add-in in Excel.
#>

# Stop on any error
$ErrorActionPreference = "Stop"

Write-Host "=====================================" -ForegroundColor Cyan
Write-Host "   Hello VSTO - Build & Launch" -ForegroundColor Cyan
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host ""

# ============================================================================
# Step 1: Locate MSBuild
# ============================================================================
Write-Host "[1/5] Locating MSBuild..." -ForegroundColor Yellow

# Try to find MSBuild using vswhere (installed with VS 2017+)
$vswherePath = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"

if (Test-Path $vswherePath) {
    $vsPath = & $vswherePath -latest -products * -requires Microsoft.Component.MSBuild -property installationPath
    if ($vsPath) {
        $msbuildPath = Join-Path $vsPath "MSBuild\Current\Bin\MSBuild.exe"
        if (-not (Test-Path $msbuildPath)) {
            # Try older VS versions
            $msbuildPath = Join-Path $vsPath "MSBuild\15.0\Bin\MSBuild.exe"
        }
    }
}

# Fallback: Check common MSBuild locations
if (-not $msbuildPath -or -not (Test-Path $msbuildPath)) {
    $msbuildLocations = @(
        "${env:ProgramFiles}\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles}\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles}\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2019\Community\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2019\Professional\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2019\Enterprise\MSBuild\Current\Bin\MSBuild.exe"
    )

    foreach ($location in $msbuildLocations) {
        if (Test-Path $location) {
            $msbuildPath = $location
            break
        }
    }
}

if (-not $msbuildPath -or -not (Test-Path $msbuildPath)) {
    Write-Host "ERROR: MSBuild.exe not found!" -ForegroundColor Red
    Write-Host "Please install Visual Studio with .NET desktop development workload." -ForegroundColor Red
    exit 1
}

Write-Host "   Found MSBuild: $msbuildPath" -ForegroundColor Green

# ============================================================================
# Step 2: Restore NuGet Packages
# ============================================================================
Write-Host ""
Write-Host "[2/5] Restoring NuGet packages..." -ForegroundColor Yellow

& $msbuildPath /t:Restore /p:Configuration=Debug ".\HelloVsto\HelloVsto.csproj"

if ($LASTEXITCODE -ne 0) {
    Write-Host "ERROR: NuGet restore failed!" -ForegroundColor Red
    exit 1
}

Write-Host "   Packages restored successfully" -ForegroundColor Green

# ============================================================================
# Step 3: Build Project
# ============================================================================
Write-Host ""
Write-Host "[3/5] Building project..." -ForegroundColor Yellow

& $msbuildPath /p:Configuration=Debug ".\HelloVsto\HelloVsto.csproj"

if ($LASTEXITCODE -ne 0) {
    Write-Host "ERROR: Build failed!" -ForegroundColor Red
    exit 1
}

Write-Host "   Build completed successfully" -ForegroundColor Green

# ============================================================================
# Step 4: Verify Build Output
# ============================================================================
Write-Host ""
Write-Host "[4/5] Verifying build output..." -ForegroundColor Yellow

$vstoPath = ".\HelloVsto\bin\Debug\HelloVsto.vsto"

if (-not (Test-Path $vstoPath)) {
    Write-Host "ERROR: HelloVsto.vsto not found at: $vstoPath" -ForegroundColor Red
    Write-Host "Build may have completed but VSTO manifest was not generated." -ForegroundColor Red
    exit 1
}

Write-Host "   VSTO manifest found: $vstoPath" -ForegroundColor Green

# ============================================================================
# Step 5: Register Add-in in Registry
# ============================================================================
Write-Host ""
Write-Host "[5/5] Registering add-in..." -ForegroundColor Yellow

# Registry path for Excel add-ins
$registryPath = "HKCU:\Software\Microsoft\Office\Excel\Addins\HelloVsto"

# Get absolute path to .vsto file
$vstoAbsolutePath = Resolve-Path -Path $vstoPath
$vstoRegistryValue = "file:///$vstoAbsolutePath|vstolocal"

# Create registry key if it doesn't exist
if (-not (Test-Path $registryPath)) {
    New-Item -Path $registryPath -Force | Out-Null
    Write-Host "   Created registry key: $registryPath" -ForegroundColor Green
}

# Set registry values
Set-ItemProperty -Path $registryPath -Name "Description" -Value "Hello VSTO Learning Project"
Set-ItemProperty -Path $registryPath -Name "FriendlyName" -Value "Hello VSTO"
Set-ItemProperty -Path $registryPath -Name "LoadBehavior" -Value 3
Set-ItemProperty -Path $registryPath -Name "Manifest" -Value $vstoRegistryValue

Write-Host "   Add-in registered successfully" -ForegroundColor Green
Write-Host "   Manifest path: $vstoRegistryValue" -ForegroundColor Gray

# ============================================================================
# Step 6: Launch Excel
# ============================================================================
Write-Host ""
Write-Host "Launching Excel with Hello VSTO add-in..." -ForegroundColor Cyan
Write-Host ""
Write-Host "IMPORTANT:" -ForegroundColor Yellow
Write-Host "  - A console window will appear alongside Excel" -ForegroundColor White
Write-Host "  - The console shows VSTO lifecycle events in real-time" -ForegroundColor White
Write-Host "  - Look for the 'Hello VSTO' tab in Excel ribbon" -ForegroundColor White
Write-Host "  - Click the 'Hello' button to write to cell A1" -ForegroundColor White
Write-Host ""

# Common Excel installation paths
$excelPaths = @(
    "${env:ProgramFiles}\Microsoft Office\root\Office16\excel.exe",
    "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\excel.exe",
    "${env:ProgramFiles}\Microsoft Office\Office16\excel.exe",
    "${env:ProgramFiles(x86)}\Microsoft Office\Office16\excel.exe"
)

$excelPath = $null
foreach ($path in $excelPaths) {
    if (Test-Path $path) {
        $excelPath = $path
        break
    }
}

if (-not $excelPath) {
    Write-Host "WARNING: Could not find Excel.exe automatically" -ForegroundColor Yellow
    Write-Host "Please launch Excel manually. The Hello VSTO add-in is registered and will load." -ForegroundColor Yellow
    exit 0
}

# Launch Excel with /x flag (suppress startup screen)
Start-Process -FilePath $excelPath -ArgumentList "/x"

Write-Host "Excel launched successfully!" -ForegroundColor Green
Write-Host ""
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host "Build complete! Watch the console window that appears." -ForegroundColor Cyan
Write-Host "=====================================" -ForegroundColor Cyan
