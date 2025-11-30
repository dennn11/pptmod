# Build script for creating standalone executable
# This script packages the application using PyInstaller

Write-Host "Building pptmod executable..." -ForegroundColor Green

# Clean previous builds
if (Test-Path ".\dist") {
    Write-Host "Cleaning previous builds..." -ForegroundColor Yellow
    Remove-Item -Recurse -Force ".\dist"
}
if (Test-Path ".\build") {
    Remove-Item -Recurse -Force ".\build"
}

# Install build dependencies if not already installed
Write-Host "Installing build dependencies..." -ForegroundColor Yellow
pip install pyinstaller python-pptx pywin32

# Build the executable
Write-Host "Creating executable with PyInstaller..." -ForegroundColor Yellow
pyinstaller --onefile `
    --name pptmod `
    --icon NONE `
    --console `
    --add-data "config.json;." `
    main.py

if ($LASTEXITCODE -eq 0) {
    Write-Host "`nBuild successful!" -ForegroundColor Green
    Write-Host "Executable location: .\dist\pptmod.exe" -ForegroundColor Cyan
    Write-Host "`nTo distribute:" -ForegroundColor Yellow
    Write-Host "  1. Share the 'dist\pptmod.exe' file" -ForegroundColor White
    Write-Host "  2. Include a 'config.json' file for users" -ForegroundColor White
    Write-Host "  3. Users can run: pptmod.exe presentation.pptx" -ForegroundColor White
} else {
    Write-Host "`nBuild failed!" -ForegroundColor Red
    exit 1
}
