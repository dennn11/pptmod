# Build script for creating standalone GUI executable
# This script packages the GUI application using PyInstaller

Write-Host "Building pptmod GUI executable..." -ForegroundColor Green

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
pip install pyinstaller python-pptx pywin32 wxPython

# Build the GUI executable
Write-Host "Creating GUI executable with PyInstaller..." -ForegroundColor Yellow
pyinstaller --onefile `
    --name pptmod-gui `
    --windowed `
    --icon NONE `
    --add-data "config.json;." `
    --hidden-import=wx `
    --hidden-import=win32com.client `
    --hidden-import=pptx `
    gui.py

if ($LASTEXITCODE -eq 0) {
    Write-Host "`nGUI Build successful!" -ForegroundColor Green
    Write-Host "Executable location: .\dist\pptmod-gui.exe" -ForegroundColor Cyan
    Write-Host "`nTo distribute:" -ForegroundColor Yellow
    Write-Host "  1. Share the 'dist\pptmod-gui.exe' file" -ForegroundColor White
    Write-Host "  2. Include a 'config.json' file for users" -ForegroundColor White
    Write-Host "  3. Users can double-click pptmod-gui.exe to launch" -ForegroundColor White
} else {
    Write-Host "`nGUI Build failed!" -ForegroundColor Red
    exit 1
}
