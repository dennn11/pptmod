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
if (Test-Path ".\pptmod-gui.spec") {
    Remove-Item -Force ".\pptmod-gui.spec"
}

# Install build dependencies if not already installed
Write-Host "Installing build dependencies..." -ForegroundColor Yellow
uv pip install pyinstaller python-pptx pywin32 wxPython

# Build the GUI executable
Write-Host "Creating GUI executable with PyInstaller..." -ForegroundColor Yellow
uv run pyinstaller --onefile `
    --name pptmod `
    --windowed `
    --icon NONE `
    --add-data "pptmodconfig.json;." `
    --hidden-import=wx `
    --hidden-import=wx.grid `
    --hidden-import=win32com.client `
    --hidden-import=pptx `
    --collect-all wx `
    gui.py

if ($LASTEXITCODE -eq 0) {
    Write-Host "`n========================================" -ForegroundColor Green
    Write-Host "   BUILD SUCCESSFUL!" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "`nExecutable created: .\dist\pptmod.exe" -ForegroundColor Cyan
    Write-Host "`nYou can now run the application by double-clicking:" -ForegroundColor Yellow
    Write-Host "  .\dist\pptmod.exe" -ForegroundColor White
    Write-Host "`nTo distribute to others:" -ForegroundColor Yellow
    Write-Host "  1. Share the 'dist\pptmod.exe' file" -ForegroundColor White
    Write-Host "  2. Include a 'pptmodconfig.json' file (optional - users can create their own)" -ForegroundColor White
    Write-Host "  3. No Python or dependencies needed!" -ForegroundColor White
    Write-Host "`nNote: The exe file is self-contained (~150MB)" -ForegroundColor Gray
} else {
    Write-Host "`n========================================" -ForegroundColor Red
    Write-Host "   BUILD FAILED!" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    exit 1
}
