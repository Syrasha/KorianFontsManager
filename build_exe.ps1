# KorianFontsManager Build Script for Windows
# This script automates the process of creating a single-file executable.

Write-Host "Checking for virtual environment..." -ForegroundColor Cyan
if (-Not (Test-Path ".venv")) {
    Write-Host "Virtual environment not found. Creating one..." -ForegroundColor Yellow
    python -m venv .venv
}

Write-Host "Installing dependencies..." -ForegroundColor Cyan
& .venv\Scripts\python.exe -m pip install -r requirements.txt pyinstaller

Write-Host "Building executable..." -ForegroundColor Cyan
# --onefile: Create a single executable
# --noconsole: Don't open a terminal window when running the app
# --name: Set the name of the output executable
# --clean: Clean PyInstaller cache and remove temporary files before building
& .venv\Scripts\pyinstaller --onefile --noconsole --name "KorianFontsManager" --clean main.py

if ($LASTEXITCODE -eq 0) {
    Write-Host "`nCleaning up build files..." -ForegroundColor Cyan
    if (Test-Path "build") { Remove-Item -Path "build" -Recurse -Force }
    if (Test-Path "KorianFontsManager.spec") { Remove-Item -Path "KorianFontsManager.spec" -Force }
    
    Write-Host "`nBuild successful! You can find the executable in the 'dist' folder." -ForegroundColor Green
} else {
    Write-Error "Build failed."
}
