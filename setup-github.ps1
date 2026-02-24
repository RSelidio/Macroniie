# GitHub Setup Script for Macronnie

$ErrorActionPreference = "Continue"

# Check and install Git
Write-Host "Checking Git installation..." -ForegroundColor Cyan
try {
    $gitVersion = git --version
    Write-Host "Git found: $gitVersion" -ForegroundColor Green
}
catch {
    Write-Host "Git not found. Installing Git..." -ForegroundColor Yellow
    
    # Try Chocolatey
    if (Get-Command choco -ErrorAction SilentlyContinue) {
        choco install git -y --no-progress
    }
    # Try direct download and installation
    else {
        Write-Host "Downloading Git installer..." -ForegroundColor Yellow
        $gitInstaller = "$env:TEMP\GitInstaller.exe"
        Invoke-WebRequest -Uri "https://github.com/git-for-windows/git/releases/download/v2.43.0.windows.1/Git-2.43.0-64-bit.exe" -OutFile $gitInstaller
        
        Write-Host "Installing Git..." -ForegroundColor Yellow
        Start-Process -FilePath $gitInstaller -ArgumentList "/SILENT /COMPONENTS=git,gitlfs" -Wait
        Remove-Item $gitInstaller -Force
    }
    
    $gitVersion = git --version
    Write-Host "Git installed: $gitVersion" -ForegroundColor Green
}

# Navigate to project directory
Set-Location "C:\Users\Yusak\Documents\GithubMacro"

# Initialize git repo
Write-Host "`nInitializing Git repository..." -ForegroundColor Cyan
git init
git config user.name "Ronnie"
git config user.email "ronnie@github.local"

# Add all files
Write-Host "Adding files to git..." -ForegroundColor Cyan
git add .

# Show what will be committed
Write-Host "`nFiles to be committed:" -ForegroundColor Yellow
git status

# Commit
Write-Host "`nCreating initial commit..." -ForegroundColor Cyan
git commit -m "Initial commit: Macronnie macro recorder with keyboard and mouse support"

# Add remote
Write-Host "`nAdding GitHub remote..." -ForegroundColor Cyan
git branch -M main
git remote add origin https://github.com/RSelidio/Macroniie.git

# Push to GitHub
Write-Host "`nPushing to GitHub..." -ForegroundColor Cyan
Write-Host "Note: You may be prompted for GitHub credentials." -ForegroundColor Yellow
git push -u origin main

Write-Host "`n✅ GitHub setup complete!" -ForegroundColor Green
Write-Host "Repository: https://github.com/RSelidio/Macroniie" -ForegroundColor Green
