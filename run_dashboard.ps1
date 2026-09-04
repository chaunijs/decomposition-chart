<#
.SYNOPSIS
    Launcher script for Decomposition Studio (Streamlit Dashboard).
.DESCRIPTION
    Runs the Streamlit application directly using the local virtual environment (.venv)
    or falls back to the system Python interpreter if venv execution is restricted by corporate policy.
#>

$ErrorActionPreference = "Stop"

# Navigate to the directory where this script is located
Set-Location -Path $PSScriptRoot

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "       Starting Decomposition Studio (Streamlit)         " -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan

# Determine the best Python interpreter
$venvPython = Join-Path $PSScriptRoot ".venv\Scripts\python.exe"
$pythonCmd = $null

if (Test-Path $venvPython) {
    try {
        & $venvPython --version | Out-Null
        $pythonCmd = $venvPython
        Write-Host "[+] Using virtual environment: .venv" -ForegroundColor Green
    }
    catch {
        Write-Host "[!] Virtual environment execution blocked or failed. Checking system Python..." -ForegroundColor Yellow
    }
}

if (-not $pythonCmd) {
    # Check if system python is available
    if (Get-Command python -ErrorAction SilentlyContinue) {
        $pythonCmd = "python"
        Write-Host "[+] Using system Python" -ForegroundColor Green
    }
    elseif (Get-Command py -ErrorAction SilentlyContinue) {
        $pythonCmd = "py"
        Write-Host "[+] Using Python launcher (py)" -ForegroundColor Green
    }
    else {
        Write-Host "[ERROR] No Python interpreter found on PATH." -ForegroundColor Red
        Write-Host "Please ensure Python 3.10+ is installed." -ForegroundColor Red
        Read-Host "Press Enter to exit..."
        exit 1
    }
}

# Verify / Run Streamlit
Write-Host "[*] Launching Streamlit application..." -ForegroundColor Cyan
try {
    & $pythonCmd -m streamlit run app.py
}
catch {
    Write-Host "[ERROR] Failed to start Streamlit: $_" -ForegroundColor Red
    Write-Host "[*] You can also try running: $pythonCmd -m pip install -r requirements.txt" -ForegroundColor Yellow
    Read-Host "Press Enter to exit..."
}
