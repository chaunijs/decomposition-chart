@echo off
cd /d "%~dp0"

echo [*] Starting Decomposition Studio...
if not exist ".venv\Scripts\python.exe" (
    echo [*] Setting up virtual environment...
    python -m venv .venv
    .\.venv\Scripts\python.exe -m pip install -r requirements.txt
)

.\.venv\Scripts\python.exe -m streamlit run app.py
pause
