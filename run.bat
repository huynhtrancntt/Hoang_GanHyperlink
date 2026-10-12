@echo off
chcp 65001 > nul
setlocal
set "DIR=%~dp0"
set "VENV=%DIR%venv"
set "PY=%VENV%\Scripts\pythonw.exe"
set "PIP=%VENV%\Scripts\pip.exe"

if not exist "%VENV%\Scripts\python.exe" (
    echo [1/3] Tao moi truong ao...
    python -m venv "%VENV%"
    if errorlevel 1 (
        echo LOI: Khong the tao venv.
        pause
        exit /b 1
    )
)

if not exist "%VENV%\Lib\site-packages\docx" (
    echo [2/3] Cai thu vien python-docx openpyxl...
    "%PIP%" install --quiet python-docx openpyxl
    if errorlevel 1 (
        echo LOI: Cai thu vien that bai.
        pause
        exit /b 1
    )
)

echo [3/3] Chay app...
start "" "%PY%" "%DIR%app.py"
endlocal
