@echo off
setlocal
set "DIR=%~dp0"
set "VENV=%DIR%venv"
set "PY=%VENV%\Scripts\pythonw.exe"
set "PIP=%VENV%\Scripts\pip.exe"

:: ── Tao moi truong ao neu chua co ──────────────────────────
if not exist "%VENV%\Scripts\python.exe" (
    echo [1/3] Chua co moi truong ao. Dang tao...
    python -m venv "%VENV%"
    if errorlevel 1 (
        echo LOI: Khong the tao venv. Kiem tra Python da duoc cai chua.
        pause
        exit /b 1
    )
    echo      Tao xong.
)

:: ── Cai thu vien neu chua co ────────────────────────────────
if not exist "%VENV%\Lib\site-packages\docx" (
    echo [2/3] Chua co thu vien. Dang cai: python-docx openpyxl...
    "%PIP%" install --quiet python-docx openpyxl
    if errorlevel 1 (
        echo LOI: Cai thu vien that bai. Kiem tra ket noi mang.
        pause
        exit /b 1
    )
    echo      Cai xong.
)

:: ── Chay app.py ─────────────────────────────────────────────
echo [3/3] Dang chay app...
start "" "%PY%" "%DIR%app.py"
endlocal
