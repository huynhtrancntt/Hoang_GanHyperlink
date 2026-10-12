@echo off
setlocal enabledelayedexpansion

:: release.bat - GanHyperlink Auto Release
:: Usage:
::   release.bat
::   release.bat 1.0.2
::   release.bat v1.0.2

set "VERSION=%~1"

:: ── Hoi version neu chua co ──────────────────────────────────
if "!VERSION!"=="" (
    echo.
    echo  ====================================================
    echo    GanHyperlink  ^|  Auto Release to GitHub
    echo  ====================================================
    echo.
    set /p "VERSION= Nhap phien ban (vi du: 1.0.2): "
    echo.
)

:: ── Tu dong them "v" neu thieu ───────────────────────────────
echo !VERSION! | findstr /r "^[0-9]" >nul
if not errorlevel 1 (
    set "VERSION=v!VERSION!"
)

:: ── Kiem tra dinh dang vX.Y.Z ────────────────────────────────
echo !VERSION! | findstr /r "^v[0-9][0-9]*\.[0-9][0-9]*\.[0-9][0-9]*$" >nul
if errorlevel 1 (
    echo.
    echo  [LOI] Dinh dang sai: !VERSION!
    echo        Phai la dang: 1.0.0 hoac v1.0.0
    echo.
    pause
    exit /b 1
)

echo.
echo  [INFO] Phien ban: !VERSION!
echo.

:: ── Kiem tra tag chua ton tai ────────────────────────────────
git tag --list "!VERSION!" | findstr "!VERSION!" >nul
if not errorlevel 1 (
    echo  [LOI] Tag !VERSION! da ton tai local!
    echo        Chay: git tag -d !VERSION!
    echo.
    pause
    exit /b 1
)

:: ── [1/4] Git status ─────────────────────────────────────────
echo  [1/4] Kiem tra thay doi...
git status --short
echo.

:: ── [2/4] Commit neu co thay doi ─────────────────────────────
echo  [2/4] Commit thay doi...
git diff --quiet 2>nul
git diff --cached --quiet 2>nul
git status --porcelain > "%TEMP%\gs_tmp.txt"
for %%A in ("%TEMP%\gs_tmp.txt") do set sz=%%~zA
if !sz! gtr 0 (
    git add -A
    git commit -m "chore: release !VERSION!"
    echo  [OK] Da commit.
) else (
    echo  [OK] Khong co thay doi moi.
)
echo.

:: ── [3/4] Push main ──────────────────────────────────────────
echo  [3/4] Push len GitHub (main)...
git push origin main
if errorlevel 1 (
    echo  [LOI] Push main that bai!
    pause
    exit /b 1
)
echo  [OK] Push main thanh cong.
echo.

:: ── [4/4] Tag, push, xoa local ───────────────────────────────
echo  [4/4] Tao tag !VERSION! va push...
git tag "!VERSION!" -m "Release !VERSION!"
git push origin "!VERSION!"
if errorlevel 1 (
    echo  [LOI] Push tag that bai!
    git tag -d "!VERSION!" 2>nul
    pause
    exit /b 1
)
git tag -d "!VERSION!" >nul 2>&1
echo  [OK] Tag !VERSION! da push va xoa khoi local.
echo.

:: ── Hoan thanh ───────────────────────────────────────────────
echo  ====================================================
echo   XONG! GitHub Actions dang build EXE...
echo.
echo   Theo doi build:
echo   https://github.com/huynhtrancntt/Hoang_GanHyperlink/actions
echo.
echo   Release se co tai:
echo   https://github.com/huynhtrancntt/Hoang_GanHyperlink/releases/tag/!VERSION!
echo  ====================================================
echo.
pause
