@echo off
setlocal enabledelayedexpansion
set "VERSION=%~1"
if "!VERSION!"=="" (
    echo.
    echo  GanHyperlink - Auto Release
    echo.
    set /p "VERSION= Nhap phien ban (vi du: 1.0.2): "
    echo.
)
if "!VERSION:~0,1!" neq "v" set "VERSION=v!VERSION!"
echo [INFO] Phien ban: !VERSION!
echo.
git tag --list "!VERSION!" | findstr "!VERSION!" >nul
if not errorlevel 1 (
    echo [LOI] Tag !VERSION! da ton tai local!
    pause & exit /b 1
)
echo [1/4] Kiem tra thay doi...
git status --short
echo.
echo [2/4] Commit thay doi...
git status --porcelain > "%TEMP%\gs.txt"
for %%A in ("%TEMP%\gs.txt") do set sz=%%~zA
if !sz! gtr 0 (git add -A && git commit -m "chore: release !VERSION!" && echo [OK] Da commit.) else echo [OK] Khong co thay doi.
echo.
echo [3/4] Push len GitHub...
git push origin main
if errorlevel 1 (echo [LOI] Push that bai! & pause & exit /b 1)
echo [OK] Push thanh cong.
echo.
echo [4/4] Tao va push tag !VERSION!...
git tag "!VERSION!" -m "Release !VERSION!"
git push origin "!VERSION!"
if errorlevel 1 (git tag -d "!VERSION!" 2>nul & echo [LOI] Push tag that bai! & pause & exit /b 1)
git tag -d "!VERSION!" >nul 2>&1
echo [OK] Tag !VERSION! da push va xoa khoi local.
echo.
echo XONG! GitHub Actions dang build EXE...
echo https://github.com/huynhtrancntt/Hoang_GanHyperlink/releases/tag/!VERSION!
echo.
pause