@echo off
setlocal enabledelayedexpansion

:: ============================================================
::  release.bat - GanHyperlink Auto Release
::  Usage: release.bat           (will prompt for version)
::         release.bat 1.0.4     (auto-add v prefix)
::         release.bat v1.0.4
:: ============================================================

set "VERSION=%~1"

:: -- Prompt if no argument --
if "!VERSION!"=="" (
    echo.
    echo  ====================================================
    echo   GanHyperlink ^| Auto Release to GitHub
    echo  ====================================================
    echo.
    set /p "VERSION= Enter version (e.g. 1.0.4): "
    echo.
)

:: -- Auto-add v prefix if missing --
if "!VERSION:~0,1!" neq "v" set "VERSION=v!VERSION!"

echo  [INFO] Version: !VERSION!
echo.

:: -- Clean local tag if exists --
git tag --list "!VERSION!" | findstr "!VERSION!" >nul 2>&1
if not errorlevel 1 (
    echo  [INFO] Deleting local tag: !VERSION!
    git tag -d "!VERSION!" >nul 2>&1
)

:: -- Clean remote tag if exists --
git ls-remote --tags origin "refs/tags/!VERSION!" | findstr "!VERSION!" >nul 2>&1
if not errorlevel 1 (
    echo  [INFO] Deleting remote tag: !VERSION!
    git push origin :refs/tags/!VERSION! >nul 2>&1
)
echo.

:: -- [1/4] Status --
echo  [1/4] Git status...
git status --short
echo.

:: -- [2/4] Auto commit if dirty --
echo  [2/4] Committing changes...
git status --porcelain > "%TEMP%\gs_release.txt"
for %%F in ("%TEMP%\gs_release.txt") do set "FSIZE=%%~zF"
if !FSIZE! gtr 0 (
    git add -A
    git commit -m "chore: release !VERSION!"
    echo  [OK] Committed.
) else (
    echo  [OK] Nothing to commit.
)
echo.

:: -- [3/4] Push main --
echo  [3/4] Pushing main to GitHub...
git push origin main
if errorlevel 1 (
    echo  [ERROR] Push failed!
    pause
    exit /b 1
)
echo  [OK] Push successful.
echo.

:: -- [4/4] Tag, push, delete local --
echo  [4/4] Creating tag !VERSION!...
git tag "!VERSION!" -m "Release !VERSION!"
git push origin "!VERSION!"
if errorlevel 1 (
    echo  [ERROR] Tag push failed!
    git tag -d "!VERSION!" >nul 2>&1
    pause
    exit /b 1
)
git tag -d "!VERSION!" >nul 2>&1
echo  [OK] Tag !VERSION! pushed and deleted locally.
echo.

:: -- Done --
echo  ====================================================
echo   DONE! GitHub Actions is building EXE...
echo.
echo   Monitor: https://github.com/huynhtrancntt/Hoang_GanHyperlink/actions
echo   Release: https://github.com/huynhtrancntt/Hoang_GanHyperlink/releases/tag/!VERSION!
echo  ====================================================
echo.
pause