@echo off
setlocal enabledelayedexpansion

:: ============================================================
::  release.bat — Tạo GitHub Release tự động
::  Cách dùng:
::      release.bat v1.0.1
::      release.bat v2.0.0
:: ============================================================

set VERSION=%~1

:: ── Nếu chưa có tham số → hỏi trực tiếp ────────────────────
if "%VERSION%"=="" (
    echo.
    echo  =====================================================
    echo    GanHyperlink ^| Cong cu phat hanh tu dong
    echo  =====================================================
    echo.
    set /p VERSION= Nhap phien ban (vi du v1.0.2): 
    echo.
)

:: ── Kiểm tra định dạng vX.Y.Z ───────────────────────────────
echo %VERSION% | findstr /r "^v[0-9][0-9]*\.[0-9][0-9]*\.[0-9][0-9]*$" >nul
if errorlevel 1 (
    echo.
    echo  [LOI] Dinh dang version sai: %VERSION%
    echo  Phai la dang: v1.0.0 / v1.2.3 / v2.0.0
    echo.
    pause
    exit /b 1
)

echo.
echo  =====================================================
echo    GanHyperlink — Release %VERSION%
echo  =====================================================
echo.

:: ── Kiểm tra tag chưa tồn tại ───────────────────────────────
git tag | findstr /x "%VERSION%" >nul
if not errorlevel 1 (
    echo  [LOI] Tag %VERSION% da ton tai!
    echo  Dung phien ban khac, vi du: release.bat v1.0.2
    echo.
    pause
    exit /b 1
)

:: ── Git status ───────────────────────────────────────────────
echo  [1/4] Kiem tra thay doi...
git status --short
echo.

:: ── Add & Commit nếu có thay đổi ────────────────────────────
git diff --quiet --cached
git diff --quiet
if errorlevel 1 (
    echo  [2/4] Co thay doi - tu dong commit...
    git add -A
    git commit -m "chore: release %VERSION%"
) else (
    echo  [2/4] Khong co thay doi moi.
)
echo.

:: ── Push main ───────────────────────────────────────────────
echo  [3/4] Push len GitHub...
git push origin main
if errorlevel 1 (
    echo  [LOI] Push that bai!
    pause
    exit /b 1
)
echo.

:: ── Tạo tag & push tag → trigger GitHub Actions ─────────────
echo  [4/4] Tao tag %VERSION% va push...
git tag %VERSION% -m "Release %VERSION%"
git push origin %VERSION%
if errorlevel 1 (
    echo  [LOI] Push tag that bai!
    pause
    exit /b 1
)

:: ── Xóa tag dưới local (chỉ giữ trên remote) ────────────────
git tag -d %VERSION%
echo  [OK] Da xoa tag %VERSION% khoi local.

echo.
echo  =====================================================
echo   XONG! GitHub Actions dang build EXE...
echo.
echo   Theo doi tai:
echo   https://github.com/huynhtrancntt/Hoang_GanHyperlink/actions
echo.
echo   File ZIP se xuat hien tai:
echo   https://github.com/huynhtrancntt/Hoang_GanHyperlink/releases/tag/%VERSION%
echo  =====================================================
echo.
pause
