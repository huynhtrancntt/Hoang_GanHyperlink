# release.ps1
# GanHyperlink - Auto Release Script
# Usage:
#   .\release.ps1
#   .\release.ps1 v1.0.2

param([string]$Version = "")

$ErrorActionPreference = "Stop"

# ── Banner ───────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ====================================================" -ForegroundColor Cyan
Write-Host "    GanHyperlink  |  Auto Release to GitHub" -ForegroundColor Cyan
Write-Host "  ====================================================" -ForegroundColor Cyan
Write-Host ""

# ── Hoi version neu chua co ──────────────────────────────────────────────────
if ($Version -eq "") {
    $Version = Read-Host "  Nhap phien ban (vi du: v1.0.2)"
    Write-Host ""
}

# ── Kiem tra dinh dang vX.Y.Z ────────────────────────────────────────────────
if ($Version -notmatch "^v\d+\.\d+\.\d+$") {
    Write-Host "  [LOI] Dinh dang sai: $Version" -ForegroundColor Red
    Write-Host "        Phai la dang: v1.0.0 / v1.2.3 / v2.0.0" -ForegroundColor Yellow
    exit 1
}

# ── Kiem tra tag chua ton tai local ──────────────────────────────────────────
$existingTag = git tag --list $Version
if ($existingTag) {
    Write-Host "  [LOI] Tag $Version da ton tai local!" -ForegroundColor Red
    Write-Host "        Chay: git tag -d $Version  de xoa truoc" -ForegroundColor Yellow
    exit 1
}

Write-Host "  Phien ban : $Version" -ForegroundColor Green
Write-Host ""

# ── [1/4] Git status ─────────────────────────────────────────────────────────
Write-Host "  [1/4] Kiem tra thay doi..." -ForegroundColor Yellow
$changes = git status --short
if ($changes) {
    Write-Host $changes -ForegroundColor Gray
}

# ── [2/4] Auto commit neu co thay doi ────────────────────────────────────────
Write-Host ""
Write-Host "  [2/4] Commit thay doi..." -ForegroundColor Yellow
$isDirty = git status --porcelain
if ($isDirty) {
    git add -A
    git commit -m "chore: release $Version"
    Write-Host "  [OK] Da commit." -ForegroundColor Green
} else {
    Write-Host "  [OK] Khong co thay doi moi." -ForegroundColor Gray
}

# ── [3/4] Push main ──────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  [3/4] Push len GitHub (main)..." -ForegroundColor Yellow
git push origin main
if ($LASTEXITCODE -ne 0) {
    Write-Host "  [LOI] Push main that bai!" -ForegroundColor Red
    exit 1
}
Write-Host "  [OK] Push main thanh cong." -ForegroundColor Green

# ── [4/4] Tao tag, push, xoa local ──────────────────────────────────────────
Write-Host ""
Write-Host "  [4/4] Tao tag $Version va push..." -ForegroundColor Yellow
git tag $Version -m "Release $Version"
git push origin $Version
if ($LASTEXITCODE -ne 0) {
    Write-Host "  [LOI] Push tag that bai!" -ForegroundColor Red
    git tag -d $Version 2>$null
    exit 1
}

# Xoa tag local sau khi push
git tag -d $Version | Out-Null
Write-Host "  [OK] Tag $Version da push va xoa khoi local." -ForegroundColor Green

# ── Hoan thanh ───────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ====================================================" -ForegroundColor Cyan
Write-Host "   XONG! GitHub Actions dang build EXE..." -ForegroundColor Green
Write-Host ""
Write-Host "   Theo doi build:" -ForegroundColor White
Write-Host "   https://github.com/huynhtrancntt/Hoang_GanHyperlink/actions" -ForegroundColor Blue
Write-Host ""
Write-Host "   Release se co tai:" -ForegroundColor White
Write-Host "   https://github.com/huynhtrancntt/Hoang_GanHyperlink/releases/tag/$Version" -ForegroundColor Blue
Write-Host "  ====================================================" -ForegroundColor Cyan
Write-Host ""
