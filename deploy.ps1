# ============================================================
#  Deploy to production (Render).  Usage:  npm run deploy
#  With message:  npm run deploy -- "fix rate"
#  (ASCII-only on purpose: PowerShell 5.1 mis-reads non-BOM UTF-8)
# ============================================================
param([string]$m = "")

$ErrorActionPreference = "Stop"
Set-Location -Path $PSScriptRoot
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

if ($m -eq "") { $m = "deploy " + (Get-Date -Format "yyyy-MM-dd HH:mm") }

Write-Host ""
Write-Host "==> 1/3 build check (abort deploy if build fails)..." -ForegroundColor Cyan
npm run build
if ($LASTEXITCODE -ne 0) {
  Write-Host "X build FAILED - deploy aborted (fix errors above)" -ForegroundColor Red
  exit 1
}

Write-Host ""
Write-Host "==> 2/3 commit + push to GitHub..." -ForegroundColor Cyan
git add -A
git commit -m $m
if ($LASTEXITCODE -ne 0) {
  Write-Host "(nothing to commit, or commit skipped - continuing to push)" -ForegroundColor Yellow
}
git push
if ($LASTEXITCODE -ne 0) {
  Write-Host "X push FAILED - check internet / GitHub access" -ForegroundColor Red
  exit 1
}

Write-Host ""
Write-Host "==> 3/3 DONE. Render is auto-building (~2-5 min)" -ForegroundColor Green
Write-Host "    status: https://dashboard.render.com" -ForegroundColor Green
Write-Host "    site:   https://trip-fuel-summarizer.onrender.com" -ForegroundColor Green
Write-Host ""
