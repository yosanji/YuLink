# YuLink - Clean Uninstaller Script
$ErrorActionPreference = 'SilentlyContinue'

Write-Host "========================================================" -ForegroundColor Cyan
Write-Host "                YuLink Office Addin Uninstaller         " -ForegroundColor Cyan
Write-Host "             Author: yosanji (YuXuanJi)   Version: 1.0  " -ForegroundColor Cyan
Write-Host "========================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Cleaning YuLink registry entries and trust certificates..." -ForegroundColor Yellow
Write-Host ""

# 1. Clean PowerPoint AddIn registry
$pptReg = 'HKCU:\Software\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn'
if (Test-Path $pptReg) {
    Remove-Item $pptReg -Recurse -Force
    Write-Host "  [OK] Removed PowerPoint Addin registry entry." -ForegroundColor Green
} else {
    Write-Host "  [-] PowerPoint Addin registry entry not found." -ForegroundColor Gray
}

# 2. Clean WPS WPP registry
$wpsReg = 'HKCU:\Software\Kingsoft\Office\WPP\AddinsWL'
if (Test-Path $wpsReg) {
    Remove-ItemProperty -Path $wpsReg -Name 'PowerPointAddIn' -Force -ErrorAction SilentlyContinue
    Write-Host "  [OK] Removed WPS AddinsWL registry entry." -ForegroundColor Green
}
$wpsRegCustom = 'HKCU:\Software\Kingsoft\Office\WPP\AddinsCustom\PowerPointAddIn'
if (Test-Path $wpsRegCustom) {
    Remove-Item $wpsRegCustom -Recurse -Force -ErrorAction SilentlyContinue
    Write-Host "  [OK] Removed WPS AddinsCustom registry entry." -ForegroundColor Green
}

# 3. Clean VSTO Inclusion trust entries
$inclusionRoot = 'HKCU:\Software\Microsoft\VSTO\Security\Inclusion'
if (Test-Path $inclusionRoot) {
    Get-ChildItem $inclusionRoot -ErrorAction SilentlyContinue | ForEach-Object {
        $urlVal = Get-ItemProperty -Path $_.PSPath -Name 'Url' -ErrorAction SilentlyContinue
        if ($urlVal -and ($urlVal.Url -like '*PowerPointAddIn.vsto*')) {
            Remove-Item $_.PSPath -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host ("  [OK] Cleaned ClickOnce Trust: " + $_.PSChildName) -ForegroundColor Green
        }
    }
}

Write-Host ""
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host "Uninstallation complete! Cleaned from system. " -ForegroundColor Green
Write-Host "==============================================" -ForegroundColor Cyan
