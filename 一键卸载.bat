@echo off
chcp 65001 >nul
title YuLink - PowerPoint & WPS 插件一键卸载程序
echo ========================================================
echo                 YuLink 幻灯片网页嵌入插件
echo              作者: yosanji (鱼玄机)   版本: 1.0.0
echo ========================================================
echo.
echo 正在清理 YuLink 注册表与信任配置...
echo.

powershell -NoProfile -ExecutionPolicy Bypass -Command "
$ErrorActionPreference = 'SilentlyContinue'

# 1. Clean PowerPoint AddIn registry
$pptReg = 'HKCU:\Software\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn'
if (Test-Path $pptReg) {
    Remove-Item $pptReg -Recurse -Force
    Write-Host '[OK] 已移除 PowerPoint 插件注册项。'
} else {
    Write-Host '[-] PowerPoint 插件注册项不存在。'
}

# 2. Clean WPS WPP registry
$wpsReg = 'HKCU:\Software\Kingsoft\Office\WPP\AddinsWL'
if (Test-Path $wpsReg) {
    Remove-ItemProperty -Path $wpsReg -Name 'PowerPointAddIn' -Force
    Write-Host '[OK] 已移除 WPS 演示插件注册项。'
} else {
    Write-Host '[-] WPS 演示插件注册项不存在。'
}

# 3. Clean VSTO Inclusion trust entries
$inclusionRoot = 'HKCU:\Software\Microsoft\VSTO\Security\Inclusion'
if (Test-Path $inclusionRoot) {
    Get-ChildItem $inclusionRoot | ForEach-Object {
        $urlVal = Get-ItemProperty -Path $_.PSPath -Name 'Url'
        if ($urlVal -and ($urlVal.Url -like '*PowerPointAddIn.vsto*')) {
            Remove-Item $_.PSPath -Recurse -Force
            Write-Host ('[OK] 已清理 ClickOnce 信任项: ' + $_.PSChildName)
        }
    }
}

Write-Host ''
Write-Host '=============================================='
Write-Host '卸载完成！YuLink 已从您的系统中完全清理。'
Write-Host '=============================================='
"

echo.
echo 按任意键退出...
pause >nul
