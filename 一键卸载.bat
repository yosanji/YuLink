@echo off
chcp 65001 >nul
title YuLink - PowerPoint & WPS 插件一键卸载程序

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0uninstall.ps1"

echo.
echo 按任意键退出...
pause >nul
