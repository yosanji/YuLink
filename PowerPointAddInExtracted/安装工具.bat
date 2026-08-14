@echo off
chcp 65001 >nul
title YuLink - 一键安装工具

if exist "%~dp0..\一键安装.bat" (
    call "%~dp0..\一键安装.bat"
) else (
    powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0..\build_and_deploy.ps1"
)