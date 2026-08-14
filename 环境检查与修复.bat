@echo off
chcp 65001 >nul
title YuLink - 运行环境一键检测与自动修复工具

echo ================================================================
echo           YuLink (鱼链) - 运行环境一键检测与自动修复
echo ================================================================
echo.
echo 正在以管理员权限检查并自动补全缺失的运行库依赖...
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0check_and_fix_env.ps1"

echo.
echo 按任意键退出...
pause >nul
