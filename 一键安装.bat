@echo off
chcp 65001 >nul
title YuLink - PowerPoint & WPS 插件一键安装程序
echo ========================================================
echo                 YuLink 幻灯片网页嵌入插件
echo              作者: yosanji (鱼玄机)   版本: 1.0.0
echo ========================================================
echo.
echo 正在自动检查环境并编译部署插件，请稍候...
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0build_and_deploy.ps1"

if %ERRORLEVEL% EQU 0 (
    echo.
    echo [OK] YuLink 安装并注册成功！
    echo 您现在可以直接打开 Microsoft PowerPoint 或 WPS 演示开始使用了。
) else (
    echo.
    echo [ERROR] 安装过程中遇到问题，请检查上方日志或以管理员身份运行。
)

echo.
echo 按任意键退出...
pause >nul
