@echo off
chcp 65001 >nul
echo ====================================
echo   文件夹监控 - 自动转换脚本
echo ====================================
echo.
echo 正在启动监控...
echo.

powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File "%~dp0Watch_Downloads.ps1"

