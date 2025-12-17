@echo off
:: 检查管理员权限
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo 请以管理员身份运行此脚本！
    pause
    exit /b 1
)

:: 创建计划任务
schtasks /create /tn "EnableUWPLoopback" /tr "PowerShell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File \"%~dp0EnableUWPLoopback.ps1\"" /sc onlogon /rl highest /f

if %errorlevel% equ 0 (
    echo.
    echo 计划任务创建成功！每次登录时会自动运行。
    echo.
    echo 现在立即运行一次脚本...
    PowerShell.exe -ExecutionPolicy Bypass -File "%~dp0EnableUWPLoopback.ps1"
) else (
    echo 创建计划任务失败！
)

pause
