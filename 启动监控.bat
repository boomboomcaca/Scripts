@echo off
chcp 65001 >nul
rem 2026-09-14 改用 PowerShell 7：pwsh 默认以 UTF-8 读取 .ps1，
rem 不像 Windows PowerShell 5.1 那样对无 BOM 文件按 GBK 解析而读坏中文。
"C:\Program Files\PowerShell\7\pwsh.exe" -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "%~dp0Watch_Downloads.ps1"
