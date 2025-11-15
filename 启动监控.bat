@echo off
chcp 65001 >nul
powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File "%~dp0Watch_Downloads.ps1"
