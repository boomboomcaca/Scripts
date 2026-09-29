' 2026-09-14 改用 PowerShell 7，理由同 启动监控.bat
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run """C:\Program Files\PowerShell\7\pwsh.exe"" -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File ""D:\Repos\Scripts\Watch_Downloads.ps1""", 0, False
