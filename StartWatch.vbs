Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File ""D:\Repos\Scripts\Watch_Downloads.ps1""", 0, False
