' 2026-09-17 无窗口启动 hostwatch 常驻监控。
' 与 StartRecap.vbs 同一套路：计划任务里直接调 powershell/python 加 -WindowStyle Hidden
' 在 Win11 上无效，因为默认终端是 Windows 终端，它自己建窗口。窗口一露头就可能被
' 顺手关掉，而关窗口会杀掉进程。wscript 是 GUI 程序，Run 的第 2 个参数 0 才是真隐藏。
'
' 用 miniconda 的绝对路径，不用裸 "python"：后者会解析到 Windows Store 的
' Python Manager 转发层，它再 spawn 一个真正的 python 子进程，于是进程列表里出现
' 两个 hostwatch，看着像重复启动。（2026-09-17 我就是这么误判并杀错了父进程。）
'
' --quiet：只写日志不打印。日志在 %LOCALAPPDATA%\hostwatch\，按月分文件，
' 只在连接失败/恢复时记录。查看：hostwatch.py history --days 7
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "C:\miniconda3\python.exe -u ""D:\Repos\Scripts\hostwatch.py"" run --quiet", 0, False
