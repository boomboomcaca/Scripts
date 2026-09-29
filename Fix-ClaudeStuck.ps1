# Claude Desktop 退出后提示"已经启动"、无法重启时用这个清理。
#
# 原因：Claude 是 MSIX（微软商店）打包应用，而打包应用派生的子进程会继承包身份。
# Claude Desktop 会为每个会话派生 Claude Code 的 CLI 子进程
# （%APPDATA%\Claude\claude-code\<版本>\claude.exe），这些进程跑在包目录之外，
# 但身份是继承的。只要还有一个带包身份的进程活着，Windows 就认为应用在运行。
# 退出时若某个子进程卡在长任务上（未返回的命令、后台 workflow），它不会跟着主进程退出，
# 于是整个包一直"在运行"，新实例起不来。
#
# 注意不同会话会各自锁定启动时的 claude-code 版本，所以可能同时存在多个版本的子进程，
# 旧版本的那些更容易变成孤儿。
#
# 用法：
#   .\Fix-ClaudeStuck.ps1          只看不动（默认）
#   .\Fix-ClaudeStuck.ps1 -Kill    实际清理
#
# 警告：如果你是在 Claude Desktop 里面跑这个脚本，-Kill 会把当前会话一起杀掉。
#       正确用法是在 Claude 已经退出、但重启失败的情况下，从 PowerShell 或终端运行。

param([switch]$Kill)

$procs = @(Get-CimInstance Win32_Process -Filter "Name='claude.exe'" -ErrorAction SilentlyContinue)

if ($procs.Count -eq 0) {
    Write-Host "没有残留的 claude 进程。" -ForegroundColor Green
    Write-Host "如果仍然起不来，试试："
    Write-Host "  1. 任务管理器里看有没有 Claude 的托盘/后台任务"
    Write-Host "  2. 注销再登录（清掉 MSIX 的包运行状态）"
    Write-Host "  3. 设置 → 应用 → Claude → 高级选项 → 终止"
    exit 0
}

# 主进程 = 父进程不是 claude.exe 的那个
$ids = $procs.ProcessId
$mains = $procs | Where-Object { $ids -notcontains $_.ParentProcessId }

Write-Host "发现 $($procs.Count) 个 claude 进程：" -ForegroundColor Yellow
foreach ($p in $procs) {
    $role = if ($mains.ProcessId -contains $p.ProcessId) { "主进程" }
            elseif ($p.CommandLine -match 'claude-code') { "CodeCLI" }
            elseif ($p.CommandLine -match '--type=([a-z-]+)') { $Matches[1] }
            else { "子进程" }
    # 从命令行里抠出 claude-code 的版本号，便于看出哪些是旧版残留
    $ver = if ($p.CommandLine -match 'claude-code\\([0-9.]+)\\') { $Matches[1] } else { "" }
    # Get-CimInstance 返回的 CreationDate 已经是 DateTime，不用再过
    # ManagementDateTimeConverter（那是 Get-WmiObject 时代的字符串格式才需要的）
    $started = if ($p.CreationDate -is [datetime]) { $p.CreationDate.ToString('MM-dd HH:mm:ss') } else { "?" }
    $age = if ($p.CreationDate -is [datetime]) { "{0,6:N0}分" -f ((Get-Date) - $p.CreationDate).TotalMinutes } else { "" }
    "{0,7}  {1,-16} {2,-9} 父={3,-7} {4} {5}" -f $p.ProcessId, $role, $ver, $p.ParentProcessId, $started, $age | Write-Host
}

# 孤儿：父进程已经不存在了
$allPids = (Get-Process -ErrorAction SilentlyContinue).Id
$orphans = $procs | Where-Object {
    $_.ParentProcessId -ne 0 -and $allPids -notcontains $_.ParentProcessId
}
if ($orphans) {
    Write-Host ""
    Write-Host "其中 $($orphans.Count) 个是孤儿（父进程已消失）—— 这些最可能是卡住的元凶：" -ForegroundColor Red
    $orphans | ForEach-Object { "  PID $($_.ProcessId)" } | Write-Host
}

if (-not $Kill) {
    Write-Host ""
    Write-Host "这是只读模式。确认要清理就加 -Kill 重跑：" -ForegroundColor Cyan
    Write-Host "  .\Fix-ClaudeStuck.ps1 -Kill"
    exit 0
}

Write-Host ""
Write-Host "正在清理..." -ForegroundColor Yellow
# 先杀子进程再杀主进程，避免主进程在退出流程里又拉起新的子进程
$children = $procs | Where-Object { $mains.ProcessId -notcontains $_.ProcessId }
foreach ($p in $children) {
    try { Stop-Process -Id $p.ProcessId -Force -ErrorAction Stop; "  已结束子进程 $($p.ProcessId)" | Write-Host }
    catch { "  跳过 $($p.ProcessId)（已退出）" | Write-Host }
}
Start-Sleep -Seconds 1
foreach ($p in $mains) {
    try { Stop-Process -Id $p.ProcessId -Force -ErrorAction Stop; "  已结束主进程 $($p.ProcessId)" | Write-Host }
    catch { "  跳过 $($p.ProcessId)（已退出）" | Write-Host }
}

Start-Sleep -Seconds 2
$left = @(Get-Process -Name claude -ErrorAction SilentlyContinue)
if ($left.Count -eq 0) {
    Write-Host ""
    Write-Host "清理完成，现在可以重新启动 Claude Desktop。" -ForegroundColor Green
} else {
    Write-Host ""
    Write-Host "仍有 $($left.Count) 个进程没清掉，可能需要管理员权限：" -ForegroundColor Red
    $left | ForEach-Object { "  PID $($_.Id)" } | Write-Host
}
