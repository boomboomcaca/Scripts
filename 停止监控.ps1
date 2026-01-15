# 停止监控脚本并发送通知

# 通知函数 - 使用 BurntToast 模块或系统气泡
function Send-ToastNotification {
    param(
        [string]$Title,
        [string]$Message,
        [string]$Type = "Warning"
    )
    
    # 使用系统气泡通知（更可靠）
    try {
        Add-Type -AssemblyName System.Windows.Forms
        $balloon = New-Object System.Windows.Forms.NotifyIcon
        $balloon.Icon = [System.Drawing.SystemIcons]::Warning
        $balloon.BalloonTipIcon = $Type
        $balloon.BalloonTipTitle = $Title
        $balloon.BalloonTipText = $Message
        $balloon.Visible = $true
        $balloon.ShowBalloonTip(5000)
        Start-Sleep -Milliseconds 500
        $balloon.Dispose()
    } catch {
        Write-Host "通知发送失败: $_" -ForegroundColor Yellow
    }
}

# 查找并停止监控进程
$found = $false
Get-Process -Name pwsh, powershell -ErrorAction SilentlyContinue | ForEach-Object {
    $cmdLine = (Get-CimInstance Win32_Process -Filter "ProcessId = $($_.Id)").CommandLine
    if ($cmdLine -match 'Watch_Downloads') {
        $found = $true
        Write-Host "正在停止监控脚本 (PID: $($_.Id))..." -ForegroundColor Yellow
        
        # 先发送通知
        Send-ToastNotification -Title "监控脚本已退出" -Message "文件夹监控已停止运行" -Type "Warning"
        
        # 等待通知显示
        Start-Sleep -Milliseconds 500
        
        # 终止进程
        Stop-Process -Id $_.Id -Force
        Write-Host "已停止" -ForegroundColor Green
    }
}

if (-not $found) {
    Write-Host "监控脚本未在运行" -ForegroundColor Gray
}
