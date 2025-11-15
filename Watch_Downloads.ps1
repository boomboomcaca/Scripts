# 文件夹监控脚本 - 监听文件夹变化并自动执行格式转换脚本

# 监控配置
$watchPath = "C:\Users\Joker\Music"

# 检查脚本文件是否存在
$convertScriptPath = "D:\Soft\Scripts\Convert_to_Mp4_Srt.ps1"

if (Test-Path $convertScriptPath) {
    Write-Host "✅ 找到转换脚本: $convertScriptPath" -ForegroundColor Green
} else {
    Write-Host "❌ 错误: 未找到 Convert_to_Mp4_Srt.ps1" -ForegroundColor Red
    exit 1
}

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "   文件夹监控已启动" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "� 监控路径:" -ForegroundColor Green
Write-Host "   $watchPath" -ForegroundColor White
Write-Host ""
Write-Host "功能:" -ForegroundColor Cyan
Write-Host "   • 视频格式转换 (MP4+H.264)" -ForegroundColor Gray
Write-Host "   • 字幕格式转换 (VTT/ASS/SSA/SUB → SRT)" -ForegroundColor Gray
Write-Host ""
Write-Host "支持格式: TS, AVI, MKV, MOV, WMV, FLV, WEBM, VTT, ASS, SSA, SUB等" -ForegroundColor Gray
Write-Host "按 Ctrl+C 停止监控" -ForegroundColor Yellow
Write-Host ""

# 创建文件监控器
$watcher = New-Object System.IO.FileSystemWatcher
$watcher.Path = $watchPath
$watcher.Filter = "*.*"
$watcher.IncludeSubdirectories = $false
$watcher.NotifyFilter = [System.IO.NotifyFilters]::FileName -bor 
                        [System.IO.NotifyFilters]::LastWrite -bor
                        [System.IO.NotifyFilters]::CreationTime
$watcher.EnableRaisingEvents = $true

# 文件创建事件处理
$onCreated = Register-ObjectEvent -InputObject $watcher -EventName "Created" -MessageData @{
    WatchPath = $watchPath
    ConvertScript = $convertScriptPath
} -Action {
    $name = $Event.SourceEventArgs.Name
    $watchPath = $Event.MessageData.WatchPath
    $convertScript = $Event.MessageData.ConvertScript
    
    # 忽略脚本本身和临时文件
    if ($name -match '\.(tmp|partial|!qB|crdownload)' -or 
        $name -match 'Convert_to_Mp4_Srt|Watch_Downloads|Convert_Subtitle_to_Srt') {
        return
    }
    
    # 获取文件扩展名
    $ext = [System.IO.Path]::GetExtension($name).ToLower()
    
    # 忽略SRT和MP4文件
    if ($ext -eq '.srt' -or $ext -eq '.mp4') {
        return
    }
    
    # 只处理视频和字幕文件
    $isVideoFile = $ext -match '\.(ts|avi|mkv|mov|wmv|flv|webm|m4v|3gp|mpg|mpeg|ogv|asf|rm|rmvb)$'
    $isSubtitleFile = $ext -match '\.(vtt|ass|ssa|sub|sbv)$'
    
    if (-not ($isVideoFile -or $isSubtitleFile)) {
        return
    }
    
    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] 检测到新文件: $name" -ForegroundColor Yellow
    Start-Sleep -Seconds 2
    
    try {
        Push-Location $watchPath
        & $convertScript -NonInteractive
        Pop-Location
        Write-Host "✅ 转换完成" -ForegroundColor Green
    } catch {
        Write-Host "❌ 错误: $_" -ForegroundColor Red
        Pop-Location -ErrorAction SilentlyContinue
    }
}

Write-Host "监控已启动，等待文件变化..." -ForegroundColor Green
Write-Host ""

# 保持脚本运行
try {
    while ($true) {
        Start-Sleep -Seconds 1
    }
} finally {
    # 清理监控器
    $watcher.EnableRaisingEvents = $false
    $watcher.Dispose()
    Unregister-Event -SourceIdentifier $onCreated.Name -ErrorAction SilentlyContinue
    
    Write-Host "`n监控已停止" -ForegroundColor Yellow
}
