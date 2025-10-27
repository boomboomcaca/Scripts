# 文件夹监控脚本 - 监听 Downloads 文件夹变化并自动执行转换脚本

$watchPath = "C:\Users\Joker\Downloads"

# 检查全局命令是否可用
$scriptExists = $false
try {
    $scriptCommand = Get-Command Convert_to_Mp4_Srt.ps1 -ErrorAction SilentlyContinue
    if ($scriptCommand) {
        $scriptPath = $scriptCommand.Source
        $scriptExists = $true
        Write-Host "✅ 找到全局命令: $scriptPath" -ForegroundColor Green
    }
} catch {
    $scriptExists = $false
}

if (-not $scriptExists) {
    Write-Host "警告: Convert_to_Mp4_Srt.ps1 未找到！" -ForegroundColor Yellow
    Write-Host "请确保 Convert_to_Mp4_Srt.ps1 已添加到系统PATH环境变量中。" -ForegroundColor Yellow
    Write-Host "当前脚本位置: d:\Soft\Scripts\Convert_to_Mp4_Srt.ps1" -ForegroundColor Yellow
    exit 1
}

Write-Host "开始监控文件夹: $watchPath" -ForegroundColor Green
Write-Host "当检测到视频文件或VTT字幕文件时，将自动执行转换" -ForegroundColor Green
Write-Host "支持的输入格式: VTT字幕、TS、AVI、MKV、MOV、WMV、FLV、WEBM等视频文件" -ForegroundColor Cyan
Write-Host "输出文件(SRT、MP4)将自动忽略，避免重复触发" -ForegroundColor Cyan
Write-Host "按 Ctrl+C 停止监控" -ForegroundColor Cyan
Write-Host ""

# 队列机制：确保同时只有一个转换任务在执行，队列中只保留一个待处理任务
$script:isProcessing = $false
$script:pendingTask = $null

# 执行脚本的函数
function Execute-ConversionScript {
    param($fileName)
    
    # 如果正在处理，更新待处理任务
    if ($script:isProcessing) {
        if ($null -eq $script:pendingTask) {
            Write-Host "[$(Get-Date -Format 'HH:mm:ss')] 检测到文件: $fileName (任务已加入队列，等待当前任务完成)" -ForegroundColor Gray
        } else {
            Write-Host "[$(Get-Date -Format 'HH:mm:ss')] 检测到文件: $fileName (替换队列中的任务: $script:pendingTask)" -ForegroundColor Gray
        }
        $script:pendingTask = $fileName
        return
    }
    
    # 标记为正在处理
    $script:isProcessing = $true
    
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] 检测到文件变化！" -ForegroundColor Yellow
    Write-Host "文件名: $fileName" -ForegroundColor Cyan
    Write-Host "开始执行转换脚本..." -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Yellow
    
    try {
        # 切换到监控目录并执行全局命令
        Push-Location $watchPath
        Convert_to_Mp4_Srt.ps1 -NonInteractive
        Pop-Location
        Write-Host "转换脚本执行完成！" -ForegroundColor Green
    } catch {
        Write-Host "错误: 执行转换脚本时出错 - $_" -ForegroundColor Red
        Pop-Location -ErrorAction SilentlyContinue
    }
    
    Write-Host ""
    
    # 标记为处理完成
    $script:isProcessing = $false
    
    # 检查是否有待处理的任务
    if ($null -ne $script:pendingTask) {
        $nextFile = $script:pendingTask
        $script:pendingTask = $null
        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] 队列中有任务，开始处理: $nextFile" -ForegroundColor Magenta
        Start-Sleep -Seconds 1
        Execute-ConversionScript -fileName $nextFile
    }
}

# 创建文件系统监视器
$fileSystemWatcher = New-Object System.IO.FileSystemWatcher
$fileSystemWatcher.Path = $watchPath
$fileSystemWatcher.Filter = "*.*"
$fileSystemWatcher.IncludeSubdirectories = $false
$fileSystemWatcher.NotifyFilter = [System.IO.NotifyFilters]::FileName -bor 
                                   [System.IO.NotifyFilters]::LastWrite -bor
                                   [System.IO.NotifyFilters]::CreationTime
$fileSystemWatcher.EnableRaisingEvents = $true

# 注册事件处理程序
$onCreated = Register-ObjectEvent -InputObject $fileSystemWatcher -EventName "Created" -Action {
    $name = $Event.SourceEventArgs.Name
    $changeType = $Event.SourceEventArgs.ChangeType
    
    # 忽略脚本本身和临时文件
    if ($name -match '\.(tmp|partial|!qB|crdownload)' -or 
        $name -match 'Convert_to_Mp4_Srt|Watch_Downloads' -or
        $name -eq 'Convert_to_Mp4_Srt.ps1' -or 
        $name -eq 'Watch_Downloads.ps1') {
        return
    }
    
    # 获取文件扩展名
    $ext = [System.IO.Path]::GetExtension($name).ToLower()
    
    # 忽略输出文件（脚本生成的文件）
    if ($ext -eq '.srt' -or $ext -eq '.mp4') {
        return
    }
    
    # 只处理输入文件：VTT 字幕或视频文件
    $isVideoFile = $ext -match '\.(ts|avi|mkv|mov|wmv|flv|webm|m4v|3gp|mpg|mpeg|ogv|asf|rm|rmvb)$'
    $isVttFile = $ext -eq '.vtt'
    
    if (-not ($isVideoFile -or $isVttFile)) {
        return
    }
    
    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] 检测到新文件: $name" -ForegroundColor Cyan
    
    # 等待文件写入完成
    Start-Sleep -Seconds 2
    Execute-ConversionScript -fileName $name
}

$onChanged = Register-ObjectEvent -InputObject $fileSystemWatcher -EventName "Changed" -Action {
    $name = $Event.SourceEventArgs.Name
    $changeType = $Event.SourceEventArgs.ChangeType
    
    # 忽略脚本本身和临时文件
    if ($name -match '\.(tmp|partial|!qB|crdownload)' -or 
        $name -match 'Convert_to_Mp4_Srt|Watch_Downloads' -or
        $name -eq 'Convert_to_Mp4_Srt.ps1' -or 
        $name -eq 'Watch_Downloads.ps1') {
        return
    }
    
    # 获取文件扩展名
    $ext = [System.IO.Path]::GetExtension($name).ToLower()
    
    # 忽略输出文件（脚本生成的文件）
    if ($ext -eq '.srt' -or $ext -eq '.mp4') {
        return
    }
    
    # 只处理输入文件：VTT 字幕或视频文件
    $isVideoFile = $ext -match '\.(ts|avi|mkv|mov|wmv|flv|webm|m4v|3gp|mpg|mpeg|ogv|asf|rm|rmvb)$'
    $isVttFile = $ext -eq '.vtt'
    
    if (-not ($isVideoFile -or $isVttFile)) {
        return
    }
    
    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] 检测到文件更改: $name" -ForegroundColor Cyan
    Start-Sleep -Seconds 2
    Execute-ConversionScript -fileName $name
}

Write-Host "监控已启动！等待文件变化..." -ForegroundColor Green
Write-Host ""

# 保持脚本运行
try {
    while ($true) {
        Start-Sleep -Seconds 1
    }
} finally {
    $fileSystemWatcher.EnableRaisingEvents = $false
    $fileSystemWatcher.Dispose()
    Unregister-Event -SourceIdentifier $onCreated.Name -ErrorAction SilentlyContinue
    Unregister-Event -SourceIdentifier $onChanged.Name -ErrorAction SilentlyContinue
    Write-Host "`n监控已停止。" -ForegroundColor Yellow
}

