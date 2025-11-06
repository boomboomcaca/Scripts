# 文件夹监控脚本 - 监听 Downloads 文件夹变化并自动执行转换脚本

# 自动添加 PyTorch CUDA 库路径到当前会话（修复 cublas64_12.dll 问题）
$findCudaPath = @"
import os, torch
cuda_lib = os.path.join(os.path.dirname(torch.__file__), 'lib')
print(cuda_lib) if os.path.exists(cuda_lib) else print('')
"@

try {
    # 优先使用 Python 3.11
    $cudaLibPath = $null
    try {
        $cudaLibPath = py -3.11 -c $findCudaPath 2>$null
    } catch {}
    
    if (-not $cudaLibPath) {
        $cudaLibPath = python -c $findCudaPath 2>$null
    }
    
    if ($cudaLibPath -and (Test-Path $cudaLibPath)) {
        if ($env:Path -notlike "*$cudaLibPath*") {
            $env:Path = $env:Path + ";" + $cudaLibPath
        }
    }
} catch {
    # 静默失败，不影响后续流程
}

# 双路径监控配置
$watchPathLocal = "C:\Users\Joker\Downloads"      # 本地路径：格式转换
$watchPathNetwork = "\\192.168.1.111\data\Scenes"  # 网络路径：语音识别

# 检查脚本文件是否存在
$convertScriptPath = "D:\Soft\Scripts\Convert_to_Mp4_Srt.ps1"
$whisperScriptPath = "D:\Soft\Scripts\Generate_Srt_From_Mp4.ps1"

if (Test-Path $convertScriptPath) {
    Write-Host "✅ 找到转换脚本: $convertScriptPath" -ForegroundColor Green
} else {
    Write-Host "⚠️  警告: 未找到 Convert_to_Mp4_Srt.ps1" -ForegroundColor Yellow
}

if (Test-Path $whisperScriptPath) {
    Write-Host "✅ 找到语音识别脚本: $whisperScriptPath" -ForegroundColor Green
} else {
    Write-Host "⚠️  警告: 未找到 Generate_Srt_From_Mp4.ps1" -ForegroundColor Yellow
}

if (-not (Test-Path $convertScriptPath) -and -not (Test-Path $whisperScriptPath)) {
    Write-Host "❌ 错误: 未找到任何处理脚本！" -ForegroundColor Red
    exit 1
}

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "   双路径监控已启动" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "📂 监控路径1（本地）:" -ForegroundColor Green
Write-Host "   $watchPathLocal" -ForegroundColor White
Write-Host "   🎬 其他视频格式 -> 转换为MP4+H.264编码" -ForegroundColor Cyan
Write-Host "   📝 VTT字幕 -> 转换为SRT格式" -ForegroundColor Cyan
Write-Host ""
Write-Host "📂 监控路径2（网络）:" -ForegroundColor Green
Write-Host "   $watchPathNetwork" -ForegroundColor White
Write-Host "   📹 MP4文件 -> 自动语音识别生成SRT字幕" -ForegroundColor Cyan
Write-Host ""
Write-Host "支持的输入格式: MP4、VTT字幕、TS、AVI、MKV、MOV、WMV、FLV、WEBM等" -ForegroundColor Gray
Write-Host "按 Ctrl+C 停止监控" -ForegroundColor Yellow
Write-Host ""

# 创建监控器1：本地路径（格式转换）
$watcherLocal = New-Object System.IO.FileSystemWatcher
$watcherLocal.Path = $watchPathLocal
$watcherLocal.Filter = "*.*"
$watcherLocal.IncludeSubdirectories = $false
$watcherLocal.NotifyFilter = [System.IO.NotifyFilters]::FileName -bor 
                              [System.IO.NotifyFilters]::LastWrite -bor
                              [System.IO.NotifyFilters]::CreationTime
$watcherLocal.EnableRaisingEvents = $true

# 创建监控器2：网络路径（语音识别）
$watcherNetwork = New-Object System.IO.FileSystemWatcher
$watcherNetwork.Path = $watchPathNetwork
$watcherNetwork.Filter = "*.*"
$watcherNetwork.IncludeSubdirectories = $false
$watcherNetwork.NotifyFilter = [System.IO.NotifyFilters]::FileName -bor 
                                [System.IO.NotifyFilters]::LastWrite -bor
                                [System.IO.NotifyFilters]::CreationTime
$watcherNetwork.EnableRaisingEvents = $true

# 本地路径事件处理（格式转换）
$onCreatedLocal = Register-ObjectEvent -InputObject $watcherLocal -EventName "Created" -MessageData @{
    WatchPath = $watchPathLocal
    ConvertScript = $convertScriptPath
} -Action {
    $name = $Event.SourceEventArgs.Name
    $watchPath = $Event.MessageData.WatchPath
    $convertScript = $Event.MessageData.ConvertScript
    
    # 忽略脚本本身和临时文件
    if ($name -match '\.(tmp|partial|!qB|crdownload)' -or 
        $name -match 'Convert_to_Mp4_Srt|Watch_Downloads|Generate_Srt_From_Mp4') {
        return
    }
    
    # 获取文件扩展名
    $ext = [System.IO.Path]::GetExtension($name).ToLower()
    
    # 忽略SRT和MP4文件（本地路径不处理MP4）
    if ($ext -eq '.srt' -or $ext -eq '.mp4') {
        return
    }
    
    # 只处理其他输入文件：VTT 字幕或其他视频格式
    $isVideoFile = $ext -match '\.(ts|avi|mkv|mov|wmv|flv|webm|m4v|3gp|mpg|mpeg|ogv|asf|rm|rmvb)$'
    $isVttFile = $ext -eq '.vtt'
    
    if (-not ($isVideoFile -or $isVttFile)) {
        return
    }
    
    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] [本地] 检测到新文件: $name" -ForegroundColor Yellow
    Write-Host "[DEBUG] 文件扩展名: $ext, isVideoFile: $isVideoFile, isVttFile: $isVttFile" -ForegroundColor DarkGray
    
    # 等待文件写入完成
    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] [本地] 等待文件写入完成..." -ForegroundColor Gray
    Start-Sleep -Seconds 2
    
    Write-Host "[DEBUG] 准备执行转换..." -ForegroundColor DarkGray
    
    try {
        Write-Host "========================================" -ForegroundColor Yellow
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [本地] 开始处理文件！" -ForegroundColor Yellow
        Write-Host "文件名: $name" -ForegroundColor Cyan
        Write-Host "路径: $watchPath" -ForegroundColor Cyan
        Write-Host "========================================" -ForegroundColor Yellow
        
        Push-Location $watchPath
        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] 正在调用转换脚本..." -ForegroundColor Cyan
        & $convertScript -NonInteractive
        Pop-Location
        Write-Host "✅ 格式转换脚本执行完成！" -ForegroundColor Green
    } catch {
        Write-Host "❌ 错误: $_" -ForegroundColor Red
        Write-Host "❌ 错误详情: $($_.Exception.Message)" -ForegroundColor Red
        Pop-Location -ErrorAction SilentlyContinue
    }
}

# 网络路径事件处理（语音识别）
$onCreatedNetwork = Register-ObjectEvent -InputObject $watcherNetwork -EventName "Created" -MessageData @{
    WatchPath = $watchPathNetwork
    WhisperScript = $whisperScriptPath
} -Action {
    $name = $Event.SourceEventArgs.Name
    $fullPath = $Event.SourceEventArgs.FullPath
    $watchPath = $Event.MessageData.WatchPath
    $whisperScript = $Event.MessageData.WhisperScript
    
    # 忽略脚本本身和临时文件
    if ($name -match '\.(tmp|partial|!qB|crdownload)' -or 
        $name -match 'Convert_to_Mp4_Srt|Watch_Downloads|Generate_Srt_From_Mp4') {
        return
    }
    
    # 获取文件扩展名
    $ext = [System.IO.Path]::GetExtension($name).ToLower()
    
    # 忽略SRT输出文件
    if ($ext -eq '.srt') {
        return
    }
    
    # 只处理MP4文件
    if ($ext -eq '.mp4') {
        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] [网络] 检测到新MP4文件: $name" -ForegroundColor Cyan
        
        # 检查是否已有字幕文件
        $srtPath = [System.IO.Path]::ChangeExtension($fullPath, "srt")
        if (Test-Path $srtPath) {
            Write-Host "[$(Get-Date -Format 'HH:mm:ss')] [网络] 跳过: 已有字幕文件" -ForegroundColor Gray
            return
        }
        
        # 等待文件写入完成
        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] [网络] 等待文件写入完成..." -ForegroundColor Gray
        Start-Sleep -Seconds 3
        
        Write-Host "[DEBUG] 准备执行语音识别..." -ForegroundColor DarkGray
        
        try {
            Write-Host "========================================" -ForegroundColor Cyan
            Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [网络] 开始语音识别！" -ForegroundColor Cyan
            Write-Host "文件名: $name" -ForegroundColor White
            Write-Host "路径: $watchPath" -ForegroundColor White
            Write-Host "========================================" -ForegroundColor Cyan
            
            Push-Location $watchPath
            Write-Host "[$(Get-Date -Format 'HH:mm:ss')] 正在调用语音识别脚本..." -ForegroundColor Cyan
            & $whisperScript -NonInteractive
            Pop-Location
            Write-Host "✅ 语音识别脚本执行完成！" -ForegroundColor Green
        } catch {
            Write-Host "❌ 错误: $_" -ForegroundColor Red
            Write-Host "❌ 错误详情: $($_.Exception.Message)" -ForegroundColor Red
            Pop-Location -ErrorAction SilentlyContinue
        }
    }
}

$onChangedLocal = Register-ObjectEvent -InputObject $watcherLocal -EventName "Changed" -MessageData @{
    WatchPath = $watchPathLocal
    ConvertScript = $convertScriptPath
} -Action {
    $name = $Event.SourceEventArgs.Name
    $watchPath = $Event.MessageData.WatchPath
    $convertScript = $Event.MessageData.ConvertScript
    
    # 忽略脚本本身和临时文件
    if ($name -match '\.(tmp|partial|!qB|crdownload)' -or 
        $name -match 'Convert_to_Mp4_Srt|Watch_Downloads|Generate_Srt_From_Mp4') {
        return
    }
    
    # 获取文件扩展名
    $ext = [System.IO.Path]::GetExtension($name).ToLower()
    
    # 忽略SRT和MP4文件
    if ($ext -eq '.srt' -or $ext -eq '.mp4') {
        return
    }
    
    # 只处理其他输入文件：VTT 字幕或其他视频格式
    $isVideoFile = $ext -match '\.(ts|avi|mkv|mov|wmv|flv|webm|m4v|3gp|mpg|mpeg|ogv|asf|rm|rmvb)$'
    $isVttFile = $ext -eq '.vtt'
    
    if (-not ($isVideoFile -or $isVttFile)) {
        return
    }
    
    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] [本地] 检测到文件更改: $name" -ForegroundColor Yellow
    Write-Host "[DEBUG] 文件扩展名: $ext, isVideoFile: $isVideoFile, isVttFile: $isVttFile" -ForegroundColor DarkGray
    
    # 等待文件写入完成
    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] [本地] 等待文件写入完成..." -ForegroundColor Gray
    Start-Sleep -Seconds 2
    
    Write-Host "[DEBUG] 准备执行转换..." -ForegroundColor DarkGray
    
    try {
        Write-Host "========================================" -ForegroundColor Yellow
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [本地] 开始处理文件！" -ForegroundColor Yellow
        Write-Host "文件名: $name" -ForegroundColor Cyan
        Write-Host "路径: $watchPath" -ForegroundColor Cyan
        Write-Host "========================================" -ForegroundColor Yellow
        
        Push-Location $watchPath
        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] 正在调用转换脚本..." -ForegroundColor Cyan
        & $convertScript -NonInteractive
        Pop-Location
        Write-Host "✅ 格式转换脚本执行完成！" -ForegroundColor Green
    } catch {
        Write-Host "❌ 错误: $_" -ForegroundColor Red
        Write-Host "❌ 错误详情: $($_.Exception.Message)" -ForegroundColor Red
        Pop-Location -ErrorAction SilentlyContinue
    }
}

$onChangedNetwork = Register-ObjectEvent -InputObject $watcherNetwork -EventName "Changed" -MessageData @{
    WatchPath = $watchPathNetwork
    WhisperScript = $whisperScriptPath
} -Action {
    $name = $Event.SourceEventArgs.Name
    $fullPath = $Event.SourceEventArgs.FullPath
    $watchPath = $Event.MessageData.WatchPath
    $whisperScript = $Event.MessageData.WhisperScript
    
    # 忽略脚本本身和临时文件
    if ($name -match '\.(tmp|partial|!qB|crdownload)' -or 
        $name -match 'Convert_to_Mp4_Srt|Watch_Downloads|Generate_Srt_From_Mp4') {
        return
    }
    
    # 获取文件扩展名
    $ext = [System.IO.Path]::GetExtension($name).ToLower()
    
    # 忽略SRT输出文件
    if ($ext -eq '.srt') {
        return
    }
    
    # 只处理MP4文件
    if ($ext -eq '.mp4') {
        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] [网络] 检测到MP4文件更改: $name" -ForegroundColor Cyan
        
        # 检查是否已有字幕文件
        $srtPath = [System.IO.Path]::ChangeExtension($fullPath, "srt")
        if (Test-Path $srtPath) {
            Write-Host "[$(Get-Date -Format 'HH:mm:ss')] [网络] 跳过: 已有字幕文件" -ForegroundColor Gray
            return
        }
        
        # 等待文件写入完成
        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] [网络] 等待文件写入完成..." -ForegroundColor Gray
        Start-Sleep -Seconds 3
        
        Write-Host "[DEBUG] 准备执行语音识别..." -ForegroundColor DarkGray
        
        try {
            Write-Host "========================================" -ForegroundColor Cyan
            Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [网络] 开始语音识别！" -ForegroundColor Cyan
            Write-Host "文件名: $name" -ForegroundColor White
            Write-Host "路径: $watchPath" -ForegroundColor White
            Write-Host "========================================" -ForegroundColor Cyan
            
            Push-Location $watchPath
            Write-Host "[$(Get-Date -Format 'HH:mm:ss')] 正在调用语音识别脚本..." -ForegroundColor Cyan
            & $whisperScript -NonInteractive
            Pop-Location
            Write-Host "✅ 语音识别脚本执行完成！" -ForegroundColor Green
        } catch {
            Write-Host "❌ 错误: $_" -ForegroundColor Red
            Write-Host "❌ 错误详情: $($_.Exception.Message)" -ForegroundColor Red
            Pop-Location -ErrorAction SilentlyContinue
        }
    }
}

Write-Host "监控已启动！等待文件变化..." -ForegroundColor Green
Write-Host ""

# 保持脚本运行
try {
    while ($true) {
        Start-Sleep -Seconds 1
    }
} finally {
    # 清理本地监控器
    $watcherLocal.EnableRaisingEvents = $false
    $watcherLocal.Dispose()
    Unregister-Event -SourceIdentifier $onCreatedLocal.Name -ErrorAction SilentlyContinue
    Unregister-Event -SourceIdentifier $onChangedLocal.Name -ErrorAction SilentlyContinue
    
    # 清理网络监控器
    $watcherNetwork.EnableRaisingEvents = $false
    $watcherNetwork.Dispose()
    Unregister-Event -SourceIdentifier $onCreatedNetwork.Name -ErrorAction SilentlyContinue
    Unregister-Event -SourceIdentifier $onChangedNetwork.Name -ErrorAction SilentlyContinue
    
    Write-Host "`n监控已停止。" -ForegroundColor Yellow
}

