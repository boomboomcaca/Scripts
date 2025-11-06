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

$watchPath = "C:\Users\Joker\Downloads"

# 检查全局命令是否可用
$convertScriptExists = $false
$whisperScriptExists = $false

try {
    $convertCommand = Get-Command Convert_to_Mp4_Srt.ps1 -ErrorAction SilentlyContinue
    if ($convertCommand) {
        $convertScriptPath = $convertCommand.Source
        $convertScriptExists = $true
        Write-Host "✅ 找到转换脚本: $convertScriptPath" -ForegroundColor Green
    }
} catch {
    $convertScriptExists = $false
}

try {
    $whisperCommand = Get-Command Generate_Srt_From_Mp4.ps1 -ErrorAction SilentlyContinue
    if ($whisperCommand) {
        $whisperScriptPath = $whisperCommand.Source
        $whisperScriptExists = $true
        Write-Host "✅ 找到语音识别脚本: $whisperScriptPath" -ForegroundColor Green
    }
} catch {
    $whisperScriptExists = $false
}

if (-not $convertScriptExists) {
    Write-Host "警告: Convert_to_Mp4_Srt.ps1 未找到！" -ForegroundColor Yellow
    Write-Host "请确保 Convert_to_Mp4_Srt.ps1 已添加到系统PATH环境变量中。" -ForegroundColor Yellow
    Write-Host "当前脚本位置: d:\Soft\Scripts\Convert_to_Mp4_Srt.ps1" -ForegroundColor Yellow
}

if (-not $whisperScriptExists) {
    Write-Host "警告: Generate_Srt_From_Mp4.ps1 未找到！" -ForegroundColor Yellow
    Write-Host "请确保 Generate_Srt_From_Mp4.ps1 已添加到系统PATH环境变量中。" -ForegroundColor Yellow
    Write-Host "当前脚本位置: d:\Soft\Scripts\Generate_Srt_From_Mp4.ps1" -ForegroundColor Yellow
}

if (-not $convertScriptExists -and -not $whisperScriptExists) {
    Write-Host "错误: 未找到任何处理脚本！" -ForegroundColor Red
    exit 1
}

Write-Host "开始监控文件夹: $watchPath" -ForegroundColor Green
Write-Host ""
Write-Host "📹 MP4文件 -> 自动语音识别生成SRT字幕" -ForegroundColor Cyan
Write-Host "🎬 其他视频格式 -> 转换为MP4+H.264编码" -ForegroundColor Cyan
Write-Host "📝 VTT字幕 -> 转换为SRT格式" -ForegroundColor Cyan
Write-Host ""
Write-Host "支持的输入格式: MP4、VTT字幕、TS、AVI、MKV、MOV、WMV、FLV、WEBM等" -ForegroundColor Gray
Write-Host "输出文件(SRT)将自动忽略，避免重复触发" -ForegroundColor Gray
Write-Host "按 Ctrl+C 停止监控" -ForegroundColor Yellow
Write-Host ""

# 执行锁：确保同时只有一个任务在执行
$script:isProcessing = $false
$script:whisperProcessing = $false

# 执行视频格式转换脚本的函数
function Execute-ConversionScript {
    param($fileName)
    
    # 如果正在处理，跳过本次调用（Convert脚本会自动处理整个文件夹）
    if ($script:isProcessing) {
        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] 检测到文件: $fileName (正在处理其他文件，将稍后自动处理)" -ForegroundColor Gray
        return
    }
    
    if (-not $convertScriptExists) {
        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] 跳过: Convert_to_Mp4_Srt.ps1 脚本不可用" -ForegroundColor Yellow
        return
    }
    
    # 标记为正在处理
    $script:isProcessing = $true
    
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] 检测到视频/字幕文件！" -ForegroundColor Yellow
    Write-Host "文件名: $fileName" -ForegroundColor Cyan
    Write-Host "开始执行格式转换脚本..." -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Yellow
    
    try {
        # 切换到监控目录并执行全局命令
        Push-Location $watchPath
        Convert_to_Mp4_Srt.ps1 -NonInteractive
        Pop-Location
        Write-Host "✅ 格式转换脚本执行完成！" -ForegroundColor Green
    } catch {
        Write-Host "❌ 错误: 执行转换脚本时出错 - $_" -ForegroundColor Red
        Pop-Location -ErrorAction SilentlyContinue
    }
    
    Write-Host ""
    
    # 标记为处理完成
    $script:isProcessing = $false
}

# 执行语音识别脚本的函数（针对MP4文件）
function Execute-WhisperScript {
    param($fileName, $fullPath)
    
    # 如果正在处理，跳过本次调用
    if ($script:whisperProcessing) {
        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] 检测到MP4: $fileName (正在处理其他文件，请稍候)" -ForegroundColor Gray
        return
    }
    
    if (-not $whisperScriptExists) {
        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] 跳过: Generate_Srt_From_Mp4.ps1 脚本不可用" -ForegroundColor Yellow
        return
    }
    
    # 检查是否已有字幕文件
    $srtPath = [System.IO.Path]::ChangeExtension($fullPath, "srt")
    if (Test-Path $srtPath) {
        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] 跳过MP4: $fileName (已有字幕文件)" -ForegroundColor Gray
        return
    }
    
    # 标记为正在处理
    $script:whisperProcessing = $true
    
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] 检测到MP4文件！" -ForegroundColor Cyan
    Write-Host "文件名: $fileName" -ForegroundColor White
    Write-Host "开始语音识别生成字幕..." -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    
    try {
        # 切换到监控目录并执行全局命令
        Push-Location $watchPath
        Generate_Srt_From_Mp4.ps1 -NonInteractive
        Pop-Location
        Write-Host "✅ 语音识别脚本执行完成！" -ForegroundColor Green
    } catch {
        Write-Host "❌ 错误: 执行语音识别脚本时出错 - $_" -ForegroundColor Red
        Pop-Location -ErrorAction SilentlyContinue
    }
    
    Write-Host ""
    
    # 标记为处理完成
    $script:whisperProcessing = $false
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
    $fullPath = $Event.SourceEventArgs.FullPath
    $changeType = $Event.SourceEventArgs.ChangeType
    
    # 忽略脚本本身和临时文件
    if ($name -match '\.(tmp|partial|!qB|crdownload)' -or 
        $name -match 'Convert_to_Mp4_Srt|Watch_Downloads|Generate_Srt_From_Mp4' -or
        $name -eq 'Convert_to_Mp4_Srt.ps1' -or 
        $name -eq 'Watch_Downloads.ps1' -or
        $name -eq 'Generate_Srt_From_Mp4.ps1') {
        return
    }
    
    # 获取文件扩展名
    $ext = [System.IO.Path]::GetExtension($name).ToLower()
    
    # 忽略SRT输出文件（脚本生成的文件）
    if ($ext -eq '.srt') {
        return
    }
    
    # MP4文件单独处理 - 进行语音识别
    if ($ext -eq '.mp4') {
        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] 检测到新MP4文件: $name" -ForegroundColor Cyan
        # 等待文件写入完成
        Start-Sleep -Seconds 3
        Execute-WhisperScript -fileName $name -fullPath $fullPath
        return
    }
    
    # 只处理其他输入文件：VTT 字幕或其他视频格式
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
    $fullPath = $Event.SourceEventArgs.FullPath
    $changeType = $Event.SourceEventArgs.ChangeType
    
    # 忽略脚本本身和临时文件
    if ($name -match '\.(tmp|partial|!qB|crdownload)' -or 
        $name -match 'Convert_to_Mp4_Srt|Watch_Downloads|Generate_Srt_From_Mp4' -or
        $name -eq 'Convert_to_Mp4_Srt.ps1' -or 
        $name -eq 'Watch_Downloads.ps1' -or
        $name -eq 'Generate_Srt_From_Mp4.ps1') {
        return
    }
    
    # 获取文件扩展名
    $ext = [System.IO.Path]::GetExtension($name).ToLower()
    
    # 忽略SRT输出文件（脚本生成的文件）
    if ($ext -eq '.srt') {
        return
    }
    
    # MP4文件单独处理 - 进行语音识别
    if ($ext -eq '.mp4') {
        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] 检测到MP4文件更改: $name" -ForegroundColor Cyan
        Start-Sleep -Seconds 3
        Execute-WhisperScript -fileName $name -fullPath $fullPath
        return
    }
    
    # 只处理其他输入文件：VTT 字幕或其他视频格式
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

