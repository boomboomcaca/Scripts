# 文件夹监控脚本 - 监听文件夹变化并自动执行格式转换脚本

# 监控配置
$watchPath = "C:\Users\Joker\Downloads"

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
Write-Host "   • 自动移动MP4和SRT文件到网络位置 (\\192.168.1.111\data\Scenes)" -ForegroundColor Gray
Write-Host ""
Write-Host "支持格式: TS, AVI, MKV, MOV, WMV, FLV, WEBM, MP4, VTT, ASS, SSA, SUB, SRT等" -ForegroundColor Gray
Write-Host "按 Ctrl+C 停止监控" -ForegroundColor Yellow
Write-Host ""

# 网络目标路径
$networkPath = "\\192.168.1.111\data\Scenes"

# 定义文件处理函数
function Process-MediaFile {
    param(
        [string]$FileName,
        [string]$SourcePath,
        [string]$DestPath
    )
    
    $sourceFile = Join-Path $SourcePath $FileName
    $destinationFile = Join-Path $DestPath $FileName
    
    # 转义路径中的方括号
    $sourceFileLiteral = $sourceFile -replace '(\[|\])', '`$1'
    $destFileLiteral = $destinationFile -replace '(\[|\])', '`$1'
    
    try {
        if (-not (Test-Path $DestPath)) {
            Write-Host "❌ 无法访问网络路径: $DestPath" -ForegroundColor Red
            return $false
        }
        
        if (Test-Path -LiteralPath $destinationFile) {
            $sourceSize = (Get-Item -LiteralPath $sourceFile).Length
            $destSize = (Get-Item -LiteralPath $destinationFile).Length
            
            if ($sourceSize -gt $destSize) {
                Move-Item -LiteralPath $sourceFile -Destination $destinationFile -Force
                Write-Host "  ✅ 已覆盖 (源: $([math]::Round($sourceSize/1MB,2))MB > 目标: $([math]::Round($destSize/1MB,2))MB)" -ForegroundColor Green
                return $true
            } else {
                Remove-Item -LiteralPath $sourceFile -Force
                Write-Host "  ✅ 已删除源文件 (源: $([math]::Round($sourceSize/1MB,2))MB <= 目标: $([math]::Round($destSize/1MB,2))MB)" -ForegroundColor Green
                return $true
            }
        }
        
        Move-Item -LiteralPath $sourceFile -Destination $destinationFile -Force
        Write-Host "  ✅ 已移动" -ForegroundColor Green
        return $true
    } catch {
        Write-Host "  ❌ 处理失败: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# 初始化：处理已存在的 MP4 和 SRT 文件
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "   正在扫描已存在的文件..." -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""

$existingFiles = Get-ChildItem -Path $watchPath -File | Where-Object { $_.Extension -eq '.mp4' -or $_.Extension -eq '.srt' }
if ($existingFiles.Count -gt 0) {
    Write-Host "找到 $($existingFiles.Count) 个文件需要处理" -ForegroundColor Yellow
    Write-Host ""
    
    $processedCount = 0
    foreach ($file in $existingFiles) {
        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] 处理: $($file.Name)" -ForegroundColor Cyan
        if (Process-MediaFile -FileName $file.Name -SourcePath $watchPath -DestPath $networkPath) {
            $processedCount++
        }
    }
    
    Write-Host ""
    Write-Host "初始化完成：已处理 $processedCount / $($existingFiles.Count) 个文件" -ForegroundColor Green
    Write-Host ""
} else {
    Write-Host "没有找到需要处理的文件" -ForegroundColor Gray
    Write-Host ""
}

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
    NetworkPath = $networkPath
} -Action {
    $name = $Event.SourceEventArgs.Name
    $watchPath = $Event.MessageData.WatchPath
    $convertScript = $Event.MessageData.ConvertScript
    $networkPath = $Event.MessageData.NetworkPath
    
    # 忽略脚本本身和临时文件
    if ($name -match '\.(tmp|partial|!qB|crdownload)' -or 
        $name -match 'Convert_to_Mp4_Srt|Watch_Downloads|Convert_Subtitle_to_Srt') {
        return
    }
    
    # 获取文件扩展名
    $ext = [System.IO.Path]::GetExtension($name).ToLower()
    
    # 处理MP4和SRT文件 - 直接移动到网络位置
    if ($ext -eq '.srt' -or $ext -eq '.mp4') {
        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] 检测到 $($ext.ToUpper()) 文件: $name" -ForegroundColor Cyan
        Start-Sleep -Seconds 2  # 等待文件完全写入
        
        try {
            $sourceFile = Join-Path $watchPath $name
            $destinationFile = Join-Path $networkPath $name
            
            # 检查网络路径是否可访问
            if (-not (Test-Path $networkPath)) {
                Write-Host "❌ 无法访问网络路径: $networkPath" -ForegroundColor Red
                return
            }
            
            # 检查目标文件是否已存在
            if (Test-Path -LiteralPath $destinationFile) {
                $sourceSize = (Get-Item -LiteralPath $sourceFile).Length
                $destSize = (Get-Item -LiteralPath $destinationFile).Length
                
                if ($sourceSize -gt $destSize) {
                    # 源文件更大，覆盖旧文件
                    Write-Host "⚠️  目标文件已存在 (源: $([math]::Round($sourceSize/1MB,2))MB > 目标: $([math]::Round($destSize/1MB,2))MB)，覆盖旧文件" -ForegroundColor Yellow
                    Move-Item -LiteralPath $sourceFile -Destination $destinationFile -Force
                    Write-Host "✅ 已覆盖到网络位置: $name" -ForegroundColor Green
                } else {
                    # 源文件不大于旧文件，删除源文件
                    Write-Host "⚠️  目标文件已存在 (源: $([math]::Round($sourceSize/1MB,2))MB <= 目标: $([math]::Round($destSize/1MB,2))MB)，删除源文件" -ForegroundColor Yellow
                    Remove-Item -LiteralPath $sourceFile -Force
                    Write-Host "✅ 已删除源文件: $name" -ForegroundColor Green
                }
                return
            }
            
            # 移动文件
            Move-Item -LiteralPath $sourceFile -Destination $destinationFile -Force
            Write-Host "✅ 已移动到网络位置: $name" -ForegroundColor Green
        } catch {
            Write-Host "❌ 移动失败: $name - $($_.Exception.Message)" -ForegroundColor Red
        }
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
