# 文件夹监控脚本 - 监听文件夹变化并自动执行格式转换脚本

# 监控配置
$watchPath = "D:\Videos"
$pollIntervalMinutes = 5  # 轮询间隔（分钟），作为 FileSystemWatcher 的备用机制

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
Write-Host "📂 监控路径:" -ForegroundColor Green
Write-Host "   $watchPath" -ForegroundColor White
Write-Host ""
Write-Host "功能:" -ForegroundColor Cyan
Write-Host "   • 视频格式转换 (MP4+H.264)" -ForegroundColor Gray
Write-Host "   • 字幕格式转换 (VTT/ASS/SSA/SUB → SRT)" -ForegroundColor Gray
Write-Host "   • NSFW 内容检测和自动分类" -ForegroundColor Gray
Write-Host "   • NSFW → \\192.168.1.111\data\Scenes" -ForegroundColor Gray
Write-Host "   • 普通 → \\192.168.1.111\data\Movies" -ForegroundColor Gray
Write-Host ""
Write-Host "支持格式: TS, AVI, MKV, MOV, WMV, FLV, WEBM, MP4, VTT, ASS, SSA, SUB, SRT等" -ForegroundColor Gray
Write-Host "按 Ctrl+C 停止监控" -ForegroundColor Yellow
Write-Host ""

# 网络目标路径
$networkPathNSFW = "\\192.168.1.111\data\Scenes"    # NSFW 内容
$networkPathSafe = "\\192.168.1.111\data\Movies"    # 普通内容
$nsfwDetectScript = "D:\Soft\Scripts\nsfw_detect.py"

# 磁盘空间检查配置
$linuxHost = "192.168.1.111"
$linuxDataPath = "/mnt/data"
$minimumFreeSpaceGB = 10  # 最小保留空间 (GB)
$script:diskSpaceWarningShown = $false  # 磁盘空间警告是否已显示

# Windows 通知函数
function Send-ToastNotification {
    param(
        [string]$Title,
        [string]$Message,
        [string]$Type = "Warning"  # Warning, Error, Info
    )
    
    try {
        [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
        [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] | Out-Null
        
        $template = @"
<toast>
    <visual>
        <binding template="ToastText02">
            <text id="1">$Title</text>
            <text id="2">$Message</text>
        </binding>
    </visual>
    <audio src="ms-winsoundevent:Notification.Default"/>
</toast>
"@
        
        $xml = New-Object Windows.Data.Xml.Dom.XmlDocument
        $xml.LoadXml($template)
        $toast = New-Object Windows.UI.Notifications.ToastNotification $xml
        [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier("Watch_Downloads").Show($toast)
    } catch {
        # 通知失败时使用系统气泡
        Add-Type -AssemblyName System.Windows.Forms
        $balloon = New-Object System.Windows.Forms.NotifyIcon
        $balloon.Icon = [System.Drawing.SystemIcons]::Warning
        $balloon.BalloonTipIcon = $Type
        $balloon.BalloonTipTitle = $Title
        $balloon.BalloonTipText = $Message
        $balloon.Visible = $true
        $balloon.ShowBalloonTip(10000)
        Start-Sleep -Seconds 1
        $balloon.Dispose()
    }
}

# 检查 Linux 目标磁盘剩余空间
function Test-LinuxDiskSpace {
    param(
        [long]$RequiredBytes = 0
    )
    
    try {
        $result = ssh root@$linuxHost "df -B1 $linuxDataPath | tail -1 | awk '{print `$4}'"
        $availableBytes = [long]$result
        $availableGB = [math]::Round($availableBytes / 1GB, 2)
        $requiredGB = [math]::Round($RequiredBytes / 1GB, 2)
        $minRequired = ($minimumFreeSpaceGB * 1GB) + $RequiredBytes
        
        if ($availableBytes -lt $minRequired) {
            $msg = "Linux 磁盘空间不足! 剩余: ${availableGB}GB"
            Write-Host "  ⚠️ $msg" -ForegroundColor Red
            # 只在首次发现空间不足时弹出通知
            if (-not $script:diskSpaceWarningShown) {
                Send-ToastNotification -Title "磁盘空间警告" -Message $msg -Type "Warning"
                $script:diskSpaceWarningShown = $true
            }
            return $false
        }
        # 空间恢复正常时重置标记，下次空间不足时可以再次通知
        if ($script:diskSpaceWarningShown) {
            $script:diskSpaceWarningShown = $false
            Write-Host "  ✅ Linux 磁盘空间已恢复: ${availableGB}GB" -ForegroundColor Green
        }
        return $true
    } catch {
        Write-Host "  ⚠️ 无法检查 Linux 磁盘空间: $($_.Exception.Message)" -ForegroundColor Yellow
        return $true  # 检查失败时默认允许继续
    }
}

# 全局变量：跟踪上次轮询时间
$script:lastPollTime = Get-Date

# 定义文件处理函数
function Invoke-MediaFileProcessing {
    param(
        [string]$WatchPath,
        [string]$ConvertScript,
        [string]$NetworkPath,
        [bool]$Silent = $false
    )
    
    $files = Get-ChildItem -Path $WatchPath -File -ErrorAction SilentlyContinue
    if (-not $files) { return }
    
    $hasWork = $false
    
    foreach ($file in $files) {
        $name = $file.Name
        $ext = $file.Extension.ToLower()
        
        # 忽略临时文件
        if ($name -match '\.(tmp|partial|!qB|crdownload)$') { continue }
        
        # 处理 MP4 和 SRT 文件 - 进行 NSFW 检测后移动
        if ($ext -eq '.srt' -or $ext -eq '.mp4') {
            if (-not $Silent) {
                Write-Host "[$(Get-Date -Format 'HH:mm:ss')] [轮询] 发现文件: $name" -ForegroundColor Cyan
            }
            Move-MediaFileWithNSFWDetection -FileName $name -SourcePath $WatchPath
            $hasWork = $true
            continue
        }
        
        # 处理需要转换的视频和字幕文件
        $isVideoFile = $ext -match '^\.(ts|avi|mkv|mov|wmv|flv|webm|m4v|3gp|mpg|mpeg|ogv|asf|rm|rmvb)$'
        $isSubtitleFile = $ext -match '^\.(vtt|ass|ssa|sub|sbv)$'
        
        if ($isVideoFile -or $isSubtitleFile) {
            if (-not $Silent) {
                Write-Host "[$(Get-Date -Format 'HH:mm:ss')] [轮询] 发现需转换文件: $name" -ForegroundColor Yellow
            }
            try {
                Push-Location $WatchPath
                & $ConvertScript -NonInteractive
                Pop-Location
                Write-Host "✅ 转换完成" -ForegroundColor Green
            } catch {
                Write-Host "❌ 错误: $_" -ForegroundColor Red
                Pop-Location -ErrorAction SilentlyContinue
            }
            $hasWork = $true
        }
    }
    
    return $hasWork
}

function Move-MediaFile {
    param(
        [string]$FileName,
        [string]$SourcePath,
        [string]$DestPath
    )
    
    $sourceFile = Join-Path $SourcePath $FileName
    $destinationFile = Join-Path $DestPath $FileName
    
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
        
        # 检查 Linux 磁盘空间
        $fileSize = (Get-Item -LiteralPath $sourceFile).Length
        if (-not (Test-LinuxDiskSpace -RequiredBytes $fileSize)) {
            Write-Host "  ⏸️ 跳过移动，等待磁盘空间释放" -ForegroundColor Yellow
            return $false
        }
        
        Move-Item -LiteralPath $sourceFile -Destination $destinationFile -Force
        Write-Host "  ✅ 已移动" -ForegroundColor Green
        return $true
    } catch {
        Write-Host "  ❌ 处理失败: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# NSFW 检测函数
function Test-NSFWContent {
    param(
        [string]$VideoPath
    )
    
    Write-Host "  🔍 正在进行 NSFW 检测..." -ForegroundColor Yellow
    
    try {
        # 调用 Python NSFW 检测脚本
        $result = python $nsfwDetectScript $VideoPath 2>&1
        $exitCode = $LASTEXITCODE
        
        # 解析 JSON 结果
        try {
            $jsonResult = $result | ConvertFrom-Json
            
            if ($jsonResult.is_nsfw) {
                Write-Host "  🔞 检测结果: NSFW (置信度: $($jsonResult.max_score))" -ForegroundColor Magenta
                return $true
            } else {
                Write-Host "  ✅ 检测结果: 普通内容" -ForegroundColor Green
                return $false
            }
        } catch {
            # 如果 JSON 解析失败，根据退出码判断
            if ($exitCode -eq 1) {
                Write-Host "  🔞 检测结果: NSFW" -ForegroundColor Magenta
                return $true
            } else {
                Write-Host "  ✅ 检测结果: 普通内容" -ForegroundColor Green
                return $false
            }
        }
    } catch {
        Write-Host "  ⚠️ NSFW 检测失败，默认归类为普通内容: $($_.Exception.Message)" -ForegroundColor Yellow
        return $false
    }
}

# 智能移动函数（带 NSFW 检测）
function Move-MediaFileWithNSFWDetection {
    param(
        [string]$FileName,
        [string]$SourcePath
    )
    
    $sourceFile = Join-Path $SourcePath $FileName
    $ext = [System.IO.Path]::GetExtension($FileName).ToLower()
    
    # 只对 MP4 视频进行 NSFW 检测
    if ($ext -eq '.mp4') {
        $isNSFW = Test-NSFWContent -VideoPath $sourceFile
        $destPath = if ($isNSFW) { $networkPathNSFW } else { $networkPathSafe }
        $categoryLabel = if ($isNSFW) { "Scenes (NSFW)" } else { "Movies (普通)" }
    } else {
        # SRT 字幕文件：查找对应的 MP4 文件的位置
        $mp4Name = [System.IO.Path]::ChangeExtension($FileName, ".mp4")
        $mp4InNSFW = Join-Path $networkPathNSFW $mp4Name
        $mp4InSafe = Join-Path $networkPathSafe $mp4Name
        
        if (Test-Path $mp4InNSFW) {
            $destPath = $networkPathNSFW
            $categoryLabel = "Scenes (跟随视频)"
        } elseif (Test-Path $mp4InSafe) {
            $destPath = $networkPathSafe
            $categoryLabel = "Movies (跟随视频)"
        } else {
            # 没有找到对应视频，默认放到普通目录
            $destPath = $networkPathSafe
            $categoryLabel = "Movies (默认)"
        }
    }
    
    Write-Host "  📁 目标: $categoryLabel" -ForegroundColor Cyan
    return Move-MediaFile -FileName $FileName -SourcePath $SourcePath -DestPath $destPath
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
        if (Move-MediaFileWithNSFWDetection -FileName $file.Name -SourcePath $watchPath) {
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
    NetworkPathNSFW = $networkPathNSFW
    NetworkPathSafe = $networkPathSafe
    NsfwDetectScript = $nsfwDetectScript
} -Action {
    $name = $Event.SourceEventArgs.Name
    $watchPath = $Event.MessageData.WatchPath
    $convertScript = $Event.MessageData.ConvertScript
    $networkPathNSFW = $Event.MessageData.NetworkPathNSFW
    $networkPathSafe = $Event.MessageData.NetworkPathSafe
    $nsfwDetectScript = $Event.MessageData.NsfwDetectScript
    
    # 忽略脚本本身和临时文件
    if ($name -match '\.(tmp|partial|!qB|crdownload)' -or 
        $name -match 'Convert_to_Mp4_Srt|Watch_Downloads|Convert_Subtitle_to_Srt') {
        return
    }
    
    # 获取文件扩展名
    $ext = [System.IO.Path]::GetExtension($name).ToLower()
    
    # 处理MP4和SRT文件 - 进行 NSFW 检测后移动到对应位置
    if ($ext -eq '.srt' -or $ext -eq '.mp4') {
        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] 检测到 $($ext.ToUpper()) 文件: $name" -ForegroundColor Cyan
        Start-Sleep -Seconds 2  # 等待文件完全写入
        
        $sourceFile = Join-Path $watchPath $name
        
        # 确定目标路径
        if ($ext -eq '.mp4') {
            # 对 MP4 视频进行 NSFW 检测
            Write-Host "  🔍 正在进行 NSFW 检测..." -ForegroundColor Yellow
            try {
                $result = python $nsfwDetectScript $sourceFile 2>&1
                $exitCode = $LASTEXITCODE
                
                try {
                    $jsonResult = $result | ConvertFrom-Json
                    $isNSFW = $jsonResult.is_nsfw
                    if ($isNSFW) {
                        Write-Host "  🔞 检测结果: NSFW (置信度: $($jsonResult.max_score))" -ForegroundColor Magenta
                    } else {
                        Write-Host "  ✅ 检测结果: 普通内容" -ForegroundColor Green
                    }
                } catch {
                    $isNSFW = ($exitCode -eq 1)
                    if ($isNSFW) {
                        Write-Host "  🔞 检测结果: NSFW" -ForegroundColor Magenta
                    } else {
                        Write-Host "  ✅ 检测结果: 普通内容" -ForegroundColor Green
                    }
                }
            } catch {
                Write-Host "  ⚠️ NSFW 检测失败，默认归类为普通内容" -ForegroundColor Yellow
                $isNSFW = $false
            }
            
            $destPath = if ($isNSFW) { $networkPathNSFW } else { $networkPathSafe }
            $categoryLabel = if ($isNSFW) { "Scenes (NSFW)" } else { "Movies (普通)" }
        } else {
            # SRT 字幕文件：查找对应的 MP4 文件的位置
            $mp4Name = [System.IO.Path]::ChangeExtension($name, ".mp4")
            $mp4InNSFW = Join-Path $networkPathNSFW $mp4Name
            $mp4InSafe = Join-Path $networkPathSafe $mp4Name
            
            if (Test-Path $mp4InNSFW) {
                $destPath = $networkPathNSFW
                $categoryLabel = "Scenes (跟随视频)"
            } elseif (Test-Path $mp4InSafe) {
                $destPath = $networkPathSafe
                $categoryLabel = "Movies (跟随视频)"
            } else {
                $destPath = $networkPathSafe
                $categoryLabel = "Movies (默认)"
            }
        }
        
        Write-Host "  📁 目标: $categoryLabel" -ForegroundColor Cyan
        
        try {
            $destinationFile = Join-Path $destPath $name
            
            if (-not (Test-Path $destPath)) {
                Write-Host "❌ 无法访问网络路径: $destPath" -ForegroundColor Red
                return
            }
            
            if (Test-Path -LiteralPath $destinationFile) {
                $sourceSize = (Get-Item -LiteralPath $sourceFile).Length
                $destSize = (Get-Item -LiteralPath $destinationFile).Length
                
                if ($sourceSize -gt $destSize) {
                    Move-Item -LiteralPath $sourceFile -Destination $destinationFile -Force
                    Write-Host "  ✅ 已覆盖 (源: $([math]::Round($sourceSize/1MB,2))MB > 目标: $([math]::Round($destSize/1MB,2))MB)" -ForegroundColor Green
                } else {
                    Remove-Item -LiteralPath $sourceFile -Force
                    Write-Host "  ✅ 已删除源文件 (源: $([math]::Round($sourceSize/1MB,2))MB <= 目标: $([math]::Round($destSize/1MB,2))MB)" -ForegroundColor Green
                }
                return
            }
            
            Move-Item -LiteralPath $sourceFile -Destination $destinationFile -Force
            Write-Host "  ✅ 已移动" -ForegroundColor Green
        } catch {
            Write-Host "  ❌ 移动失败: $name - $($_.Exception.Message)" -ForegroundColor Red
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

# 保持脚本运行，同时定期轮询作为备用
Write-Host "轮询间隔: 每 $pollIntervalMinutes 分钟" -ForegroundColor Gray
Write-Host ""

try {
    while ($true) {
        Start-Sleep -Seconds 10
        
        # 检查是否到达轮询时间
        $now = Get-Date
        $elapsed = ($now - $script:lastPollTime).TotalMinutes
        
        if ($elapsed -ge $pollIntervalMinutes) {
            $script:lastPollTime = $now
            
            # 执行轮询扫描
            $hasWork = Invoke-MediaFileProcessing -WatchPath $watchPath -ConvertScript $convertScriptPath -Silent $false
            
            if (-not $hasWork) {
                Write-Host "[$(Get-Date -Format 'HH:mm:ss')] [轮询] 无待处理文件" -ForegroundColor DarkGray
            }
        }
    }
} finally {
    # 清理监控器
    $watcher.EnableRaisingEvents = $false
    $watcher.Dispose()
    Unregister-Event -SourceIdentifier $onCreated.Name -ErrorAction SilentlyContinue
    
    Write-Host "`n监控已停止" -ForegroundColor Yellow
}
