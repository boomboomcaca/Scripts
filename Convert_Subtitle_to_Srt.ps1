# 通用字幕格式转换工具 - 转换所有字幕为SRT格式
# 作者: Claude
# 功能: 将ASS、SSA、VTT、SUB、SBV等字幕格式统一转换为SRT格式

param(
    [string]$Path = ".",
    [switch]$NoDelete,
    [switch]$Help,
    [switch]$NonInteractive,
    [switch]$Recursive
)

# 显示帮助信息
if ($Help) {
    Write-Host @"
通用字幕格式转换工具 - 转换所有字幕为SRT格式

用法:
    .\Convert_Subtitle_to_Srt.ps1 [-Path <目录路径>] [-NoDelete] [-Recursive] [-NonInteractive] [-Help]

参数:
    -Path          指定要处理的目录路径 (默认: 当前目录)
    -NoDelete      保留原始字幕文件，不删除
    -Recursive     递归处理子目录中的字幕文件
    -NonInteractive 非交互模式，不等待按键退出（用于自动化调用）
    -Help          显示此帮助信息

功能:
    1. 检查ffmpeg环境（用于格式转换）
    2. 自动扫描以下字幕格式:
       • VTT (WebVTT) - Web视频字幕
       • ASS (Advanced SubStation) - 高级字幕
       • SSA (SubStation Alpha) - 字幕编辑器格式
       • SUB (MicroDVD/SubViewer) - 通用字幕格式
       • SBV (YouTube字幕格式)
       • DFXP/TTML (时序文本标记语言)
       • LRC (歌词格式)
    3. 智能转换为标准SRT格式
    4. 清理HTML/格式化标签
    5. 修复编码问题（统一为UTF-8）

示例:
    .\Convert_Subtitle_to_Srt.ps1                        # 转换当前目录的所有字幕
    .\Convert_Subtitle_to_Srt.ps1 -Path "D:\Videos"      # 处理指定目录
    .\Convert_Subtitle_to_Srt.ps1 -Recursive             # 递归处理子目录
    .\Convert_Subtitle_to_Srt.ps1 -NoDelete              # 保留原始字幕文件
    .\Convert_Subtitle_to_Srt.ps1 -NonInteractive        # 自动化模式

注意事项:
    - 需要安装ffmpeg并添加到PATH环境变量
    - ASS/SSA格式会移除特效和样式，保留纯文本
    - VTT格式会自动清理HTML标签
    - 默认会删除原始字幕文件（使用-NoDelete保留）
"@
    exit 0
}

# 设置控制台编码为UTF-8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$Host.UI.RawUI.WindowTitle = "通用字幕格式转换工具"

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "   通用字幕格式转换工具 → SRT" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""

# 切换到指定目录
if ($Path -ne ".") {
    if (Test-Path $Path) {
        Set-Location $Path
        Write-Host "📁 处理目录: $Path" -ForegroundColor Yellow
    } else {
        Write-Host "❌ 错误: 目录不存在: $Path" -ForegroundColor Red
        exit 1
    }
} else {
    Write-Host "📁 处理目录: $(Get-Location)" -ForegroundColor Yellow
}

# 检查ffmpeg
Write-Host ""
Write-Host "[1/4] 检查ffmpeg环境..." -ForegroundColor Green
try {
    $ffmpegVersion = ffmpeg -version 2>&1 | Select-Object -First 1
    if ($ffmpegVersion -match "ffmpeg version") {
        Write-Host "✅ 检测到ffmpeg: $ffmpegVersion" -ForegroundColor Green
    } else {
        throw "ffmpeg未正确安装"
    }
} catch {
    Write-Host "❌ 未找到ffmpeg，请确保已安装并添加到PATH" -ForegroundColor Red
    Write-Host "下载地址: https://ffmpeg.org/download.html" -ForegroundColor Yellow
    if (-not $NonInteractive) {
        Read-Host "按任意键退出"
    }
    exit 1
}

# 定义字幕文件扩展名
$SubtitleExtensions = @("*.vtt", "*.ass", "*.ssa", "*.sub", "*.sbv", "*.dfxp", "*.ttml", "*.lrc")

# 获取所有字幕文件
Write-Host ""
Write-Host "[2/4] 扫描字幕文件..." -ForegroundColor Green

$allSubtitleFiles = @()
$scanParams = @{
    Include = $SubtitleExtensions
    File = $true
    ErrorAction = "SilentlyContinue"
}

if ($Recursive) {
    $scanParams.Recurse = $true
    Write-Host "🔍 递归扫描模式 - 包含所有子目录" -ForegroundColor Cyan
}

$allSubtitleFiles = Get-ChildItem @scanParams

if ($allSubtitleFiles.Count -eq 0) {
    Write-Host "⚠️  未找到任何需要转换的字幕文件" -ForegroundColor Yellow
    Write-Host "支持的格式: VTT, ASS, SSA, SUB, SBV, DFXP, TTML, LRC" -ForegroundColor Gray
    if (-not $NonInteractive) {
        Read-Host "按任意键退出"
    }
    exit 0
}

# 按格式分类统计
$formatStats = @{}
foreach ($file in $allSubtitleFiles) {
    $ext = $file.Extension.ToUpper()
    if ($formatStats.ContainsKey($ext)) {
        $formatStats[$ext] += 1
    } else {
        $formatStats[$ext] = 1
    }
}

Write-Host "📊 找到 $($allSubtitleFiles.Count) 个字幕文件:" -ForegroundColor White
foreach ($format in $formatStats.Keys | Sort-Object) {
    $icon = switch ($format) {
        ".VTT" { "🌐" }
        ".ASS" { "🎨" }
        ".SSA" { "🎨" }
        ".SUB" { "📝" }
        ".SBV" { "▶️" }
        ".DFXP" { "📺" }
        ".TTML" { "📺" }
        ".LRC" { "🎵" }
        default { "📄" }
    }
    Write-Host "  $icon $format`: $($formatStats[$format]) 个文件" -ForegroundColor Cyan
}

# 清理HTML和格式标签的函数
function Clear-SubtitleText {
    param([string]$Text)
    
    # 移除HTML标签
    $cleaned = $Text -replace '<[^>]*>', ''
    
    # 移除ASS/SSA格式标签 {\...}
    $cleaned = $cleaned -replace '\{[^}]*\}', ''
    
    # 移除VTT样式标签 <v ...>
    $cleaned = $cleaned -replace '<v[^>]*>', ''
    $cleaned = $cleaned -replace '</v>', ''
    
    # 替换HTML实体
    $cleaned = $cleaned -replace '&amp;', '&'
    $cleaned = $cleaned -replace '&lt;', '<'
    $cleaned = $cleaned -replace '&gt;', '>'
    $cleaned = $cleaned -replace '&quot;', '"'
    $cleaned = $cleaned -replace '&#39;', "'"
    $cleaned = $cleaned -replace '&nbsp;', ' '
    
    # 移除多余的空白
    $cleaned = $cleaned -replace '\s+', ' '
    $cleaned = $cleaned.Trim()
    
    return $cleaned
}

# 手动解析ASS/SSA格式的函数
function Convert-AssToSrt {
    param(
        [string]$InputFile,
        [string]$OutputFile
    )
    
    try {
        $lines = Get-Content $InputFile -Encoding UTF8
        $events = @()
        $inEvents = $false
        
        foreach ($line in $lines) {
            if ($line -match '^\[Events\]') {
                $inEvents = $true
                continue
            }
            
            if ($inEvents -and $line -match '^(Dialogue|Comment):\s*(.+)') {
                $parts = $matches[2] -split ',', 10
                
                if ($parts.Count -ge 10) {
                    $startTime = $parts[1].Trim()
                    $endTime = $parts[2].Trim()
                    $text = $parts[9].Trim()
                    
                    # 转换时间格式 (H:MM:SS.cc -> HH:MM:SS,mmm)
                    $startTime = $startTime -replace '(\d+):(\d+):(\d+)\.(\d+)', {
                        $h = [int]$matches[1]
                        $m = [int]$matches[2]
                        $s = [int]$matches[3]
                        $cs = [int]$matches[4]
                        $ms = $cs * 10
                        "{0:D2}:{1:D2}:{2:D2},{3:D3}" -f $h, $m, $s, $ms
                    }
                    
                    $endTime = $endTime -replace '(\d+):(\d+):(\d+)\.(\d+)', {
                        $h = [int]$matches[1]
                        $m = [int]$matches[2]
                        $s = [int]$matches[3]
                        $cs = [int]$matches[4]
                        $ms = $cs * 10
                        "{0:D2}:{1:D2}:{2:D2},{3:D3}" -f $h, $m, $s, $ms
                    }
                    
                    # 清理文本
                    $text = Clear-SubtitleText -Text $text
                    
                    # 替换换行符
                    $text = $text -replace '\\N', "`n"
                    $text = $text -replace '\\n', "`n"
                    
                    if ($text) {
                        $events += @{
                            Start = $startTime
                            End = $endTime
                            Text = $text
                        }
                    }
                }
            }
        }
        
        if ($events.Count -eq 0) {
            throw "未找到有效的字幕条目"
        }
        
        # 写入SRT文件
        $srtContent = ""
        for ($i = 0; $i -lt $events.Count; $i++) {
            $srtContent += "$($i + 1)`n"
            $srtContent += "$($events[$i].Start) --> $($events[$i].End)`n"
            $srtContent += "$($events[$i].Text)`n"
            $srtContent += "`n"
        }
        
        [System.IO.File]::WriteAllText($OutputFile, $srtContent, [System.Text.Encoding]::UTF8)
        return $true
    }
    catch {
        Write-Host "⚠️  ASS/SSA手动解析失败: $($_.Exception.Message)" -ForegroundColor Yellow
        return $false
    }
}

# 转换字幕文件
Write-Host ""
Write-Host "[3/4] 转换字幕格式..." -ForegroundColor Green
Write-Host ""

$successCount = 0
$failureCount = 0
$skippedCount = 0

foreach ($file in $allSubtitleFiles) {
    $relativePath = if ($Recursive) {
        $file.FullName.Substring((Get-Location).Path.Length + 1)
    } else {
        $file.Name
    }
    
    $outputFile = [System.IO.Path]::ChangeExtension($file.FullName, "srt")
    
    # 检查是否已存在SRT文件
    if (Test-Path $outputFile) {
        Write-Host "⏭️  跳过 (SRT已存在): $relativePath" -ForegroundColor Gray
        $skippedCount++
        continue
    }
    
    Write-Host "🔄 转换: $relativePath" -ForegroundColor White
    
    $converted = $false
    
    # ASS/SSA格式优先使用手动解析
    if ($file.Extension -in @('.ass', '.ssa')) {
        Write-Host "   使用ASS/SSA解析器..." -ForegroundColor Gray
        $converted = Convert-AssToSrt -InputFile $file.FullName -OutputFile $outputFile
    }
    
    # 如果ASS/SSA手动解析失败或其他格式，使用ffmpeg
    if (-not $converted) {
        if ($file.Extension -in @('.ass', '.ssa')) {
            Write-Host "   降级到ffmpeg转换..." -ForegroundColor Gray
        }
        
        try {
            $process = Start-Process -FilePath "ffmpeg" -ArgumentList @(
                "-i", "`"$($file.FullName)`"",
                "-y",
                "`"$outputFile`""
            ) -Wait -PassThru -NoNewWindow -RedirectStandardError "$env:TEMP\ffmpeg_error.txt"
            
            if ($process.ExitCode -eq 0 -and (Test-Path $outputFile)) {
                $converted = $true
            }
        } catch {
            Write-Host "❌ ffmpeg转换失败: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    
    if ($converted) {
        # 后处理：清理HTML标签和格式化
        try {
            $content = Get-Content $outputFile -Raw -Encoding UTF8
            
            # 清理各种标签
            $content = Clear-SubtitleText -Text $content
            
            # 保存清理后的内容
            [System.IO.File]::WriteAllText($outputFile, $content, [System.Text.Encoding]::UTF8)
            
            Write-Host "✅ 成功: $relativePath" -ForegroundColor Green
            $successCount++
            
            # 删除原始文件
            if (-not $NoDelete) {
                Remove-Item $file.FullName -Force
                Write-Host "   删除原文件: $($file.Name)" -ForegroundColor Gray
            }
        } catch {
            Write-Host "⚠️  后处理失败，但文件已转换: $relativePath" -ForegroundColor Yellow
            $successCount++
        }
    } else {
        Write-Host "❌ 失败: $relativePath" -ForegroundColor Red
        $failureCount++
    }
    
    Write-Host ""
}

# 清理现有SRT文件的HTML标签
Write-Host ""
Write-Host "[4/4] 清理现有SRT文件..." -ForegroundColor Green

$srtScanParams = @{
    Filter = "*.srt"
    File = $true
    ErrorAction = "SilentlyContinue"
}

if ($Recursive) {
    $srtScanParams.Recurse = $true
}

$srtFiles = Get-ChildItem @srtScanParams
$cleanedCount = 0

foreach ($file in $srtFiles) {
    $relativePath = if ($Recursive) {
        $file.FullName.Substring((Get-Location).Path.Length + 1)
    } else {
        $file.Name
    }
    
    try {
        $content = Get-Content $file.FullName -Raw -Encoding UTF8
        $originalContent = $content
        
        # 清理标签
        $content = Clear-SubtitleText -Text $content
        
        if ($content -ne $originalContent) {
            [System.IO.File]::WriteAllText($file.FullName, $content, [System.Text.Encoding]::UTF8)
            Write-Host "✅ 清理: $relativePath" -ForegroundColor Green
            $cleanedCount++
        }
    } catch {
        Write-Host "⚠️  清理失败: $relativePath" -ForegroundColor Yellow
    }
}

# 显示最终结果
Write-Host ""
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "              处理完成！" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "🎉 转换结果统计:" -ForegroundColor Green
Write-Host "  ✅ 成功转换: $successCount 个文件" -ForegroundColor Green
if ($skippedCount -gt 0) {
    Write-Host "  ⏭️  跳过: $skippedCount 个文件 (SRT已存在)" -ForegroundColor Gray
}
if ($failureCount -gt 0) {
    Write-Host "  ❌ 转换失败: $failureCount 个文件" -ForegroundColor Red
}
if ($cleanedCount -gt 0) {
    Write-Host "  🧹 清理标签: $cleanedCount 个现有SRT文件" -ForegroundColor Cyan
}

$finalSrtFiles = Get-ChildItem @srtScanParams
Write-Host ""
Write-Host "📝 当前目录SRT字幕文件总数: $($finalSrtFiles.Count) 个" -ForegroundColor White

if (-not $NoDelete -and $successCount -gt 0) {
    Write-Host ""
    Write-Host "🗑️  已删除 $successCount 个原始字幕文件" -ForegroundColor Gray
}

Write-Host ""
Write-Host "✨ 所有任务已完成！" -ForegroundColor Green

if (-not $NonInteractive) {
    Read-Host "按任意键退出"
}
