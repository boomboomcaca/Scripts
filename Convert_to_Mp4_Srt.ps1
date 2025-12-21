# GPU加速视频转换 + 字幕清理 + 编码分析工具

param(
    [string]$Path = ".",
    [switch]$NoDelete,
    [switch]$AnalyzeOnly,
    [switch]$Detailed,
    [switch]$ExportCsv,
    [switch]$Help,
    [switch]$NonInteractive
)

# 显示帮助信息
if ($Help) {
    Write-Host @"
GPU加速视频转换 + 字幕清理 + 编码分析工具

用法:
    .\Convert_to_Mp4_Srt.ps1 [-Path <目录路径>] [-NoDelete] [-AnalyzeOnly] [-Detailed] [-ExportCsv] [-NonInteractive] [-Help]

参数:
    -Path          指定要处理的目录路径 (默认: 当前目录)
    -NoDelete      保留原始文件，不删除
    -AnalyzeOnly   仅分析编码格式，不进行转换
    -Detailed      显示详细的编码信息
    -ExportCsv     导出详细的CSV分析报告
    -NonInteractive 非交互模式，不等待按键退出（用于自动化调用）
    -Help          显示此帮助信息

功能:
    1. 检查ffmpeg环境
    2. 智能检测视频编码格式（H.264、H.265、AV1、VP9等40+种编码）
    3. 彩色分类显示编码格式并提供统计报告
    4. GPU加速转换所有非MP4+H.264文件为MP4+H.264编码
    5. 转换VTT/ASS/SSA/SUB/SBV字幕为SRT格式并清理格式标签
    6. 生成详细的编码分析报告（可选CSV导出）

编码分类支持:
    ✅ 现代编码 (H.264) - 目标格式，最佳兼容性
    📱 高效编码 (H.265/HEVC, AV1, VP9) - 高效但兼容性有限
    🔄 传统编码 (MPEG-1/2/4, WMV系列, VC-1) - 需要转换
    🎬 专业编码 (Apple ProRes, DV Video) - 视情况转换
    🌐 Web编码 (VP8, Theora) - 建议转换
    ⚡ 专有格式 (RealVideo, Indeo, Cinepak) - 必须转换
    💎 无损编码 (MJPEG, HuffYUV, FFV1) - 特殊处理

支持的视频格式:
    TS、AVI、MKV、MOV、WMV、FLV、WEBM、M4V、3GP、MPG、MPEG、OGV、ASF、RM、RMVB

示例:
    .\Convert_to_Mp4_Srt.ps1                        # 分析并转换当前目录
    .\Convert_to_Mp4_Srt.ps1 -Path "D:\Videos"      # 处理指定目录
    .\Convert_to_Mp4_Srt.ps1 -AnalyzeOnly           # 仅分析编码格式
    .\Convert_to_Mp4_Srt.ps1 -AnalyzeOnly -Detailed # 详细分析编码格式
    .\Convert_to_Mp4_Srt.ps1 -ExportCsv             # 分析转换并导出CSV报告
    .\Convert_to_Mp4_Srt.ps1 -NoDelete              # 保留原始文件

注意事项:
    - 需要安装ffmpeg并添加到PATH环境变量
    - 建议使用NVIDIA显卡以获得GPU加速（可选）
    - 使用-NoDelete参数可保留原始文件
    - 脚本会自动检测并转换所有非MP4+H.264编码的视频文件
"@
    exit 0
}

# 设置控制台编码为UTF-8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$Host.UI.RawUI.WindowTitle = "GPU加速视频转换 + 字幕清理 + 编码分析工具"

# 增强的视频编码检测函数
function Get-VideoCodec {
    param([string]$FilePath)
    
    try {
        $ffprobeOutput = ffprobe -v quiet -print_format json -show_streams "$FilePath" 2>&1 | ConvertFrom-Json
        $videoStream = $ffprobeOutput.streams | Where-Object { $_.codec_type -eq "video" } | Select-Object -First 1
        
        if ($videoStream) {
            $codecName = $videoStream.codec_name
            
            # 定义编码分类 - 扩展支持40+种编码格式
            $modernCodecs = @("h264")  # 只有H.264被认为是目标格式
            $highEfficiencyCodecs = @("hevc", "av1", "vp9")  # 高效但可能兼容性不足
            $legacyCodecs = @("mpeg1video", "mpeg2video", "mpeg4", "wmv1", "wmv2", "wmv3", "vc1", "msmpeg4v1", "msmpeg4v2", "msmpeg4v3")
            $professionalCodecs = @("prores", "dvvideo", "dnxhd", "cineform")
            $webCodecs = @("vp8", "theora")
            $proprietaryCodecs = @("rv10", "rv20", "rv30", "rv40", "indeo2", "indeo3", "indeo4", "indeo5", "cinepak", "truemotion1", "truemotion2")
            $losslessCodecs = @("mjpeg", "huffyuv", "ffv1", "utvideo", "lagarith")
            $rawCodecs = @("rawvideo", "v210", "v308", "v408", "v410")
            
            # 确定编码类型和建议
            $codecCategory = "unknown"
            $shouldConvert = $true
            
            if ($codecName -in $modernCodecs) {
                $codecCategory = "modern"
                $shouldConvert = $false
            } elseif ($codecName -in $highEfficiencyCodecs) {
                $codecCategory = "high_efficiency"
            } elseif ($codecName -in $legacyCodecs) {
                $codecCategory = "legacy"
            } elseif ($codecName -in $professionalCodecs) {
                $codecCategory = "professional"
            } elseif ($codecName -in $webCodecs) {
                $codecCategory = "web"
            } elseif ($codecName -in $proprietaryCodecs) {
                $codecCategory = "proprietary"
            } elseif ($codecName -in $losslessCodecs) {
                $codecCategory = "lossless"
            } elseif ($codecName -in $rawCodecs) {
                $codecCategory = "raw"
            }
            
            # 获取详细视频信息
            $duration = if ($videoStream.duration) { [Math]::Round([double]$videoStream.duration, 2) } else { 0 }
            $bitrate = if ($videoStream.bit_rate) { [Math]::Round([double]$videoStream.bit_rate / 1000000, 2) } else { 0 }
            $resolution = if ($videoStream.width -and $videoStream.height) { "$($videoStream.width)x$($videoStream.height)" } else { "未知" }
            $fps = if ($videoStream.r_frame_rate) { 
                $frameParts = $videoStream.r_frame_rate -split "/"
                if ($frameParts.Length -eq 2 -and $frameParts[1] -ne "0") {
                    [Math]::Round([double]$frameParts[0] / [double]$frameParts[1], 2)
                } else { 0 }
            } else { 0 }
            
            return @{
                CodecName = $codecName
                CodecLongName = $videoStream.codec_long_name
                CodecCategory = $codecCategory
                ShouldConvert = $shouldConvert
                Duration = $duration
                Bitrate = $bitrate
                Resolution = $resolution
                FrameRate = $fps
                PixelFormat = $videoStream.pix_fmt
                IsH264 = $codecName -eq "h264"
                IsH265 = $codecName -eq "hevc"
                IsAV1 = $codecName -eq "av1"
                IsVP9 = $codecName -eq "vp9"
                IsLegacy = $codecName -in $legacyCodecs
                IsProprietary = $codecName -in $proprietaryCodecs
                IsHighEfficiency = $codecName -in $highEfficiencyCodecs
                IsWeb = $codecName -in $webCodecs
                IsProfessional = $codecName -in $professionalCodecs
                IsLossless = $codecName -in $losslessCodecs
                IsRaw = $codecName -in $rawCodecs
            }
        } else {
            return $null
        }
    } catch {
        Write-Host "警告: 无法检测视频编码格式: $FilePath" -ForegroundColor Yellow
        return $null
    }
}

# 显示编码统计函数
function Show-CodecStatistics {
    param([array]$VideoFiles)
    
    Write-Host ""
    Write-Host "🔍 编码格式统计:" -ForegroundColor Cyan
    $codecStats = @{}
    $codecCounts = @{}
    $totalFiles = $VideoFiles.Count
    
    foreach ($file in $VideoFiles) {
        $codecInfo = Get-VideoCodec -FilePath $file.FullName
        if ($codecInfo) {
            # 按类别统计
            $category = $codecInfo.CodecCategory
            if ($codecStats.ContainsKey($category)) {
                $codecStats[$category] += 1
            } else {
                $codecStats[$category] = 1
            }
            
            # 按具体编码统计
            $codec = $codecInfo.CodecName.ToUpper()
            if ($codecCounts.ContainsKey($codec)) {
                $codecCounts[$codec] += 1
            } else {
                $codecCounts[$codec] = 1
            }
        }
    }
    
    # 显示按类别统计
    foreach ($category in $codecStats.Keys | Sort-Object) {
        $count = $codecStats[$category]
        $percentage = [Math]::Round(($count / $totalFiles) * 100, 1)
        $emoji = switch ($category) {
            "modern" { "✅" }
            "high_efficiency" { "📱" }
            "legacy" { "🔄" }
            "professional" { "🎬" }
            "web" { "🌐" }
            "proprietary" { "⚡" }
            "lossless" { "💎" }
            "raw" { "🎞️" }
            "unknown" { "❓" }
        }
        $description = switch ($category) {
            "modern" { "现代编码 (H.264)" }
            "high_efficiency" { "高效编码 (H.265/AV1/VP9)" }
            "legacy" { "传统编码 (MPEG系列/WMV)" }
            "professional" { "专业编码 (ProRes/DV)" }
            "web" { "Web编码 (VP8/Theora)" }
            "proprietary" { "专有格式 (RealVideo/Indeo)" }
            "lossless" { "无损编码 (MJPEG/FFV1)" }
            "raw" { "原始格式 (RAW/YUV)" }
            "unknown" { "未知编码" }
        }
        Write-Host "  $emoji $description`: $count 个文件 ($percentage%)" -ForegroundColor White
    }
    
    # 显示具体编码格式分布
    if ($Detailed) {
        Write-Host ""
        Write-Host "🎯 具体编码格式分布:" -ForegroundColor Green
        foreach ($codec in $codecCounts.Keys | Sort-Object) {
            $count = $codecCounts[$codec]
            $percentage = [Math]::Round(($count / $totalFiles) * 100, 1)
            Write-Host "    • $codec`: $count 个文件 ($percentage%)" -ForegroundColor Cyan
        }
    }
    
    return @{
        CodecStats = $codecStats
        CodecCounts = $codecCounts
        TotalFiles = $totalFiles
    }
}

# 生成分析报告函数
function New-AnalysisReport {
    param([array]$VideoFiles)
    
    $analysisResults = @()
    $totalFiles = $VideoFiles.Count
    $currentFile = 0
    
    Write-Host "🔍 生成详细分析报告..." -ForegroundColor Green
    
    foreach ($file in $VideoFiles) {
        $currentFile++
        $progress = [Math]::Round(($currentFile / $totalFiles) * 100, 1)
        
        if ($Detailed) {
            Write-Host "[$progress%] 分析: $($file.Name)" -ForegroundColor Cyan
        }
        
        $codecInfo = Get-VideoCodec -FilePath $file.FullName
        
        if ($codecInfo) {
            $analysisResults += @{
                FileName = $file.Name
                FileSize = [Math]::Round($file.Length / 1MB, 2)
                Extension = $file.Extension.ToUpper()
                CodecName = $codecInfo.CodecName
                CodecLongName = $codecInfo.CodecLongName
                CodecCategory = $codecInfo.CodecCategory
                Duration = $codecInfo.Duration
                Bitrate = $codecInfo.Bitrate
                Resolution = $codecInfo.Resolution
                FrameRate = $codecInfo.FrameRate
                PixelFormat = $codecInfo.PixelFormat
                ShouldConvert = $codecInfo.ShouldConvert
            }
            
            if ($Detailed) {
                $emoji = switch ($codecInfo.CodecCategory) {
                    "modern" { "✅" }
                    "high_efficiency" { "📱" }
                    "legacy" { "🔄" }
                    "professional" { "🎬" }
                    "web" { "🌐" }
                    "proprietary" { "⚡" }
                    "lossless" { "💎" }
                    "raw" { "🎞️" }
                    "unknown" { "❓" }
                }
                
                Write-Host "  $emoji $($codecInfo.CodecName.ToUpper()) ($($codecInfo.CodecCategory)) - $($codecInfo.Resolution)" -ForegroundColor White
            }
        } else {
            $analysisResults += @{
                FileName = $file.Name
                FileSize = [Math]::Round($file.Length / 1MB, 2)
                Extension = $file.Extension.ToUpper()
                CodecName = "检测失败"
                CodecLongName = "无法检测"
                CodecCategory = "unknown"
                Duration = 0
                Bitrate = 0
                Resolution = "未知"
                FrameRate = 0
                PixelFormat = "未知"
                ShouldConvert = $true
            }
        }
    }
    
    return $analysisResults
}

# 显示转换建议函数
function Show-ConversionRecommendations {
    param([array]$AnalysisResults)
    
    Write-Host ""
    Write-Host "💡 转换建议:" -ForegroundColor Green
    $needConversion = $AnalysisResults | Where-Object { $_.ShouldConvert -eq $true }
    
    if ($needConversion.Count -gt 0) {
        Write-Host "  🔄 建议转换 $($needConversion.Count) 个文件为H.264编码以获得更好的兼容性" -ForegroundColor Yellow
        
        $highPriority = $needConversion | Where-Object { $_.CodecCategory -in @("proprietary", "legacy") }
        if ($highPriority.Count -gt 0) {
            Write-Host "  ⚡ 高优先级转换: $($highPriority.Count) 个专有/传统格式文件" -ForegroundColor Red
        }
        
        $modernHigh = $needConversion | Where-Object { $_.CodecCategory -eq "high_efficiency" }
        if ($modernHigh.Count -gt 0) {
            Write-Host "  📱 兼容性转换: $($modernHigh.Count) 个高效编码文件（可选）" -ForegroundColor Cyan
        }
    } else {
        Write-Host "  ✅ 所有文件都是H.264编码，无需转换" -ForegroundColor Green
    }
}

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host " GPU加速视频转换 + 字幕清理 + 编码分析工具" -ForegroundColor Cyan
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
Write-Host "[1/6] 检查ffmpeg环境..." -ForegroundColor Green
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

# 定义支持的视频文件格式
$VideoExtensions = @("*.mp4", "*.ts", "*.avi", "*.mkv", "*.mov", "*.wmv", "*.flv", "*.webm", "*.m4v", "*.3gp", "*.mpg", "*.mpeg", "*.ogv", "*.asf", "*.rm", "*.rmvb")

# 获取所有视频文件
Write-Host ""
Write-Host "[2/6] 扫描视频文件..." -ForegroundColor Green
$allVideoFiles = @()
foreach ($ext in $VideoExtensions) {
    $files = Get-ChildItem -Filter $ext -ErrorAction SilentlyContinue | Where-Object { 
        # 排除临时文件
        $_.Name -notmatch "\.temp\." -and $_.Name -notmatch "\.tmp\."
    }
    if ($files) {
        $allVideoFiles += $files
    }
}

# 统计字幕文件
$vttFiles = Get-ChildItem -Filter "*.vtt" -ErrorAction SilentlyContinue
$assFiles = Get-ChildItem -Filter "*.ass" -ErrorAction SilentlyContinue
$ssaFiles = Get-ChildItem -Filter "*.ssa" -ErrorAction SilentlyContinue
$subFiles = Get-ChildItem -Filter "*.sub" -ErrorAction SilentlyContinue
$sbvFiles = Get-ChildItem -Filter "*.sbv" -ErrorAction SilentlyContinue
$srtFiles = Get-ChildItem -Filter "*.srt" -ErrorAction SilentlyContinue
$totalSubFiles = $vttFiles.Count + $assFiles.Count + $ssaFiles.Count + $subFiles.Count + $sbvFiles.Count

if ($allVideoFiles.Count -eq 0) {
    Write-Host "⚠️  未找到任何视频文件" -ForegroundColor Yellow
    Write-Host "支持的格式: TS, AVI, MKV, MOV, WMV, FLV, WEBM, M4V, 3GP, MPG, MPEG, OGV, ASF, RM, RMVB" -ForegroundColor Gray
    if ($totalSubFiles -gt 0 -or $srtFiles.Count -gt 0) {
        Write-Host "但发现字幕文件，将处理字幕转换..." -ForegroundColor Yellow
    } else {
        exit 0
    }
}

Write-Host "📊 找到 $($allVideoFiles.Count) 个视频文件" -ForegroundColor White
Write-Host "📊 找到 $totalSubFiles 个需转换的字幕文件 (VTT: $($vttFiles.Count), ASS: $($assFiles.Count), SSA: $($ssaFiles.Count), SUB: $($subFiles.Count), SBV: $($sbvFiles.Count))" -ForegroundColor White
Write-Host "📊 找到 $($srtFiles.Count) 个SRT字幕文件" -ForegroundColor White

# 执行编码分析
if ($allVideoFiles.Count -gt 0) {
    Write-Host ""
    Write-Host "[3/6] 视频编码分析..." -ForegroundColor Green
    
    # 生成详细分析报告
    $analysisResults = New-AnalysisReport -VideoFiles $allVideoFiles
    
    # 显示编码统计
    Show-CodecStatistics -VideoFiles $allVideoFiles | Out-Null
    
    # 显示转换建议
    Show-ConversionRecommendations -AnalysisResults $analysisResults
    
    # 导出CSV报告
    if ($ExportCsv) {
        $csvPath = "Video_Codec_Analysis_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        $analysisResults | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
        Write-Host ""
        Write-Host "📋 详细分析报告已导出到: $csvPath" -ForegroundColor Green
    }
    
    # 如果只是分析模式，则退出
    if ($AnalyzeOnly) {
        Write-Host ""
        Write-Host "✨ 编码分析完成!" -ForegroundColor Green
        if (-not $NonInteractive) {
            Read-Host "按任意键退出"
        }
        exit 0
    }
    
    # 分类文件
    Write-Host ""
    Write-Host "[4/6] 分类视频文件..." -ForegroundColor Green
    
    $mp4H264Files = @()
    $nonMp4H264Files = @()
    
    foreach ($file in $allVideoFiles) {
        try {
            $codecInfo = Get-VideoCodec -FilePath $file.FullName
            
            if ($codecInfo -and $file.Extension -eq '.mp4' -and $codecInfo.IsH264) {
                $mp4H264Files += $file
                Write-Host "  ✅ $($file.Name): MP4 + H.264" -ForegroundColor Green
            } else {
                $nonMp4H264Files += $file
                if ($codecInfo) {
                    $codecDisplay = $codecInfo.CodecName.ToUpper()
                    $statusIcon = "⚠"
                    $displayColor = "Yellow"
                    
                    # 根据编码类别显示不同的信息和颜色
                    switch ($codecInfo.CodecCategory) {
                        "high_efficiency" {
                            $statusIcon = "📱"
                            $displayColor = "Cyan"
                            if ($codecInfo.IsAV1) { $codecDisplay += " (AV1-高效编码)" }
                            elseif ($codecInfo.IsH265) { $codecDisplay += " (H.265/HEVC)" }
                            elseif ($codecInfo.IsVP9) { $codecDisplay += " (VP9-Web)" }
                        }
                        "legacy" {
                            $statusIcon = "🔄"
                            $displayColor = "Magenta"
                            $codecDisplay += " (传统编码)"
                        }
                        "professional" {
                            $statusIcon = "🎬"
                            $displayColor = "Blue"
                            $codecDisplay += " (专业编码)"
                        }
                        "web" {
                            $statusIcon = "🌐"
                            $displayColor = "DarkCyan"
                            $codecDisplay += " (Web编码)"
                        }
                        "proprietary" {
                            $statusIcon = "⚡"
                            $displayColor = "Red"
                            $codecDisplay += " (专有格式)"
                        }
                        "lossless" {
                            $statusIcon = "💎"
                            $displayColor = "DarkYellow"
                            $codecDisplay += " (无损编码)"
                        }
                        "raw" {
                            $statusIcon = "🎞️"
                            $displayColor = "DarkGray"
                            $codecDisplay += " (原始格式)"
                        }
                        "unknown" {
                            $statusIcon = "❓"
                            $displayColor = "Gray"
                            $codecDisplay += " (未知编码)"
                        }
                    }
                    
                    Write-Host "  $statusIcon $($file.Name): $($file.Extension.ToUpper()) + $codecDisplay" -ForegroundColor $displayColor
                } else {
                    Write-Host "  ❓ $($file.Name): 无法检测编码" -ForegroundColor Gray
                }
            }
        } catch {
            $nonMp4H264Files += $file
            Write-Host "  ❌ $($file.Name): 检测失败" -ForegroundColor Red
        }
    }
    
    Write-Host ""
    Write-Host "📈 文件分类结果:" -ForegroundColor White
    Write-Host "  - MP4 + H.264编码文件: $($mp4H264Files.Count) 个" -ForegroundColor Green
    Write-Host "  - 需要转换的文件: $($nonMp4H264Files.Count) 个" -ForegroundColor Yellow
} else {
    $nonMp4H264Files = @()
}

# 转换成MP4 + H.264编码文件
Write-Host ""
Write-Host "[5/6] GPU加速转换为MP4 + H.264编码..." -ForegroundColor Green

if ($nonMp4H264Files.Count -gt 0) {
    Write-Host "🚀 正在使用NVIDIA GPU加速转换..." -ForegroundColor Yellow
    Write-Host "📊 发现 $($nonMp4H264Files.Count) 个需要转换的视频文件" -ForegroundColor White
    
    $successCount = 0
    $failureCount = 0
    
    foreach ($file in $nonMp4H264Files) {
        $outputFile = [System.IO.Path]::ChangeExtension($file.FullName, "mp4")
        
        # 如果是同名MP4文件，使用临时文件名
        if ($file.Extension -eq '.mp4') {
            $outputFile = [System.IO.Path]::ChangeExtension($file.FullName, "temp.mp4")
        }
        
        Write-Host "🔄 转换中: $($file.Name) -> $([System.IO.Path]::GetFileName($outputFile))" -ForegroundColor White
        
        # 检查输出文件是否已存在
        if ((Test-Path $outputFile) -and ($file.Extension -ne '.mp4')) {
            Write-Host "⚠️  目标文件已存在，跳过: $([System.IO.Path]::GetFileName($outputFile))" -ForegroundColor Yellow
            continue
        }
        
        try {
            # 根据文件格式选择不同的转换参数
            $ffmpegArgs = @(
                "-hwaccel", "cuda",
                "-i", "`"$($file.FullName)`"",
                "-c:v", "h264_nvenc",
                "-preset", "fast",
                "-crf", "23",
                "-c:a", "aac",
                "-ar", "48000",
                "-af", "loudnorm=I=-16:TP=-1.5:LRA=11",
                "-map_metadata", "0",
                "-y",
                "`"$outputFile`""
            )
            
            # 对于某些格式，可能需要使用软件解码
            if ($file.Extension -in @('.rm', '.rmvb', '.asf')) {
                $ffmpegArgs[1] = "auto"  # 不强制使用CUDA硬件加速解码
            }
            
            $process = Start-Process -FilePath "ffmpeg" -ArgumentList $ffmpegArgs -Wait -PassThru -NoNewWindow
            
            if ($process.ExitCode -eq 0) {
                Write-Host "✅ 成功转换: $($file.Name)" -ForegroundColor Green
                $successCount++
                
                # 如果是MP4文件重新编码，替换原文件
                if ($file.Extension -eq '.mp4') {
                    Remove-Item $file.FullName -Force
                    Move-Item $outputFile $file.FullName
                    Write-Host "✅ 已更新编码格式: $($file.Name)" -ForegroundColor Green
                } else {
                    # 删除原始文件（如果不是NoDelete模式）
                    if (-not $NoDelete) {
                        Remove-Item $file.FullName -Force
                        Write-Host "✅ 已删除源文件: $($file.Name)" -ForegroundColor Green
                    }
                }
            } else {
                Write-Host "❌ 转换失败: $($file.Name)" -ForegroundColor Red
                $failureCount++
                # 清理临时文件
                if (Test-Path $outputFile) {
                    Remove-Item $outputFile -Force
                }
            }
        } catch {
            Write-Host "❌ 转换出错: $($file.Name) - $($_.Exception.Message)" -ForegroundColor Red
            $failureCount++
            # 清理临时文件
            if (Test-Path $outputFile) {
                Remove-Item $outputFile -Force
            }
        }
        Write-Host ""
    }
    
    Write-Host "📊 视频转换统计:" -ForegroundColor Green
    Write-Host "  ✅ 成功转换: $successCount 个文件" -ForegroundColor Green
    if ($failureCount -gt 0) {
        Write-Host "  ❌ 转换失败: $failureCount 个文件" -ForegroundColor Red
    }
} else {
    Write-Host "✅ 所有视频文件已经是MP4 + H.264格式，跳过转换" -ForegroundColor Yellow
}

# 清理字幕文本的函数
function Clear-SubtitleText {
    param([string]$Text)
    
    # 移除HTML标签
    $cleaned = $Text -replace '<[^>]*>', ''
    
    # 移除ASS/SSA格式标签 {\...}
    $cleaned = $cleaned -replace '\{[^}]*\}', ''
    
    # 替换HTML实体
    $cleaned = $cleaned -replace '&amp;', '&'
    $cleaned = $cleaned -replace '&lt;', '<'
    $cleaned = $cleaned -replace '&gt;', '>'
    $cleaned = $cleaned -replace '&quot;', '"'
    $cleaned = $cleaned -replace '&#39;', "'"
    $cleaned = $cleaned -replace '&nbsp;', ' '
    
    return $cleaned
}

# 转换各种字幕格式并清理HTML标签
Write-Host ""
Write-Host "[6/7] 处理字幕文件..." -ForegroundColor Green

# 合并所有需要转换的字幕文件
$allSubtitleFiles = @()
$allSubtitleFiles += $vttFiles | ForEach-Object { @{ File = $_; Format = "VTT" } }
$allSubtitleFiles += $assFiles | ForEach-Object { @{ File = $_; Format = "ASS" } }
$allSubtitleFiles += $ssaFiles | ForEach-Object { @{ File = $_; Format = "SSA" } }
$allSubtitleFiles += $subFiles | ForEach-Object { @{ File = $_; Format = "SUB" } }
$allSubtitleFiles += $sbvFiles | ForEach-Object { @{ File = $_; Format = "SBV" } }

if ($allSubtitleFiles.Count -gt 0) {
    Write-Host "🔄 转换字幕文件为SRT格式..." -ForegroundColor Yellow
    $subSuccessCount = 0
    
    foreach ($item in $allSubtitleFiles) {
        $file = $item.File
        $format = $item.Format
        $outputFile = [System.IO.Path]::ChangeExtension($file.FullName, "srt")
        
        # 检查是否已存在SRT文件
        if (Test-Path $outputFile) {
            Write-Host "⏭️  跳过 (SRT已存在): $($file.Name)" -ForegroundColor Gray
            continue
        }
        
        Write-Host "📝 转换字幕 [$format]: $($file.Name) -> $([System.IO.Path]::GetFileName($outputFile))" -ForegroundColor White
        
        try {
            $process = Start-Process -FilePath "ffmpeg" -ArgumentList @(
                "-i", "`"$($file.FullName)`"",
                "-y",
                "`"$outputFile`""
            ) -Wait -PassThru -NoNewWindow -RedirectStandardError "$env:TEMP\ffmpeg_error.txt"
            
            if ($process.ExitCode -eq 0 -and (Test-Path $outputFile)) {
                Write-Host "✅ 转换完成，正在清理格式标签..." -ForegroundColor Green
                
                # 清理格式标签
                try {
                    $content = Get-Content $outputFile -Raw -Encoding UTF8
                    $content = Clear-SubtitleText -Text $content
                    [System.IO.File]::WriteAllText($outputFile, $content, [System.Text.Encoding]::UTF8)
                    
                    Write-Host "✅ 成功转换并清理: $($file.Name)" -ForegroundColor Green
                    $subSuccessCount++
                    
                    if (-not $NoDelete) {
                        Remove-Item $file.FullName -Force
                        Write-Host "✅ 已删除源文件: $($file.Name)" -ForegroundColor Gray
                    }
                } catch {
                    Write-Host "⚠️  格式标签清理失败，但文件已转换: $($file.Name)" -ForegroundColor Yellow
                    $subSuccessCount++
                }
            } else {
                Write-Host "❌ 字幕转换失败: $($file.Name)" -ForegroundColor Red
            }
        } catch {
            Write-Host "❌ 转换出错: $($file.Name) - $($_.Exception.Message)" -ForegroundColor Red
        }
        Write-Host ""
    }
    
    Write-Host "📊 字幕转换统计: 成功 $subSuccessCount/$($allSubtitleFiles.Count) 个文件" -ForegroundColor Green
} else {
    Write-Host "📝 未发现需要转换的字幕文件，跳过转换" -ForegroundColor Gray
}

# 清理现有SRT文件的格式标签
Write-Host ""
Write-Host "🧹 清理现有SRT文件中的格式标签..." -ForegroundColor Yellow
$cleanedCount = 0
$currentSrtFiles = Get-ChildItem -Filter "*.srt" -ErrorAction SilentlyContinue

foreach ($file in $currentSrtFiles) {
    Write-Host "🔍 检查: $($file.Name)" -ForegroundColor White
    
    try {
        $content = Get-Content $file.FullName -Raw -Encoding UTF8
        $originalContent = $content
        $content = Clear-SubtitleText -Text $content
        
        if ($content -ne $originalContent) {
            [System.IO.File]::WriteAllText($file.FullName, $content, [System.Text.Encoding]::UTF8)
            Write-Host "✅ 已清理: $($file.Name)" -ForegroundColor Green
            $cleanedCount++
        } else {
            Write-Host "✨ 无需清理: $($file.Name)" -ForegroundColor Gray
        }
    } catch {
        Write-Host "❌ 清理失败: $($file.Name)" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "📊 SRT清理统计: 清理了 $cleanedCount 个文件" -ForegroundColor Green

# 显示最终结果
Write-Host ""
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "              处理完成！" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "🎉 最终结果统计:" -ForegroundColor Green
$finalMp4Files = Get-ChildItem -Filter "*.mp4" -ErrorAction SilentlyContinue
$finalSrtFiles = Get-ChildItem -Filter "*.srt" -ErrorAction SilentlyContinue

Write-Host "  📹 MP4+H.264编码文件: $($finalMp4Files.Count) 个" -ForegroundColor White
Write-Host "  📝 SRT字幕文件: $($finalSrtFiles.Count) 个" -ForegroundColor White

if ($ExportCsv -and (Test-Path "Video_Codec_Analysis_*.csv")) {
    $csvFile = Get-ChildItem -Filter "Video_Codec_Analysis_*.csv" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    Write-Host "  📋 分析报告: $($csvFile.Name)" -ForegroundColor White
}

Write-Host ""
Write-Host "📋 已处理的项目:" -ForegroundColor Green
Write-Host "  🎬 视频编码分析和分类显示" -ForegroundColor White
Write-Host "  🔄 视频格式转换为MP4+H.264" -ForegroundColor White
Write-Host "  📝 字幕格式转换为SRT (支持VTT/ASS/SSA/SUB/SBV)" -ForegroundColor White
Write-Host "  🧹 格式标签清理 (HTML/ASS/SSA)" -ForegroundColor White


# 移动MP4和SRT文件到网络文件夹（带 NSFW 检测）
Write-Host ""
Write-Host "[7/7] NSFW 检测并移动文件到网络文件夹..." -ForegroundColor Green

$networkPathNSFW = "\\192.168.1.111\data\Scenes"    # NSFW 内容
$networkPathSafe = "\\192.168.1.111\data\Movies"    # 普通内容
$nsfwDetectScript = "D:\Soft\Scripts\nsfw_detect.py"

# NSFW 检测函数
function Test-NSFWContentLocal {
    param([string]$VideoPath)
    
    Write-Host "  🔍 正在进行 NSFW 检测..." -ForegroundColor Yellow
    
    try {
        $result = python $nsfwDetectScript $VideoPath 2>&1
        $exitCode = $LASTEXITCODE
        
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
            if ($exitCode -eq 1) {
                Write-Host "  🔞 检测结果: NSFW" -ForegroundColor Magenta
                return $true
            } else {
                Write-Host "  ✅ 检测结果: 普通内容" -ForegroundColor Green
                return $false
            }
        }
    } catch {
        Write-Host "  ⚠️ NSFW 检测失败，默认归类为普通内容" -ForegroundColor Yellow
        return $false
    }
}

# 检查网络路径是否可访问
$nsfwPathOK = Test-Path $networkPathNSFW
$safePathOK = Test-Path $networkPathSafe

if ($nsfwPathOK -or $safePathOK) {
    Write-Host "  📁 NSFW 目录: $networkPathNSFW $(if($nsfwPathOK){'✅'}else{'❌'})" -ForegroundColor $(if($nsfwPathOK){'Green'}else{'Red'})
    Write-Host "  📁 普通目录: $networkPathSafe $(if($safePathOK){'✅'}else{'❌'})" -ForegroundColor $(if($safePathOK){'Green'}else{'Red'})
    
    # 获取所有MP4和SRT文件
    $mp4FilesToMove = Get-ChildItem -Filter "*.mp4" -ErrorAction SilentlyContinue
    $srtFilesToMove = Get-ChildItem -Filter "*.srt" -ErrorAction SilentlyContinue
    
    if ($mp4FilesToMove.Count -gt 0 -or $srtFilesToMove.Count -gt 0) {
        Write-Host "  📊 找到文件: MP4: $($mp4FilesToMove.Count), SRT: $($srtFilesToMove.Count)" -ForegroundColor White
        
        $movedNSFW = 0
        $movedSafe = 0
        $skipCount = 0
        $errorCount = 0
        
        # 先处理 MP4 文件（进行 NSFW 检测）
        foreach ($file in $mp4FilesToMove) {
            Write-Host ""
            Write-Host "  📹 处理: $($file.Name)" -ForegroundColor Cyan
            
            $isNSFW = Test-NSFWContentLocal -VideoPath $file.FullName
            $destPath = if ($isNSFW) { $networkPathNSFW } else { $networkPathSafe }
            $categoryLabel = if ($isNSFW) { "Scenes (NSFW)" } else { "Movies (普通)" }
            
            Write-Host "  📁 目标: $categoryLabel" -ForegroundColor Cyan
            
            if (-not (Test-Path $destPath)) {
                Write-Host "  ❌ 无法访问目标路径: $destPath" -ForegroundColor Red
                $errorCount++
                continue
            }
            
            $destinationPath = Join-Path $destPath $file.Name
            
            try {
                if (Test-Path $destinationPath) {
                    $sourceSize = $file.Length
                    $destSize = (Get-Item $destinationPath).Length
                    if ($sourceSize -gt $destSize) {
                        Move-Item -Path $file.FullName -Destination $destinationPath -Force
                        Write-Host "  ✅ 已覆盖 (源更大)" -ForegroundColor Green
                        if ($isNSFW) { $movedNSFW++ } else { $movedSafe++ }
                    } else {
                        Remove-Item $file.FullName -Force
                        Write-Host "  ✅ 已删除源文件 (目标更大)" -ForegroundColor Green
                        $skipCount++
                    }
                } else {
                    Move-Item -Path $file.FullName -Destination $destinationPath -Force
                    Write-Host "  ✅ 已移动" -ForegroundColor Green
                    if ($isNSFW) { $movedNSFW++ } else { $movedSafe++ }
                }
            } catch {
                Write-Host "  ❌ 移动失败: $($_.Exception.Message)" -ForegroundColor Red
                $errorCount++
            }
        }
        
        # 处理 SRT 文件（跟随对应的 MP4 文件）
        foreach ($file in $srtFilesToMove) {
            Write-Host ""
            Write-Host "  📝 处理: $($file.Name)" -ForegroundColor Cyan
            
            $mp4Name = [System.IO.Path]::ChangeExtension($file.Name, ".mp4")
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
            
            Write-Host "  📁 目标: $categoryLabel" -ForegroundColor Cyan
            
            if (-not (Test-Path $destPath)) {
                Write-Host "  ❌ 无法访问目标路径: $destPath" -ForegroundColor Red
                $errorCount++
                continue
            }
            
            $destinationPath = Join-Path $destPath $file.Name
            
            try {
                if (Test-Path $destinationPath) {
                    Write-Host "  ⏭️ 跳过 (目标已存在)" -ForegroundColor Yellow
                    $skipCount++
                } else {
                    Move-Item -Path $file.FullName -Destination $destinationPath -Force
                    Write-Host "  ✅ 已移动" -ForegroundColor Green
                    if ($destPath -eq $networkPathNSFW) { $movedNSFW++ } else { $movedSafe++ }
                }
            } catch {
                Write-Host "  ❌ 移动失败: $($_.Exception.Message)" -ForegroundColor Red
                $errorCount++
            }
        }
        
        Write-Host ""
        Write-Host "  📊 文件移动统计:" -ForegroundColor Green
        Write-Host "    🔞 移动到 Scenes (NSFW): $movedNSFW 个文件" -ForegroundColor Magenta
        Write-Host "    📁 移动到 Movies (普通): $movedSafe 个文件" -ForegroundColor Green
        if ($skipCount -gt 0) {
            Write-Host "    ⏭️ 跳过: $skipCount 个文件" -ForegroundColor Yellow
        }
        if ($errorCount -gt 0) {
            Write-Host "    ❌ 失败: $errorCount 个文件" -ForegroundColor Red
        }
    } else {
        Write-Host "  未找到需要移动的MP4或SRT文件" -ForegroundColor Yellow
    }
} else {
    Write-Host "  ❌ 无法访问任何网络路径" -ForegroundColor Red
    Write-Host "    请确认网络连接和路径权限" -ForegroundColor Yellow
}
Write-Host ""
Write-Host "✨ 所有任务已完成！" -ForegroundColor Green

if (-not $AnalyzeOnly -and -not $NonInteractive) {
    Read-Host "按任意键退出"
}
