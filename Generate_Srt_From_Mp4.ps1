# MP4语音识别转SRT字幕工具
# 作者: Claude
# 功能: 使用faster-whisper对MP4视频进行语音识别并生成SRT字幕

param(
    [string]$Path = ".",
    [switch]$NonInteractive,
    [switch]$Help,
    [string]$Model = "large-v3",
    [string]$Language = "auto"
)

# 显示帮助信息
if ($Help) {
    Write-Host @"
MP4语音识别转SRT字幕工具

用法:
    .\Generate_Srt_From_Mp4.ps1 [-Path <目录路径>] [-Model <模型>] [-Language <语言>] [-NonInteractive] [-Help]

参数:
    -Path          指定要处理的目录路径 (默认: 当前目录)
    -Model         Whisper模型大小 (tiny/base/small/medium/large-v3/turbo, 默认: large-v3)
    -Language      语言代码 (auto=自动检测, zh=中文, en=英文, ja=日语等, 默认: auto)
    -NonInteractive 非交互模式，不等待按键退出（用于自动化调用）
    -Help          显示此帮助信息

功能:
    1. 检查Python和faster-whisper环境
    2. 扫描指定目录下的MP4文件
    3. 跳过已有同名SRT字幕的MP4文件
    4. 智能语言识别和翻译:
       • 中文音频 → 中文字幕
       • 英文音频 → 英文字幕
       • 其他语言 → 英文字幕（自动翻译）
    5. 生成SRT格式字幕文件

模型说明:
    - tiny:     最快，准确度较低，约39M
    - base:     快速，准确度一般，约74M
    - small:    较快，准确度中等，约244M
    - medium:   平衡，准确度高，约769M
    - turbo:    快速推荐，准确度接近large，速度快8倍，约809M
    - large-v3: 最高准确度⭐，处理较慢，约1550M (默认)

示例:
    .\Generate_Srt_From_Mp4.ps1                           # 自动检测语言，智能生成字幕
    .\Generate_Srt_From_Mp4.ps1 -Path "D:\Videos"         # 处理指定目录
    .\Generate_Srt_From_Mp4.ps1 -Model turbo              # 使用turbo模型（更快）
    .\Generate_Srt_From_Mp4.ps1 -Language zh              # 强制识别为中文
    .\Generate_Srt_From_Mp4.ps1 -Language en              # 强制识别为英文
    .\Generate_Srt_From_Mp4.ps1 -Language ja              # 日语音频翻译为英文字幕
    .\Generate_Srt_From_Mp4.ps1 -NonInteractive           # 自动化模式

智能翻译说明:
    • 中文视频 → 自动生成中文字幕
    • 英文视频 → 自动生成英文字幕
    • 日语视频 → 自动翻译为英文字幕
    • 韩语视频 → 自动翻译为英文字幕
    • 法语视频 → 自动翻译为英文字幕
    （Whisper只支持翻译成英文，不支持翻译成中文）

注意事项:
    - 需要安装Python 3.8+
    - 需要安装faster-whisper库: pip install faster-whisper
    - 首次使用时会自动下载模型
    - 语音识别需要较长时间，请耐心等待
"@
    exit 0
}

# 设置控制台编码为UTF-8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$Host.UI.RawUI.WindowTitle = "MP4语音识别转SRT字幕工具"

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

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "   MP4语音识别转SRT字幕工具" -ForegroundColor Cyan
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

# 检查Python环境
Write-Host ""
Write-Host "[1/4] 检查Python环境..." -ForegroundColor Green

# 优先使用 Python 3.11（有 CUDA 支持）
$pythonCmd = "python"
$usePy311 = $false

try {
    # 尝试 py -3.11
    $py311Version = py -3.11 --version 2>&1
    if ($py311Version -match "Python 3\.11") {
        $pythonCmd = "py -3.11"
        $usePy311 = $true
        Write-Host "✅ 检测到Python 3.11: $py311Version" -ForegroundColor Green
        Write-Host "   (使用 Python 3.11 以支持 CUDA GPU 加速)" -ForegroundColor Gray
    }
} catch {}

if (-not $usePy311) {
    # 使用默认 python
    try {
        $pythonVersion = python --version 2>&1
        if ($pythonVersion -match "Python (\d+)\.(\d+)") {
            $majorVersion = [int]$matches[1]
            $minorVersion = [int]$matches[2]
            if ($majorVersion -ge 3 -and $minorVersion -ge 8) {
                Write-Host "✅ 检测到Python: $pythonVersion" -ForegroundColor Green
                if ($majorVersion -eq 3 -and $minorVersion -ge 13) {
                    Write-Host "   ⚠️  Python 3.13 可能不支持 CUDA，建议降级到 Python 3.11" -ForegroundColor Yellow
                }
            } else {
                throw "Python版本过低，需要3.8或更高版本"
            }
        } else {
            throw "无法检测Python版本"
        }
    } catch {
        Write-Host "❌ 未找到Python或版本不符合要求" -ForegroundColor Red
        Write-Host "请安装Python 3.8或更高版本" -ForegroundColor Yellow
        Write-Host "下载地址: https://www.python.org/downloads/" -ForegroundColor Yellow
        if (-not $NonInteractive) {
            Read-Host "按任意键退出"
        }
        exit 1
    }
}

# 检查faster-whisper库
Write-Host ""
Write-Host "[2/4] 检查faster-whisper库..." -ForegroundColor Green
$checkScript = @"
try:
    import faster_whisper
    print('installed')
except ImportError:
    print('not_installed')
"@

if ($pythonCmd -eq "py -3.11") {
    $checkResult = py -3.11 -c $checkScript 2>&1
} else {
    $checkResult = python -c $checkScript 2>&1
}
if ($checkResult -match "installed") {
    Write-Host "✅ faster-whisper已安装" -ForegroundColor Green
} else {
    Write-Host "❌ 未找到faster-whisper库" -ForegroundColor Red
    Write-Host ""
    Write-Host "请运行以下命令安装:" -ForegroundColor Yellow
    Write-Host "  pip install faster-whisper" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "或使用清华镜像加速:" -ForegroundColor Yellow
    Write-Host "  pip install -i https://pypi.tuna.tsinghua.edu.cn/simple faster-whisper" -ForegroundColor Cyan
    if (-not $NonInteractive) {
        Read-Host "按任意键退出"
    }
    exit 1
}

# 扫描MP4文件
Write-Host ""
Write-Host "[3/4] 扫描MP4文件..." -ForegroundColor Green
$mp4Files = Get-ChildItem -Filter "*.mp4" -ErrorAction SilentlyContinue | Where-Object { 
    # 排除临时文件
    $_.Name -notmatch "\.temp\." -and $_.Name -notmatch "\.tmp\."
}

if ($mp4Files.Count -eq 0) {
    Write-Host "⚠️  未找到任何MP4文件" -ForegroundColor Yellow
    if (-not $NonInteractive) {
        Read-Host "按任意键退出"
    }
    exit 0
}

Write-Host "📊 找到 $($mp4Files.Count) 个MP4文件" -ForegroundColor White

# 过滤出需要处理的文件（没有对应SRT字幕的）
$filesToProcess = @()
foreach ($file in $mp4Files) {
    $srtFile = [System.IO.Path]::ChangeExtension($file.FullName, "srt")
    if (Test-Path $srtFile) {
        Write-Host "  ⏭️  跳过 (已有字幕): $($file.Name)" -ForegroundColor Gray
    } else {
        $filesToProcess += $file
        Write-Host "  ✅ 待处理: $($file.Name)" -ForegroundColor Green
    }
}

if ($filesToProcess.Count -eq 0) {
    Write-Host ""
    Write-Host "✨ 所有MP4文件都已有SRT字幕，无需处理" -ForegroundColor Green
    if (-not $NonInteractive) {
        Read-Host "按任意键退出"
    }
    exit 0
}

Write-Host ""
Write-Host "📋 需要处理 $($filesToProcess.Count) 个MP4文件" -ForegroundColor Cyan

# 生成Python脚本用于语音识别
Write-Host ""
Write-Host "[4/4] 开始语音识别..." -ForegroundColor Green
Write-Host "🎯 使用模型: $Model" -ForegroundColor Cyan
Write-Host "🌐 识别语言: $Language" -ForegroundColor Cyan
Write-Host ""

$pythonScript = @"
import sys
import os
from faster_whisper import WhisperModel
import datetime

def format_time(seconds):
    """将秒数转换为SRT时间格式 (HH:MM:SS,mmm)"""
    td = datetime.timedelta(seconds=seconds)
    hours = int(td.total_seconds() // 3600)
    minutes = int((td.total_seconds() % 3600) // 60)
    seconds = td.total_seconds() % 60
    milliseconds = int((seconds % 1) * 1000)
    seconds = int(seconds)
    return f"{hours:02d}:{minutes:02d}:{seconds:02d},{milliseconds:03d}"

def transcribe_video(video_path, output_path, model_name, language):
    """使用faster-whisper转录视频并生成SRT字幕"""
    try:
        print(f"正在加载模型: {model_name}...")
        
        # 初始化模型（优先使用CUDA GPU加速）
        device = "cpu"
        compute_type = "int8"
        
        try:
            # 尝试使用CUDA GPU
            model = WhisperModel(model_name, device="cuda", compute_type="float16")
            device = "cuda"
            compute_type = "float16"
            print("✅ 使用CUDA GPU加速 (float16)")
        except Exception as e:
            error_msg = str(e)
            if "cublas" in error_msg.lower() or "cudnn" in error_msg.lower():
                print(f"❌ CUDA库缺失: {error_msg}")
                print("")
                print("⚠️  需要安装CUDA Toolkit和cuDNN:")
                print("   1. CUDA Toolkit 12.x: https://developer.nvidia.com/cuda-downloads")
                print("   2. cuDNN: https://developer.nvidia.com/cudnn")
                print("   3. 或安装PyTorch (包含CUDA): pip install torch --index-url https://download.pytorch.org/whl/cu121")
                print("")
                print("⏳ 正在降级到CPU模式...")
            else:
                print(f"⚠️  CUDA初始化失败: {error_msg}")
                print("⏳ 降级到CPU模式...")
            
            # 降级到CPU模式
            try:
                model = WhisperModel(model_name, device="cpu", compute_type="int8")
                device = "cpu"
                print("✅ 使用CPU模式（处理速度较慢）")
            except Exception as cpu_error:
                print(f"❌ CPU模式也失败: {str(cpu_error)}")
                raise
        
        print(f"正在转录: {os.path.basename(video_path)}")
        print("⏳ 这可能需要几分钟时间，请耐心等待...")
        print("")
        
        # 智能识别和翻译逻辑
        if language == "auto":
            # 先检测语言
            segments, info = model.transcribe(video_path, beam_size=5)
            detected_lang = info.language
            lang_prob = info.language_probability
            
            print(f"🌐 检测到语言: {detected_lang} (概率: {lang_prob:.2%})")
            
            # 判断是否需要翻译
            if detected_lang in ['zh', 'en']:
                # 中文或英文，直接转录
                print(f"✅ 使用转录模式 - 生成{('中文' if detected_lang == 'zh' else '英文')}字幕")
            else:
                # 其他语言，翻译成英文
                print(f"🔄 使用翻译模式 - 将 {detected_lang} 翻译为英文字幕")
                print("⏳ 正在重新处理...")
                segments, info = model.transcribe(video_path, task="translate", beam_size=5)
                print(f"✅ 翻译完成 - 已生成英文字幕")
        else:
            # 手动指定语言
            if language in ['zh', 'en']:
                # 中文或英文，直接转录
                segments, info = model.transcribe(video_path, language=language, beam_size=5)
                print(f"✅ 使用转录模式 - 生成{('中文' if language == 'zh' else '英文')}字幕")
            else:
                # 其他语言，翻译成英文
                print(f"🔄 使用翻译模式 - 将 {language} 翻译为英文字幕")
                segments, info = model.transcribe(video_path, language=language, task="translate", beam_size=5)
                print(f"✅ 翻译完成 - 已生成英文字幕")
        
        print("")
        
        # 生成SRT字幕
        with open(output_path, 'w', encoding='utf-8') as f:
            for i, segment in enumerate(segments, start=1):
                # SRT格式：序号、时间轴、字幕文本、空行
                f.write(f"{i}\n")
                f.write(f"{format_time(segment.start)} --> {format_time(segment.end)}\n")
                f.write(f"{segment.text.strip()}\n")
                f.write("\n")
                
                # 显示进度
                if i % 10 == 0:
                    print(f"  处理进度: {i} 条字幕")
        
        print(f"✅ 成功生成字幕: {os.path.basename(output_path)}")
        return True
        
    except Exception as e:
        print(f"❌ 转录失败: {str(e)}")
        return False

if __name__ == "__main__":
    if len(sys.argv) < 5:
        print("用法: python script.py <video_path> <output_path> <model_name> <language>")
        sys.exit(1)
    
    video_path = sys.argv[1]
    output_path = sys.argv[2]
    model_name = sys.argv[3]
    language = sys.argv[4]
    
    success = transcribe_video(video_path, output_path, model_name, language)
    sys.exit(0 if success else 1)
"@

# 保存Python脚本到临时文件
$tempPythonScript = Join-Path $env:TEMP "whisper_transcribe_temp.py"
$pythonScript | Out-File -FilePath $tempPythonScript -Encoding UTF8

$successCount = 0
$failureCount = 0

foreach ($file in $filesToProcess) {
    $outputFile = [System.IO.Path]::ChangeExtension($file.FullName, "srt")
    
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host "🎬 处理文件: $($file.Name)" -ForegroundColor Cyan
    Write-Host "📊 文件大小: $([Math]::Round($file.Length / 1MB, 2)) MB" -ForegroundColor White
    Write-Host ""
    
    try {
        # 调用Python脚本进行转录
        $pythonArgs = @(
            "`"$tempPythonScript`"",
            "`"$($file.FullName)`"",
            "`"$outputFile`"",
            $Model,
            $Language
        )
        
        if ($pythonCmd -eq "py -3.11") {
            $process = Start-Process -FilePath "py" -ArgumentList (@("-3.11") + $pythonArgs) -Wait -PassThru -NoNewWindow
        } else {
            $process = Start-Process -FilePath "python" -ArgumentList $pythonArgs -Wait -PassThru -NoNewWindow
        }
        
        if ($process.ExitCode -eq 0) {
            Write-Host "✅ 成功生成字幕: $($file.Name)" -ForegroundColor Green
            $successCount++
        } else {
            Write-Host "❌ 语音识别失败: $($file.Name)" -ForegroundColor Red
            $failureCount++
        }
    } catch {
        Write-Host "❌ 处理出错: $($file.Name) - $($_.Exception.Message)" -ForegroundColor Red
        $failureCount++
    }
    
    Write-Host ""
}

# 清理临时文件
if (Test-Path $tempPythonScript) {
    Remove-Item $tempPythonScript -Force
}

# 显示最终结果
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "              处理完成！" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "🎉 处理结果统计:" -ForegroundColor Green
Write-Host "  ✅ 成功生成字幕: $successCount 个文件" -ForegroundColor Green
if ($failureCount -gt 0) {
    Write-Host "  ❌ 处理失败: $failureCount 个文件" -ForegroundColor Red
}

$finalSrtFiles = Get-ChildItem -Filter "*.srt" -ErrorAction SilentlyContinue
Write-Host ""
Write-Host "📝 当前目录SRT字幕文件总数: $($finalSrtFiles.Count) 个" -ForegroundColor White

Write-Host ""
Write-Host "✨ 所有任务已完成！" -ForegroundColor Green

if (-not $NonInteractive) {
    Read-Host "按任意键退出"
}

