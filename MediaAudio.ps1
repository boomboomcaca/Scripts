# 两遍响度标准化 + 音轨完整性校验（供 Convert_to_Mp4_Srt.ps1 / Watch_Downloads.ps1 共用）

$script:LoudnormI = '-16'
$script:LoudnormTP = '-1.5'
$script:LoudnormLRA = '11'

function Test-LoudnormApplied {
    param([string]$FilePath)

    try {
        $tag = ffprobe -v quiet -show_entries format_tags=loudnorm_applied -of default=nw=1:nk=1 "$FilePath" 2>$null
        return ("$tag".Trim() -eq '1')
    }
    catch {
        return $false
    }
}

function Test-AudioStreamHealthy {
    param([string]$FilePath)

    $errFile = Join-Path $env:TEMP ("audio_check_{0}_{1}.log" -f $PID, [Guid]::NewGuid().ToString('N'))
    try {
        $ffmpegArgs = @(
            '-hide_banner', '-v', 'error', '-xerror',
            '-i', "`"$FilePath`"",
            '-vn', '-sn', '-dn', '-f', 'null', 'NUL'
        )
        $process = Start-Process -FilePath 'ffmpeg' -ArgumentList $ffmpegArgs -Wait -PassThru -NoNewWindow -RedirectStandardError $errFile
        if ($process.ExitCode -ne 0) { return $false }
        $err = Get-Content -LiteralPath $errFile -Raw -ErrorAction SilentlyContinue
        if ($err -and $err -match 'Invalid data|channel element|Decode error') { return $false }
        return $true
    }
    catch {
        return $false
    }
    finally {
        Remove-Item -LiteralPath $errFile -Force -ErrorAction SilentlyContinue
    }
}

function Get-LoudnormMeasurement {
    param([string]$FilePath)

    $logFile = Join-Path $env:TEMP ("loudnorm_measure_{0}_{1}.log" -f $PID, [Guid]::NewGuid().ToString('N'))
    try {
        $ffmpegArgs = @(
            '-hide_banner', '-nostats',
            '-i', "`"$FilePath`"",
            '-vn', '-sn', '-dn',
            '-af', "loudnorm=I=$($script:LoudnormI):TP=$($script:LoudnormTP):LRA=$($script:LoudnormLRA):print_format=json",
            '-f', 'null', 'NUL'
        )
        $process = Start-Process -FilePath 'ffmpeg' -ArgumentList $ffmpegArgs -Wait -PassThru -NoNewWindow -RedirectStandardError $logFile
        if ($process.ExitCode -ne 0) { return $null }

        $text = Get-Content -LiteralPath $logFile -Raw -ErrorAction SilentlyContinue
        if (-not $text -or $text -notmatch '(?s)\{\s*"input_i".*?\}') { return $null }

        $measured = $Matches[0] | ConvertFrom-Json
        foreach ($prop in @('input_i', 'input_tp', 'input_lra', 'input_thresh', 'target_offset')) {
            $value = [string]$measured.$prop
            if (-not $value -or $value -match 'inf') { return $null }
        }
        return $measured
    }
    catch {
        return $null
    }
    finally {
        Remove-Item -LiteralPath $logFile -Force -ErrorAction SilentlyContinue
    }
}

function ConvertTo-LoudnormFilter {
    param($Measurement)

    return (
        "loudnorm=I=$($script:LoudnormI):TP=$($script:LoudnormTP):LRA=$($script:LoudnormLRA)" +
        ":measured_I=$($Measurement.input_i)" +
        ":measured_LRA=$($Measurement.input_lra)" +
        ":measured_TP=$($Measurement.input_tp)" +
        ":measured_thresh=$($Measurement.input_thresh)" +
        ":offset=$($Measurement.target_offset)" +
        ':linear=true'
    )
}

function Get-LoudnormAudioFfmpegArgs {
    param([string]$FilePath)

    Write-Host '  📏 第一遍：测量音频响度...' -ForegroundColor Cyan
    $measured = Get-LoudnormMeasurement -FilePath $FilePath
    if (-not $measured) {
        Write-Host '  ⚠️ 响度测量失败，将仅转为 AAC（不标准化）' -ForegroundColor Yellow
        return @('-c:a', 'aac', '-b:a', '192k', '-ar', '48000')
    }

    Write-Host "  🔊 第二遍：应用响度标准化 (输入 $($measured.input_i) LUFS → $($script:LoudnormI) LUFS)..." -ForegroundColor Cyan
    return @(
        '-c:a', 'aac',
        '-b:a', '192k',
        '-ar', '48000',
        '-af', (ConvertTo-LoudnormFilter -Measurement $measured),
        '-metadata', 'loudnorm_applied=1'
    )
}

function Invoke-TwoPassAudioLoudnorm {
    param([string]$FilePath)

    $name = [System.IO.Path]::GetFileName($FilePath)
    if (Test-LoudnormApplied -FilePath $FilePath) {
        Write-Host "  🔉 已标准化过，跳过: $name" -ForegroundColor DarkGray
        return $true
    }

    Write-Host "  🔊 音频响度标准化: $name" -ForegroundColor White
    Write-Host '  📏 第一遍：测量音频响度...' -ForegroundColor Cyan
    $measured = Get-LoudnormMeasurement -FilePath $FilePath
    if (-not $measured) {
        Write-Host '  ❌ 响度测量失败，保留原音轨' -ForegroundColor Red
        return $false
    }

    $tempFile = [System.IO.Path]::Combine(
        [System.IO.Path]::GetDirectoryName($FilePath),
        [System.IO.Path]::GetFileNameWithoutExtension($FilePath) + '.loudnorm.temp.mp4'
    )

    Write-Host "  🔊 第二遍：应用响度标准化 (输入 $($measured.input_i) LUFS → $($script:LoudnormI) LUFS)..." -ForegroundColor Cyan
    try {
        $ffmpegArgs = @(
            '-i', "`"$FilePath`"",
            '-c:v', 'copy',
            '-c:a', 'aac',
            '-b:a', '192k',
            '-ar', '48000',
            '-af', (ConvertTo-LoudnormFilter -Measurement $measured),
            '-map_metadata', '0',
            '-metadata', 'loudnorm_applied=1',
            '-movflags', '+faststart',
            '-y',
            "`"$tempFile`""
        )

        $process = Start-Process -FilePath 'ffmpeg' -ArgumentList $ffmpegArgs -Wait -PassThru -NoNewWindow
        if ($process.ExitCode -ne 0 -or -not (Test-Path -LiteralPath $tempFile)) {
            Write-Host '  ❌ 音频标准化失败，保留原文件' -ForegroundColor Red
            if (Test-Path -LiteralPath $tempFile) { Remove-Item -LiteralPath $tempFile -Force }
            return $false
        }

        Write-Host '  🔍 校验输出音轨...' -ForegroundColor Cyan
        if (-not (Test-AudioStreamHealthy -FilePath $tempFile)) {
            Write-Host '  ❌ 输出音轨校验失败，保留原文件' -ForegroundColor Red
            Remove-Item -LiteralPath $tempFile -Force
            return $false
        }

        Remove-Item -LiteralPath $FilePath -Force
        Move-Item -LiteralPath $tempFile -Destination $FilePath
        Write-Host '  ✅ 音频标准化完成' -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "  ⚠️ 音频标准化出错，保留原文件: $($_.Exception.Message)" -ForegroundColor Yellow
        if (Test-Path -LiteralPath $tempFile) { Remove-Item -LiteralPath $tempFile -Force }
        return $false
    }
}
