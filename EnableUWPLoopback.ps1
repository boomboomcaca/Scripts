# 以管理员身份运行此脚本
# 为所有 UWP 应用启用 loopback 豁免
param (
    [switch]$InstallTask,
    [switch]$UninstallTask
)

# 检查是否以管理员身份运行
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "请以管理员身份运行此脚本！"
    exit
}

if ($UninstallTask) {
    Write-Host "正在删除计划任务..." -ForegroundColor Cyan
    $delCount = 0
    $tasks = @("EnableUWPLoopback_Combined", "EnableUWPLoopback", "EnableUWPLoopback_MSI", "EnableUWPLoopback_Fallback")
    
    foreach ($task in $tasks) {
        schtasks /query /tn "$task" 2>$null | Out-Null
        if ($LASTEXITCODE -eq 0) {
            schtasks /delete /tn "$task" /f 2>$null | Out-Null
            if ($LASTEXITCODE -eq 0) {
                Write-Host "[OK] 已删除任务: $task" -ForegroundColor Green
                $delCount++
            } else {
                Write-Host "[ERROR] 删除失败: $task" -ForegroundColor Red
            }
        } else {
            Write-Host "[SKIP] 未找到任务: $task" -ForegroundColor DarkGray
        }
    }
    
    Write-Host ""
    if ($delCount -gt 0) {
        Write-Host "成功删除了 $delCount 个计划任务。" -ForegroundColor Green
    } else {
        Write-Host "没有发现相关的计划任务，无需清理。" -ForegroundColor Yellow
    }
    exit
}

if ($InstallTask) {
    $taskName = "EnableUWPLoopback_Combined"
    $scriptPath = $MyInvocation.MyCommand.Definition

    Write-Host "正在清理旧任务..." -ForegroundColor Cyan
    schtasks /delete /tn "EnableUWPLoopback" /f 2>$null
    schtasks /delete /tn "EnableUWPLoopback_MSI" /f 2>$null
    schtasks /delete /tn "EnableUWPLoopback_Fallback" /f 2>$null
    # 也尝试清理自身（如果是更新）
    schtasks /delete /tn "$taskName" /f 2>$null
    Write-Host "旧任务清理完毕。" -ForegroundColor Green

    $xml = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Description>自动为 UWP 应用添加回环代理豁免 (合并版: 包含 UWP/MSI 安装监控及定时检查)</Description>
    <URI>\$taskName</URI>
  </RegistrationInfo>
  <Triggers>
    <EventTrigger>
      <Enabled>true</Enabled>
      <Subscription>&lt;QueryList&gt;&lt;Query Id="0" Path="Microsoft-Windows-AppXDeploymentServer/Operational"&gt;&lt;Select Path="Microsoft-Windows-AppXDeploymentServer/Operational"&gt;*[System[(EventID=19)]]&lt;/Select&gt;&lt;/Query&gt;&lt;/QueryList&gt;</Subscription>
    </EventTrigger>
    <EventTrigger>
      <Enabled>true</Enabled>
      <Subscription>&lt;QueryList&gt;&lt;Query Id="0" Path="Application"&gt;&lt;Select Path="Application"&gt;*[System[Provider[@Name='MsiInstaller'] and (EventID=11707)]]&lt;/Select&gt;&lt;/Query&gt;&lt;/QueryList&gt;</Subscription>
    </EventTrigger>
    <CalendarTrigger>
      <StartBoundary>2023-01-01T00:00:00</StartBoundary>
      <Enabled>true</Enabled>
      <ScheduleByDay>
        <DaysInterval>1</DaysInterval>
      </ScheduleByDay>
      <Repetition>
        <Interval>PT6H</Interval>
        <Duration>P1D</Duration>
      </Repetition>
    </CalendarTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <LogonType>InteractiveToken</LogonType>
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>Queue</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>true</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT72H</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>PowerShell.exe</Command>
      <Arguments>-ExecutionPolicy Bypass -WindowStyle Hidden -File "$scriptPath"</Arguments>
    </Exec>
  </Actions>
</Task>
"@

    $xmlPath = "$env:TEMP\task.xml"
    $xml | Out-File $xmlPath -Encoding Unicode

    Write-Host "正在创建合并后的计划任务: $taskName..." -ForegroundColor Cyan
    schtasks /create /tn "$taskName" /xml "$xmlPath" /f
    Remove-Item $xmlPath

    if ($LASTEXITCODE -eq 0) {
        Write-Host "`n[成功] 计划任务已安装！" -ForegroundColor Green
        Write-Host "触发条件:"
        Write-Host "1. UWP 应用安装 (Event 19)"
        Write-Host "2. MSI 软件安装 (Event 11707)"
        Write-Host "3. 每 6 小时自动运行"
    } else {
        Write-Host "`n[失败] 创建任务失败，请检查错误信息。" -ForegroundColor Red
    }
    
    exit
}

Write-Host "正在获取所有 AppContainer 应用..." -ForegroundColor Cyan

# 获取所有 AppX 包的 SID
$packages = Get-AppxPackage -AllUsers | Where-Object { $_.PackageFamilyName }

$existing = @{}
try {
    $existingOutput = CheckNetIsolation.exe LoopbackExempt -s 2>$null
    foreach ($line in $existingOutput) {
        # 同时支持中文和英文系统
        if ($line -match '^\s*(名称|Name):\s*(.+)\s*$') {
            $existing[$matches[2]] = $true
        }
    }
}
catch {
}

$seen = @{}

$count = 0
foreach ($package in $packages) {
    try {
        # 获取包的 SID
        $familyName = $package.PackageFamilyName
        if (-not $familyName) { continue }
        if ($seen.ContainsKey($familyName)) { continue }
        $seen[$familyName] = $true
        if ($existing.ContainsKey($familyName)) { continue }
        # 使用 CheckNetIsolation 添加 loopback 豁免
        $result = CheckNetIsolation.exe LoopbackExempt -a -n="$familyName" 2>&1
        if ($LASTEXITCODE -eq 0) {
            $count++
            $existing[$familyName] = $true
            Write-Host "已添加: $familyName" -ForegroundColor Green
        }
    }
    catch {
        # 静默处理错误
    }
}

Write-Host "`n正在清理失效的 exemption..." -ForegroundColor Cyan
$orphanCount = 0
try {
    $currentExemptions = CheckNetIsolation.exe LoopbackExempt -s 2>$null
    
    for ($i = 0; $i -lt $currentExemptions.Count; $i++) {
        $line = $currentExemptions[$i]
        
        # 同时支持中文和英文系统
        if ($line -match '^\s*(名称|Name):\s*AppContainer NOT FOUND\s*$') {
            # 找到孤儿记录，寻找下一行的 SID
            if (($i + 1) -lt $currentExemptions.Count) {
                $nextNode = $currentExemptions[$i + 1]
                if ($nextNode -match '^\s*SID:\s*(S-1-15-\d+[-\d]+)\s*$') {
                    $sidToRemove = $matches[1]
                    Write-Host "移除失效 SID: $sidToRemove" -ForegroundColor Yellow
                    # 使用 Start-Process 并在后台运行以减少交互干扰可能性，虽然这主要是命令行工具
                    $proc = Start-Process CheckNetIsolation.exe -ArgumentList "LoopbackExempt -d -p=""$sidToRemove""" -NoNewWindow -PassThru -Wait
                    $orphanCount++
                }
            }
        }
    }
} catch {
    Write-Host "清理过程出错: $_" -ForegroundColor Red
}

if ($orphanCount -gt 0) {
    Write-Host "已清理 $orphanCount 个失效条目。" -ForegroundColor Green
} else {
    Write-Host "未发现失效条目。" -ForegroundColor DarkGray
}

Write-Host "`n完成！所有 UWP 应用已添加 loopback 豁免。" -ForegroundColor Green
Write-Host "请重启 Clash Verge 以使更改生效。" -ForegroundColor Yellow


