# 以管理员身份运行此脚本
# 为所有 UWP 应用启用 loopback 豁免

# 检查是否以管理员身份运行
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "请以管理员身份运行此脚本！"
    pause
    exit
}

Write-Host "正在获取所有 AppContainer 应用..." -ForegroundColor Cyan

# 获取所有 AppX 包的 SID
$packages = Get-AppxPackage -AllUsers | Where-Object { $_.PackageFamilyName }

$count = 0
foreach ($package in $packages) {
    try {
        # 获取包的 SID
        $sid = (Get-AppxPackage -Name $package.Name -AllUsers | Get-AppxPackageManifest).Package.Applications.Application | ForEach-Object {
            $familyName = $package.PackageFamilyName
            # 使用 CheckNetIsolation 添加 loopback 豁免
            $result = CheckNetIsolation.exe LoopbackExempt -a -n="$familyName" 2>&1
            if ($LASTEXITCODE -eq 0) {
                $count++
                Write-Host "已添加: $familyName" -ForegroundColor Green
            }
        }
    }
    catch {
        # 静默处理错误
    }
}

# 也可以直接通过 PackageFamilyName 添加
Write-Host "`n正在通过 PackageFamilyName 批量添加..." -ForegroundColor Cyan

$allPackages = Get-AppxPackage -AllUsers
foreach ($pkg in $allPackages) {
    if ($pkg.PackageFamilyName) {
        CheckNetIsolation.exe LoopbackExempt -a -n="$($pkg.PackageFamilyName)" 2>$null
    }
}

Write-Host "`n完成！所有 UWP 应用已添加 loopback 豁免。" -ForegroundColor Green
Write-Host "请重启 Clash Verge 以使更改生效。" -ForegroundColor Yellow

pause
