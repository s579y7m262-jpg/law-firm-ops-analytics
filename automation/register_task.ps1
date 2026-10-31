<#
.SYNOPSIS
    注册 / 删除「经营监测仪表盘」的 Windows 计划任务（周一至周五，多时点自动刷新）。

.DESCRIPTION
    用 Register-ScheduledTask 一条命令复现调度，不必在任务计划程序里手动点。
    默认 6 个时点：10:00 / 10:20 / 10:40 / 11:00 / 14:00 / 17:00。
    为什么是这几个：CRM 导出通常在上午落盘，10:00 起每 20 分钟重试一次，
    防"导出还没生成"就白跑；下午 14:00 / 17:00 覆盖补录。
    任务不需要管理员权限（当前用户 + 交互式令牌即可）。

.EXAMPLE
    .\register_task.ps1                       # 注册（同名任务会被覆盖）
    .\register_task.ps1 -Times "10:00,17:00"  # 只跑两个时点
    .\register_task.ps1 -RunNow               # 注册后立刻跑一次
    .\register_task.ps1 -Unregister           # 删除任务
#>
param(
    [string]$TaskName = "经营监测仪表盘",
    [string]$Times    = "10:00,10:20,10:40,11:00,14:00,17:00",
    [switch]$Unregister,
    [switch]$RunNow
)

$ErrorActionPreference = "Stop"

$RepoRoot = Split-Path -Parent $PSScriptRoot
$Bat      = Join-Path $PSScriptRoot "run_monitor.bat"

if ($Unregister) {
    if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
        Write-Host "[ok] 已删除计划任务：$TaskName"
    } else {
        Write-Host "[skip] 没有找到计划任务：$TaskName"
    }
    return
}

if (-not (Test-Path $Bat)) { throw "找不到启动器：$Bat" }

$triggers = foreach ($t in ($Times -split ",")) {
    $at = [datetime]::ParseExact($t.Trim(), "HH:mm", [Globalization.CultureInfo]::InvariantCulture)
    New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday, Tuesday, Wednesday, Thursday, Friday -At $at
}

# 用 $env:ComSpec 取系统命令解释器，避免把绝对路径写死
$action = New-ScheduledTaskAction -Execute $env:ComSpec `
    -Argument "/c `"$Bat`" --silent" `
    -WorkingDirectory $RepoRoot

$settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 30) `
    -MultipleInstances IgnoreNew `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $triggers `
    -Settings $settings `
    -Description "经营监测仪表盘：工作日自动刷新 dashboards/monitor.html" -Force | Out-Null

$task = Get-ScheduledTask -TaskName $TaskName
Write-Host "[ok] 已注册：$TaskName"
Write-Host "     触发器 $($task.Triggers.Count) 个：$Times"
Write-Host "     动作：$($task.Actions[0].Execute) $($task.Actions[0].Arguments)"

if ($RunNow) {
    Start-ScheduledTask -TaskName $TaskName
    Write-Host "[ok] 已触发一次运行（结果见 build\monitor.log）"
}
