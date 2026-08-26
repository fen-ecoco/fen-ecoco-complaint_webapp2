<#
.SYNOPSIS
    把 ECOCO 客訴分析註冊成 Windows 排程工作。

.DESCRIPTION
    預設每天 08:30 執行一次 scriptsun_analysis.bat。
    需以「系統管理員」身分執行 PowerShell。

.EXAMPLE
    .\scriptsegister_task.ps1
    .\scriptsegister_task.ps1 -Time "07:00"
    .\scriptsegister_task.ps1 -Remove          # 移除排程
#>
param(
    [string]$TaskName = "ECOCO客訴分析",
    [string]$Time     = "08:30",
    [switch]$Remove
)

$ErrorActionPreference = "Stop"
$projectDir = Split-Path -Parent $PSScriptRoot
$batPath    = Join-Path $PSScriptRoot "run_analysis.bat"

if ($Remove) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    Write-Host "已移除排程工作：$TaskName"
    return
}

if (-not (Test-Path $batPath)) { throw "找不到 $batPath" }

$action   = New-ScheduledTaskAction -Execute $batPath -WorkingDirectory $projectDir
$trigger  = New-ScheduledTaskTrigger -Daily -At $Time
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -DontStopOnIdleEnd -ExecutionTimeLimit (New-TimeSpan -Hours 2) -RestartCount 2 -RestartInterval (New-TimeSpan -Minutes 10)

$desc = "每日自動分析客訴並產出報告（ECOCO）"
Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings -Description $desc -Force | Out-Null

Write-Host "已註冊排程工作：$TaskName（每天 $Time）"
Write-Host "  執行檔案：$batPath"
Write-Host "  工作目錄：$projectDir"
Write-Host ""
Write-Host "立即測試一次： Start-ScheduledTask -TaskName $TaskName"
Write-Host "查看執行結果： Get-ScheduledTaskInfo -TaskName $TaskName"
Write-Host "日誌位置：     $projectDir\logs\"
