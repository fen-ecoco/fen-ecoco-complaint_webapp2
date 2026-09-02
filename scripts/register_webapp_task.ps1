<#
.SYNOPSIS
    讓 ECOCO 客訴分析「網頁介面」自動常駐，不會因為終端機視窗關掉就停。

.DESCRIPTION
    依序嘗試兩種方式，用得成哪個就用哪個：
      1. Windows 排程工作（登入時啟動）—— 需要系統管理員權限
      2. 啟動資料夾捷徑（登入時啟動）—— 不需要任何權限
    加 -AtStartup 可改成「開機即啟動、不必登入」，但那一定要管理員權限。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File scriptsegister_webapp_task.ps1
    powershell -ExecutionPolicy Bypass -File scriptsegister_webapp_task.ps1 -Port 8600
    powershell -ExecutionPolicy Bypass -File scriptsegister_webapp_task.ps1 -AtStartup   # 需管理員
    powershell -ExecutionPolicy Bypass -File scriptsegister_webapp_task.ps1 -Remove
#>
param(
    [string]$TaskName = "ECOCO客訴分析網頁",
    [int]$Port = 8501,
    [switch]$AtStartup,
    [switch]$Remove
)

$ErrorActionPreference = "Stop"
$projectDir  = Split-Path -Parent $PSScriptRoot
$batPath     = Join-Path $PSScriptRoot "start_webapp.bat"
$startupDir  = [Environment]::GetFolderPath("Startup")
$lnkPath     = Join-Path $startupDir "$TaskName.lnk"

if ($Remove) {
    $done = @()
    try { Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false; $done += "排程工作" } catch { }
    if (Test-Path $lnkPath) { Remove-Item $lnkPath -Force; $done += "啟動資料夾捷徑" }
    if ($done.Count) { Write-Host ("已移除：" + ($done -join "、")) } else { Write-Host "沒有找到已註冊的項目。" }
    return
}

if (-not (Test-Path $batPath)) { throw "找不到 $batPath" }

# ── 方式 1：排程工作 ─────────────────────────────────────────
$registered = $false
try {
    $action = New-ScheduledTaskAction -Execute "cmd.exe" `
        -Argument "/c set ECOCO_PORT=$Port && `"$batPath`"" `
        -WorkingDirectory $projectDir
    $trigger = if ($AtStartup) { New-ScheduledTaskTrigger -AtStartup } else { New-ScheduledTaskTrigger -AtLogOn }
    # ExecutionTimeLimit 0 = 不限時；否則排程器預設 3 天後會終止這個長駐行程
    $settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -DontStopOnIdleEnd `
        -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1) `
        -ExecutionTimeLimit ([TimeSpan]::Zero)
    Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger `
        -Settings $settings -Description "ECOCO 客訴分析網頁介面" -Force -ErrorAction Stop | Out-Null
    $registered = $true
    $when = if ($AtStartup) { "開機時啟動" } else { "登入時啟動" }
    Write-Host "已註冊排程工作：$TaskName（$when，埠 $Port）"
    Write-Host "  立即啟動： Start-ScheduledTask -TaskName '$TaskName'"
    Write-Host "  查看狀態： Get-ScheduledTaskInfo -TaskName '$TaskName'"
} catch {
    Write-Host "排程工作註冊失敗（$($_.Exception.Message.Trim())）"
    if ($AtStartup) {
        Write-Host "「開機時啟動」一定要系統管理員權限，請用管理員身分重跑。"
        return
    }
    Write-Host "改用不需要權限的方式：啟動資料夾捷徑。"
}

# ── 方式 2：啟動資料夾捷徑 ───────────────────────────────────
if (-not $registered) {
    $shell = New-Object -ComObject WScript.Shell
    $lnk = $shell.CreateShortcut($lnkPath)
    $lnk.TargetPath       = "cmd.exe"
    $lnk.Arguments        = "/c set ECOCO_PORT=$Port && `"$batPath`""
    $lnk.WorkingDirectory = $projectDir
    $lnk.WindowStyle      = 7          # 最小化
    $lnk.Description      = "ECOCO 客訴分析網頁介面"
    $lnk.Save()
    Write-Host "已建立啟動資料夾捷徑：$lnkPath"
    Write-Host "  下次登入這台機器時會自動啟動（最小化視窗）。"
    Write-Host "  要現在就啟動： Start-Process -FilePath `"$batPath`""
}

Write-Host ""
Write-Host "網址： http://localhost:$Port"
Write-Host "移除： powershell -ExecutionPolicy Bypass -File scriptsegister_webapp_task.ps1 -Remove"
