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
    [switch]$Daily,            # 只在上班時段執行
    [string]$StartTime = "08:00",
    [string]$StopTime  = "19:00",
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
    try { Unregister-ScheduledTask -TaskName "$TaskName-停止" -Confirm:$false; $done += "每日停止工作" } catch { }
    if (Test-Path $lnkPath) { Remove-Item $lnkPath -Force; $done += "啟動資料夾捷徑" }
    if ($done.Count) { Write-Host ("已移除：" + ($done -join "、")) } else { Write-Host "沒有找到已註冊的項目。" }
    return
}

if (-not (Test-Path $batPath)) { throw "找不到 $batPath" }

# ── 方式 1：排程工作（自我修復的長駐） ──────────────────────
# 觸發器用 -Once + 每 5 分鐘重複：一般使用者就能註冊
#（AtLogOn / AtStartup 都需要管理員權限，實測會被拒）。
# MultipleInstances = IgnoreNew：服務還活著時後續觸發直接略過；
# 服務掛了下一次觸發就把它拉回來。
# 搭配 start_webapp.bat 的 ECOCO_NO_PORT_HUNT=1，
# 埠已被自己佔用時乾淨結束，不會另外開一份在別的埠。
$registered = $false
try {
    $action = New-ScheduledTaskAction -Execute "cmd.exe" `
        -Argument "/c set ECOCO_PORT=$Port && set ECOCO_NO_PORT_HUNT=1 && `"$batPath`"" `
        -WorkingDirectory $projectDir

    if ($AtStartup) {
        $trigger = New-ScheduledTaskTrigger -AtStartup
    } elseif ($Daily) {
        # 上班時段模式：每天 StartTime 啟動，期間每 5 分鐘自我修復，
        # 到 StopTime 由另一個工作停掉，把記憶體還回去。
        $span = [datetime]$StopTime - [datetime]$StartTime
        if ($span.TotalMinutes -le 0) { throw "StopTime 必須晚於 StartTime" }
        $trigger = New-ScheduledTaskTrigger -Daily -At $StartTime
        $trigger.Repetition = (New-ScheduledTaskTrigger -Once -At (Get-Date) `
            -RepetitionInterval (New-TimeSpan -Minutes 5) `
            -RepetitionDuration $span).Repetition
    } else {
        $trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) `
            -RepetitionInterval (New-TimeSpan -Minutes 5) `
            -RepetitionDuration (New-TimeSpan -Days 3650)
    }

    # ExecutionTimeLimit 0 = 不限時；否則排程器預設 3 天後會終止長駐行程
    $settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -DontStopOnIdleEnd `
        -MultipleInstances IgnoreNew `
        -ExecutionTimeLimit ([TimeSpan]::Zero)

    Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger `
        -Settings $settings -Description "ECOCO 客訴分析網頁介面" -Force -ErrorAction Stop | Out-Null
    $registered = $true

    $when = if ($AtStartup) { "開機時啟動" }
            elseif ($Daily)  { "每天 $StartTime 啟動、$StopTime 停止" }
            else             { "每 5 分鐘檢查，沒在跑就拉起來" }
    Write-Host "已註冊排程工作：$TaskName（$when，埠 $Port）"

    if ($Daily) {
        $stopAction  = New-ScheduledTaskAction -Execute "powershell.exe" `
            -Argument "-NoProfile -Command `"Stop-ScheduledTask -TaskName '$TaskName'`""
        $stopTrigger = New-ScheduledTaskTrigger -Daily -At $StopTime
        Register-ScheduledTask -TaskName "$TaskName-停止" -Action $stopAction `
            -Trigger $stopTrigger -Description "每天 $StopTime 停止 ECOCO 網頁介面" `
            -Force -ErrorAction Stop | Out-Null
        Write-Host "已註冊每日停止工作：$TaskName-停止（每天 $StopTime）"
    }

    Start-ScheduledTask -TaskName $TaskName
    Write-Host "已立即啟動一次。"
    Write-Host "  查看狀態： Get-ScheduledTaskInfo -TaskName '$TaskName'"
    Write-Host "  手動停止： Stop-ScheduledTask -TaskName '$TaskName'"
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
