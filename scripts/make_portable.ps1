<#
.SYNOPSIS
    打包成「可攜版」——目標主機不需要安裝 Python、Git 或任何東西。

.DESCRIPTION
    把目前這台機器上可以正常運作的 Python（含所有套件）與專案程式碼
    包成一個資料夾，複製到目標主機後直接執行裡面的「啟動.bat」即可。

    這個 Python 是可搬移安裝（python-build-standalone 版面配置），
    換路徑、換機器都能執行——已實測從完全不同的目錄啟動，
    streamlit / pandas / gspread 皆正常，CLI 也跑得起來。

.PARAMETER Destination
    輸出資料夾。預設為專案外層的 ECOCO_可攜版。

.PARAMETER IncludeSecrets
    一併複製 .streamlit\secrets.toml（Google 服務帳戶金鑰）。
    預設不複製——那是憑證，複製到別台機器應該是明確的決定。

.PARAMETER Zip
    打包完再壓成一個 .zip，方便傳輸。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File scripts\make_portable.ps1
    powershell -ExecutionPolicy Bypass -File scripts\make_portable.ps1 -IncludeSecrets -Zip
#>
param(
    [string]$Destination = "",
    [switch]$IncludeSecrets,
    [switch]$Zip
)

$ErrorActionPreference = "Stop"
$projectDir = Split-Path -Parent $PSScriptRoot
if (-not $Destination) {
    $Destination = Join-Path (Split-Path -Parent $projectDir) "ECOCO_可攜版"
}

# ── 找出目前正在用、而且真的有 streamlit 的 Python ──────────────
$pyExe = $null
$candidates = @(
    $env:ECOCO_PYTHON,
    (Join-Path $env:LOCALAPPDATA "Python\pythoncore-3.14-64\python.exe"),
    (Get-Command python -ErrorAction SilentlyContinue).Source
) | Where-Object { $_ }

foreach ($c in $candidates) {
    if (-not (Test-Path $c)) { continue }
    $probe = & $c -c "import streamlit,sys;sys.stdout.write('OK')" 2>$null
    if ($probe -eq "OK") { $pyExe = $c; break }
}
if (-not $pyExe) { throw "找不到裝有 streamlit 的 Python。請先 pip install -r requirements.txt" }

$pyHome = Split-Path -Parent $pyExe
Write-Host "來源 Python : $pyHome"
Write-Host "來源專案   : $projectDir"
Write-Host "輸出到     : $Destination"
Write-Host ""

if (Test-Path $Destination) {
    Write-Host "輸出資料夾已存在，先清空…"
    Remove-Item $Destination -Recurse -Force
}
New-Item -ItemType Directory -Force -Path $Destination | Out-Null

# ── 1. Python（整份可搬移安裝）─────────────────────────────────
Write-Host "[1/4] 複製 Python 執行環境（約 730 MB，需要一點時間）…"
$null = robocopy $pyHome (Join-Path $Destination "python") /E /NFL /NDL /NJH /NJS /NP /MT:8
if ($LASTEXITCODE -ge 8) { throw "Python 複製失敗（robocopy $LASTEXITCODE）" }

# ── 2. 專案程式碼（排除版控、快取、產出與本機資料）──────────────
Write-Host "[2/4] 複製專案程式碼…"
$exclude = @(".git", "__pycache__", "logs", "output", "inbox",
             "history_backup", "tmp_verify", ".tmp_gsheet_tools",
             ".streamlit", "ecoco-manual")
$null = robocopy $projectDir (Join-Path $Destination "app") /E /NFL /NDL /NJH /NJS /NP /MT:8 `
    /XD $exclude /XF "*.pyc" ".automation_state.json"
if ($LASTEXITCODE -ge 8) { throw "專案複製失敗（robocopy $LASTEXITCODE）" }

# 憑證：預設不帶
$stCfg = Join-Path $Destination "app\.streamlit"
New-Item -ItemType Directory -Force -Path $stCfg | Out-Null
Copy-Item (Join-Path $projectDir ".streamlit\config.toml") $stCfg -Force -ErrorAction SilentlyContinue
if ($IncludeSecrets) {
    $src = Join-Path $projectDir ".streamlit\secrets.toml"
    if (Test-Path $src) {
        Copy-Item $src $stCfg -Force
        Write-Host "      已一併複製 secrets.toml（內含 Google 服務帳戶金鑰）"
    }
} else {
    Write-Host "      未複製 secrets.toml。目標主機要連 Google Sheets 的話，"
    Write-Host "      請另外把它放到 app\.streamlit\ 底下，或設成系統環境變數。"
}

# ── 3. 啟動與常駐腳本 ──────────────────────────────────────────
Write-Host "[3/4] 產生啟動腳本…"

$startBat = @'
@echo off
rem ECOCO complaint analysis - portable launcher
rem Nothing needs to be installed on this machine.
setlocal
chcp 65001 >nul
set "PYTHONIOENCODING=utf-8"
set "PYTHONUTF8=1"
set "HERE=%~dp0"
set "ECOCO_PYTHON=%HERE%python\python.exe"
if not exist "%ECOCO_PYTHON%" (
    echo [ERROR] Bundled Python not found: %ECOCO_PYTHON%
    pause
    exit /b 1
)
pushd "%HERE%app"
call "scripts\start_webapp.bat"
popd
endlocal
'@

$regBat = @'
@echo off
rem Register the portable web UI so it keeps running.
setlocal
chcp 65001 >nul
set "HERE=%~dp0"
set "ECOCO_PYTHON=%HERE%python\python.exe"
pushd "%HERE%app"
powershell -ExecutionPolicy Bypass -File "scripts\register_webapp_task.ps1"
popd
pause
endlocal
'@

[IO.File]::WriteAllText((Join-Path $Destination "啟動.bat"), $startBat.Replace("`n", "`r`n"), [Text.Encoding]::ASCII)
[IO.File]::WriteAllText((Join-Path $Destination "註冊常駐.bat"), $regBat.Replace("`n", "`r`n"), [Text.Encoding]::ASCII)

$readme = @"
ECOCO 客訴分析平台 — 可攜版
============================================================

這個資料夾自帶 Python 與所有套件，目標主機不需要安裝任何東西。

【怎麼用】
  1. 把整個資料夾複製到目標主機（例如 D:\ECOCO_可攜版）
  2. 雙擊「啟動.bat」
  3. 視窗會印出網址，預設 http://localhost:8501
     同網段的其他電腦用 http://<這台主機IP>:8501

【要它一直活著（關掉視窗也不停）】
  雙擊「註冊常駐.bat」
  會註冊成 Windows 排程工作，掛掉最多 5 分鐘自動拉回來。

【Google Sheets 憑證】
  打包時預設不含 secrets.toml。要用歷史紀錄與趨勢儀表板的話，
  把 secrets.toml 放到  app\.streamlit\  底下。
  沒有憑證也能用「上傳檔案 → 分析 → 下載」這條主線。

【檢查設定】
  在這個資料夾開命令提示字元，執行：
    python\python.exe -m automation.cli doctor  （先 cd app）

打包時間：$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
來源主機：$env:COMPUTERNAME
"@
[IO.File]::WriteAllText((Join-Path $Destination "README.txt"), $readme, [Text.Encoding]::UTF8)

# ── 4. 驗收 ────────────────────────────────────────────────────
Write-Host "[4/4] 驗收打包結果…"
$bundledPy = Join-Path $Destination "python\python.exe"
$probe = & $bundledPy -c "import streamlit,pandas,gspread,sys;sys.stdout.write('OK')" 2>$null
if ($probe -ne "OK") { throw "打包後的 Python 無法載入必要套件" }

$size = (Get-ChildItem $Destination -Recurse -File -ErrorAction SilentlyContinue | Measure-Object Length -Sum).Sum / 1MB
Write-Host ""
Write-Host ("完成：{0}  （{1:N0} MB）" -f $Destination, $size)
Write-Host "  啟動.bat        直接啟動網頁介面"
Write-Host "  註冊常駐.bat    註冊成排程工作，關掉視窗也不停"
Write-Host "  README.txt      使用說明"

if ($Zip) {
    $zipPath = "$Destination.zip"
    Write-Host ""
    Write-Host "壓縮中…（檔案大，需要幾分鐘）"
    if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
    Compress-Archive -Path "$Destination\*" -DestinationPath $zipPath -CompressionLevel Optimal
    $zsize = (Get-Item $zipPath).Length / 1MB
    Write-Host ("已壓縮：{0}  （{1:N0} MB）" -f $zipPath, $zsize)
}
