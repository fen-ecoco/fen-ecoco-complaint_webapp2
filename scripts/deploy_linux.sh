#!/usr/bin/env bash
# ============================================================
#  ECOCO 客訴分析平台 — Ubuntu / Debian 部署腳本
#
#  用法（在目標主機上）：
#      bash scripts/deploy_linux.sh
#      bash scripts/deploy_linux.sh --port 8501 --user-service
#
#  做的事：
#    1. 檢查 Python 版本（需要 3.11 以上）
#    2. 建立 venv 並安裝 requirements.txt
#    3. 安裝中文字型（PDF 與圖表需要，沒有會變成空白方框）
#    4. 註冊 systemd 服務，開機自動啟動、當掉自動重啟
#
#  預設裝成「系統服務」（需要 sudo）。
#  沒有 sudo 時加 --user-service 改裝成使用者服務。
# ============================================================
set -euo pipefail

PORT=8501
USER_SERVICE=0
SERVICE_NAME="ecoco-webapp"

while [[ $# -gt 0 ]]; do
    case "$1" in
        --port)         PORT="$2"; shift 2 ;;
        --user-service) USER_SERVICE=1; shift ;;
        --name)         SERVICE_NAME="$2"; shift 2 ;;
        -h|--help)      sed -n '2,20p' "$0"; exit 0 ;;
        *) echo "未知參數：$1" >&2; exit 1 ;;
    esac
done

PROJECT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
VENV="$PROJECT_DIR/.venv"
cd "$PROJECT_DIR"

say() { printf '\n\033[1m%s\033[0m\n' "$*"; }

# ── 1. Python ────────────────────────────────────────────────
say "[1/4] 檢查 Python"
PY=""
for cand in python3.13 python3.12 python3.11 python3; do
    if command -v "$cand" >/dev/null 2>&1; then
        ver="$("$cand" -c 'import sys;print("%d.%d"%sys.version_info[:2])')"
        major="${ver%%.*}"; minor="${ver##*.}"
        if [[ "$major" -eq 3 && "$minor" -ge 11 ]]; then PY="$cand"; break; fi
    fi
done
if [[ -z "$PY" ]]; then
    echo "找不到 Python 3.11 以上。請先安裝：" >&2
    echo "    sudo apt update && sudo apt install -y python3 python3-venv python3-pip" >&2
    exit 1
fi
echo "使用 $PY（$("$PY" --version)）"

# ── 2. venv 與套件 ───────────────────────────────────────────
say "[2/4] 建立虛擬環境並安裝套件"
if [[ ! -d "$VENV" ]]; then
    "$PY" -m venv "$VENV" || {
        echo "建立 venv 失敗。可能需要： sudo apt install -y python3-venv" >&2
        exit 1
    }
fi
"$VENV/bin/python" -m pip install --upgrade pip >/dev/null
# 有 lock 檔就用它：requirements.txt 沒有鎖版本，直接裝會拿到當下最新版，
# 不同時間部署的機器會拿到不同組合。
if [[ -f requirements-lock.txt ]]; then
    echo "使用 requirements-lock.txt（已驗證的版本組合）"
    "$VENV/bin/python" -m pip install -r requirements-lock.txt
else
    "$VENV/bin/python" -m pip install -r requirements.txt
fi
"$VENV/bin/python" -c "import streamlit,pandas,gspread;print('套件就緒')"

# ── 3. 中文字型 ──────────────────────────────────────────────
say "[3/4] 檢查中文字型"
# 直接檢查檔案路徑，跟 _ensure_cjk_font() 的判斷一致。
# fc-list 需要 fontconfig 快取正確，字型檔明明在也可能查不到。
FONT_FOUND=""
for f in /usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc          /usr/share/fonts/opentype/noto/NotoSansCJK-Medium.ttc          /usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc          /usr/share/fonts/noto-cjk/NotoSansCJK-Regular.ttc          /usr/share/fonts/truetype/wqy/wqy-microhei.ttc; do
    [[ -f "$f" ]] && FONT_FOUND="$f" && break
done
if [[ -n "$FONT_FOUND" ]]; then
    echo "已有中文字型：$FONT_FOUND"
else
    echo "找不到中文字型，PDF 與圖表的中文會變成空白方框。"
    # sudo 不一定免密碼，失敗不能讓整個部署中斷（腳本有 set -e）
    if sudo -n true 2>/dev/null; then
        echo "安裝 fonts-noto-cjk…"
        sudo -n apt-get update -qq && sudo -n apt-get install -y fonts-noto-cjk && fc-cache -f || true
    else
        echo "沒有免密碼 sudo，請自行執行： sudo apt install -y fonts-noto-cjk"
        echo "（程式在找不到系統字型時會自動下載一份到暫存資料夾，仍可運作）"
    fi
fi

# ── 4. systemd 服務 ──────────────────────────────────────────
say "[4/4] 註冊 systemd 服務"
EXEC="$VENV/bin/python -m streamlit run complaint_webapp.py \
--server.port $PORT --server.address 0.0.0.0 \
--server.headless true --browser.gatherUsageStats false"

UNIT_BODY="[Unit]
Description=ECOCO 客訴分析平台（網頁介面）
After=network-online.target
Wants=network-online.target

[Service]
Type=simple
WorkingDirectory=$PROJECT_DIR
Environment=PYTHONUNBUFFERED=1
Environment=PYTHONIOENCODING=utf-8
EnvironmentFile=-$PROJECT_DIR/.env
ExecStart=$EXEC
Restart=always
RestartSec=10

[Install]
WantedBy=WANTED_TARGET
"

if [[ "$USER_SERVICE" -eq 1 ]]; then
    UNIT_DIR="$HOME/.config/systemd/user"
    mkdir -p "$UNIT_DIR"
    echo "${UNIT_BODY//WANTED_TARGET/default.target}" > "$UNIT_DIR/$SERVICE_NAME.service"
    systemctl --user daemon-reload
    systemctl --user enable --now "$SERVICE_NAME"
    echo
    echo "已啟動使用者服務：$SERVICE_NAME"
    echo "  狀態： systemctl --user status $SERVICE_NAME"
    echo "  日誌： journalctl --user -u $SERVICE_NAME -f"
    echo
    echo "注意：使用者服務在你登出後會停止。要讓它一直跑，請管理者執行："
    echo "  sudo loginctl enable-linger $USER"
else
    TMP_UNIT="$(mktemp)"
    echo "${UNIT_BODY//WANTED_TARGET/multi-user.target}" > "$TMP_UNIT"
    sudo install -m 644 "$TMP_UNIT" "/etc/systemd/system/$SERVICE_NAME.service"
    rm -f "$TMP_UNIT"
    sudo systemctl daemon-reload
    sudo systemctl enable --now "$SERVICE_NAME"
    echo
    echo "已啟動系統服務：$SERVICE_NAME（開機自動啟動、當掉自動重啟）"
    echo "  狀態： sudo systemctl status $SERVICE_NAME"
    echo "  日誌： sudo journalctl -u $SERVICE_NAME -f"
    echo "  重啟： sudo systemctl restart $SERVICE_NAME"
fi

IP="$(hostname -I 2>/dev/null | awk '{print $1}')"
echo
echo "============================================================"
echo "  本機   ： http://localhost:$PORT"
[[ -n "$IP" ]] && echo "  內網   ： http://$IP:$PORT"
echo "============================================================"
echo
echo "若同網段其他電腦連不上，可能是防火牆："
echo "  sudo ufw allow $PORT/tcp"
echo
echo "檢查憑證與試算表設定："
echo "  $VENV/bin/python -m automation.cli doctor"
