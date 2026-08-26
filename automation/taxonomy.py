"""問題分類法（唯一來源）。

由 complaint_webapp.py 原樣抽出，供 Streamlit 介面與無介面排程共用。
新增／調整分類請只改這裡。
"""

import re

TOPIC_DETAIL_MAP = {
    "APP使用問題類型": [
        "APP畫面顯示與機台狀態不符",
        "APP商家頁面空白",
        "APP點數顯示異常",
        "APP多重異常狀況",
        "app畫面顯示與機台狀態不符",
        "app點數顯示異常",
        "app商家頁面空白",
        "app多重異常狀況",
    ],
    "APP帳號設定問題類型": [
        "忘記密碼/無法重設密碼",
        "帳號資訊修改/設定",
        "無法接收簡訊驗證碼",
        "APP無法登入",
        "app無法登入",
    ],
    "優惠券問題類型": [
        "優惠券無法使用",
        "使用規則/限制條件說明",
        "查詢優惠券序號紀錄",
    ],
    "回收點數問題類型": [
        "點數重複入點",
        "點數未入帳號",
        "投入後未獲點數/點數未記錄",
    ],
    "機台問題類型": [
        # 合併後的細項（原本互相分不開的幾項已併為一項，見 DETAIL_ALIASES）
        "機台故障/無法運作",
        "螢幕/操作介面異常",
        "辨識/重量判定異常",
        "黑色分選門異常或卡瓶堵塞",
        "履帶未作動或異常抖動",
        "機台網路連線失敗",
        "機台髒污/需要清潔",
        "網路中斷或不穩定",
        "投入物卡住_瓶罐/電池",
        "投入後未獲點數/點數未記錄",
        "瓶蓋桶已滿",
        "寶特瓶滿艙",
        "回收箱滿艙",
        "回收艙門開啟",
    ],
    "顧客關係類型": [
        "許願新增站點/設站建議",
        "申請刪除帳號",
        "更換帳號",
        "其他建議",
        "規則與活動諮詢",
    ],
}

TYPE_OPTIONS = list(TOPIC_DETAIL_MAP.keys())
DETAIL_OPTIONS = [d for lst in TOPIC_DETAIL_MAP.values() for d in lst]

DEPT_OPTIONS = [
    "營運部", "研發部", "廠務部", "人資部", "行銷部",
    "資訊部", "企劃部", "財務部", "開發部", "總經理室"
]

DEPT_MAP = {
    "機台問題類型": "營運部",
    "機台相關問題": "營運部",
    "APP帳號設定問題類型": "資訊部",
    "APP使用問題類型": "資訊部",
    "回收點數問題類型": "",
    "優惠券問題類型": "行銷部",
    "顧客關係類型": "營運部",
}


# 各問題類型的預設細項（分類失敗時的保底值）
FALLBACK_TYPE = "顧客關係類型"
FALLBACK_DETAIL = "其他建議"


def default_detail_for(topic: str) -> str:
    """回傳該類型的第一個細項，作為細項不合法時的保底。"""
    return TOPIC_DETAIL_MAP.get(topic, [FALLBACK_DETAIL])[0]


def details_of(topic: str) -> list[str]:
    return list(TOPIC_DETAIL_MAP.get(topic, []))


# 由公司明訂政策決定的細項：必須依客訴內容區分，不受舊標記影響。
# 這些細項的判斷以內建規則為準，不會被「從舊資料挖出的規則」覆蓋，
# 因為舊資料在政策訂立前把三種滿艙混用同一個標記。
POLICY_DETAILS = {
    "瓶蓋桶已滿",
    "寶特瓶滿艙",
    "回收箱滿艙",
}
POLICY_CONFIDENCE = 0.85


# 已廢除／更名的舊標記 → 現行標記
# 讓歷史資料不會因為分類法調整而失效（知識庫仍能學到那些筆）
RETIRED_TOPICS = {
    "APP帳密登入問題": "APP帳號設定問題類型",   # 細項與帳號設定完全重疊，已廢除
    "機台相關問題": "機台問題類型",             # 舊類型名
}

DETAIL_ALIASES: dict[str, str] = {
    # 歷史寫法 → 現行細項（例："瓶蓋箱已滿" 與 "瓶蓋桶已滿" 為同一項）
    "瓶蓋箱已滿": "瓶蓋桶已滿",
    "瓶蓋桶滿艙": "瓶蓋桶已滿",
    "寶特瓶已滿": "寶特瓶滿艙",
    "回收箱已滿": "回收箱滿艙",
    # ── 2026/08 細項合併（方案 C）──
    # 這些細項在實測中彼此分不開（雙向混淆），合併後細項正確率
    # 63.0% → 68.9%、完全自動採用 55.4% → 59.8%。
    # 舊名稱保留為別名，歷史資料與內建規則吐出的舊名稱都會自動歸位。
    "機台操作畫面無法登入": "螢幕/操作介面異常",
    "螢幕異常顯示/畫面異常": "螢幕/操作介面異常",
    "螢幕西曬導致黑屏或反光": "螢幕/操作介面異常",
    "辨識失敗異常或錯誤": "辨識/重量判定異常",
    "重量偵測異常": "辨識/重量判定異常",
    "投口綠燈拒收容器": "辨識/重量判定異常",
    "機台運作中斷/重啟": "機台故障/無法運作",
    "機台當機/無回應": "機台故障/無法運作",
    "機台關閉/無法啟動": "機台故障/無法運作",
    "機台需維護/故障提醒": "機台故障/無法運作",
    "操作流程異常/無法正常操作": "機台故障/無法運作",
    "兌換失敗/顯示錯誤": "優惠券無法使用",
    "無法進行兌換操作": "優惠券無法使用",
    "回收物使用規則": "規則與活動諮詢",
    "相關活動規則疑問": "規則與活動諮詢",
}


# 細項 → 所屬類型（同名細項可能出現在多個類型，取第一個）
DETAIL_TO_TOPIC: dict[str, str] = {}
for _topic, _details in TOPIC_DETAIL_MAP.items():
    for _d in _details:
        DETAIL_TO_TOPIC.setdefault(_d, _topic)


SEPARATOR_CHARS = set(" 	　/|、,，-—－_或") | {chr(92), chr(65295)}


def _loose_key(name: str) -> str:
    """去掉分隔符與空白後的比對鍵。

    歷史資料的細項常只差分隔符寫法（「機台需維護-故障提醒」對「機台需維護/故障提醒」、
    「履帶未作動 or 異常抖動」對「履帶未作動或異常抖動」），
    這些應視為同一個細項，不該當成無法對應而被丟棄。
    """
    text = str(name).strip().lower()
    parts = [t for t in text.split() if t != "or"]   # 拿掉獨立的 or
    return "".join(ch for ch in "".join(parts) if ch not in SEPARATOR_CHARS)


# 寬鬆比對表；同一個鍵對應到多個細項時視為不明確，不納入
_LOOSE_DETAIL_MAP: dict[str, str] = {}
_loose_conflicts: set[str] = set()
for _d in DETAIL_TO_TOPIC:
    _k = _loose_key(_d)
    if not _k:
        continue
    if _k in _LOOSE_DETAIL_MAP:
        # 同一個鍵對到不同「類型」才算真衝突；APP／app 這種大小寫變體不算
        if DETAIL_TO_TOPIC[_LOOSE_DETAIL_MAP[_k]] != DETAIL_TO_TOPIC[_d]:
            _loose_conflicts.add(_k)
    else:
        _LOOSE_DETAIL_MAP[_k] = _d
for _k in _loose_conflicts:
    _LOOSE_DETAIL_MAP.pop(_k, None)


# 大小寫變體正規化：分類法裡同時有 "APP無法登入" 與 "app無法登入" 這種寫法，
# 而 normalize_problem_labels() 會把細項的英文轉小寫，
# 若不統一，同一個細項會被當成兩種標記（知識庫學不到一起、評估也會誤判成錯）。
CASE_CANONICAL: dict[str, str] = {}
for _topic, _details in TOPIC_DETAIL_MAP.items():
    _by_lower: dict[str, list[str]] = {}
    for _d in _details:
        _by_lower.setdefault(_d.lower(), []).append(_d)
    for _lower, _variants in _by_lower.items():
        if len(_variants) > 1:
            # 以小寫版本為準（與 normalize_problem_labels 的輸出一致）
            _canon = next((v for v in _variants if v == _lower), _variants[-1])
            for _v in _variants:
                if _v != _canon:
                    CASE_CANONICAL[_v] = _canon


def legalize_pair(topic: str, detail: str) -> tuple[str, str] | None:
    """把 (類型, 細項) 修正成合法組合；完全無法對應時回傳 None。

    用於學習歷史標記：寧可略過不合分類法的舊資料，也不要用猜測的細項污染知識庫。
    """
    topic = (topic or "").strip()
    detail = (detail or "").strip()
    topic = RETIRED_TOPICS.get(topic, topic)
    detail = DETAIL_ALIASES.get(detail, detail)
    detail = CASE_CANONICAL.get(detail, detail)
    if topic in TOPIC_DETAIL_MAP and detail in TOPIC_DETAIL_MAP[topic]:
        return topic, detail
    if detail in DETAIL_TO_TOPIC:          # 細項合法但掛錯類型 → 歸回細項真正的類型
        return DETAIL_TO_TOPIC[detail], detail
    loose = _LOOSE_DETAIL_MAP.get(_loose_key(detail))   # 只差分隔符寫法
    # 合併模擬會暫時把細項移出 DETAIL_TO_TOPIC，寬鬆比對表可能仍指向舊名稱，
    # 這裡一律用 get()，對不到就當成不合法，不要讓評估整個中斷。
    if loose and loose in DETAIL_TO_TOPIC:
        return DETAIL_TO_TOPIC[loose], loose
    return None


def coerce_pair(topic: str, detail: str) -> tuple[str, str]:
    """把任意輸出夾回合法組合；無法對應時退回該類型的預設細項或保底分類。"""
    fixed = legalize_pair(topic, detail)
    if fixed is not None:
        return fixed
    topic = (topic or "").strip()
    if topic in TOPIC_DETAIL_MAP:
        return topic, default_detail_for(topic)
    return FALLBACK_TYPE, FALLBACK_DETAIL
