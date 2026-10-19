"""L1 分類規則：內建人工關鍵字規則（最高優先的覆寫層）。

analyze_complaint 為原本 complaint_webapp.py 內的規則鏈，原樣搬移。
Phase 2 起，另有從歷史紀錄自動挖掘的規則（automation/knowledge.py），
兩者的優先順序由 automation/classifier.py 決定：人工規則永遠優先。
"""

import re

from .taxonomy import TOPIC_DETAIL_MAP


# 滿艙相關用詞（「滿」的說法很多，集中在這裡維護）
FULL_KEYWORDS = [
    # 直接說滿
    "滿倉", "滿艙", "滿倉庫", "收滿", "滿台", "已滿", "滿了", "爆滿", "全滿", "都滿",
    "滿出來", "太滿", "很滿", "塞滿", "滿到", "滿滿", "滿袋", "滿瓶", "滿載", "滿機",
    "額滿", "容量滿", "裝滿", "是滿的", "常滿", "老是滿", "一直滿",
    # 因為滿而無法投遞的說法
    "可投數量都是零", "可投數量是0", "可投數量0", "顯示滿",
    # 客訴實際出現過的其他說法（含簡體與誤字）
    "客滿", "滿團", "滿团", "满了", "已满", "沒空位", "沒有空位",
]
# 含「滿」但與「裝滿」無關的詞：單字「滿」要觸發滿艙判定之前先把這些剔除，
# 否則「我很不滿意」「服務滿分」都會被判成機台滿艙。
NOT_FULL_WORDS = [
    "滿意", "不滿", "滿足", "圓滿", "美滿", "滿分", "充滿",
    "滿心", "滿頭", "滿臉", "滿腹", "滿街", "滿天",
]

# 需要清運／換袋的說法（多半就是滿了）
FULL_CONTEXT_KEYWORDS = [
    "需清運", "來清運", "請清運", "沒人清", "沒人收", "沒有清運", "來收", "清機",
    "換袋", "清一下", "幫忙清", "派人清", "收空瓶", "清理一下",
    "加強清運", "什麼時候會清運", "清運", "去清理", "都沒清", "沒清",
]


CAP_KEYWORDS = ["瓶蓋", "蓋子", "瓶蓋桶", "瓶蓋箱", "瓶蓋回收"]

# 重量相關：出現任一個就一律歸「重量偵測異常」
WEIGHT_KEYWORDS = ["重量", "秤重", "偵測重量", "稱重", "磅重", "重量感測", "克數", "公克數"]

# 寶特瓶滿艙：寶特瓶與其他塑膠類容器（依實際客訴用語整理）
PET_KEYWORDS = [
    "寶特瓶", "保特瓶", "宝特瓶", "寶特", "保特",
    "塑膠瓶", "塑胶瓶", "塑膠", "塑胶", "塑料",
    "牛奶瓶", "鮮奶瓶", "鲜奶瓶", "優酪乳", "优酪乳", "養樂多", "养乐多",
    "塑膠杯", "塑胶杯", "飲料杯", "饮料杯",
    "ｐｅｔ", "ｐｖｃ",
]
# 英文簡寫需用邊界比對，否則 "pp" 會命中 "app"、"pet" 會命中 "carpet"
PET_LATIN_RE = re.compile(r"(?<![a-z])(pet|pvc|pp)(?![a-z])")
# 「瓶滿」「瓶子滿」「瓶罐都滿」這類：瓶與滿必須相鄰（中間最多 3 個字），
# 否則整句話只要出現「瓶」就會把「回收箱滿艙」的客訴搶走。
BOTTLE_FULL_RE = re.compile(r"瓶[^，。！？!?\s]{0,3}[滿满]")


def full_bin_detail(text: str) -> str:
    """已知是滿艙客訴時，依內容提到的回收物決定是哪一種滿艙。

    順序＝由具體到概括：有提到瓶蓋→瓶蓋桶；有提到寶特瓶／塑膠瓶→寶特瓶滿艙；
    都沒提到具體回收物→回收箱滿艙。
    """
    t = str(text).lower()
    # 瓶蓋一定要先判：「瓶蓋」本身含「瓶」，順序顛倒的話會被當成寶特瓶。
    if any(k in t for k in CAP_KEYWORDS):
        return "瓶蓋桶已滿"
    if any(k in t for k in PET_KEYWORDS) or PET_LATIN_RE.search(t):
        return "寶特瓶滿艙"
    # 「瓶滿／瓶子滿／瓶罐滿」→ 寶特瓶（機台收的瓶就是寶特瓶）。
    # 必須是「瓶」緊接著「滿」才算，不能整句出現「瓶」就算 ——
    # 實測後者會把 95 筆人工標為「回收箱滿艙」的客訴搶走
    #（例：「回收箱滿了，瓶子投不進去」重點是箱子滿，不是瓶子）。
    # 這一步放在瓶蓋之後，所以「瓶蓋滿」不會誤落到這裡。
    if BOTTLE_FULL_RE.search(t):
        return "寶特瓶滿艙"
    return "回收箱滿艙"


def _is_full_complaint(text: str) -> bool:
    """判斷是否在反映「裝滿了、需要清運」。

    「滿」的說法很多（滿倉／滿袋／額滿／滿載／可投數量歸零），
    這裡集中維護，避免散落在規則鏈各處。
    """
    if any(k in text for k in FULL_KEYWORDS):
        return True
    # 只講「請來清運／清機」而沒講滿，也算（機台需要清運就是因為滿）
    if any(k in text for k in FULL_CONTEXT_KEYWORDS):
        return True
    # 只要出現「滿」字就算 —— 詞表再長都列不完（實測有 16.4% 的滿艙客訴
    # 因為用了表上沒有的說法而漏判）。先剔除「滿意／不滿」這類無關的詞，
    # 剩下還有「滿」才算。
    stripped = text
    for w in NOT_FULL_WORDS:
        stripped = stripped.replace(w, "")
    return "滿" in stripped or "满" in stripped


def analyze_complaint(subject: str, content: str) -> tuple[str, str]:
    s = subject if isinstance(subject, str) else ""
    c = content if isinstance(content, str) else ""
    t = (s + " " + c).lower()

    # 顧客關係類型
    if "註冊" in t and "無法" in t:
        return "顧客關係類型", "其他建議"
    if any(k in t for k in ["不處理", "態度", "搞什麼", "不願意"]):
        return "顧客關係類型", "其他建議"
    if any(k in t for k in ["刪除帳號", "註銷"]):
        return "顧客關係類型", "申請刪除帳號"
    if any(k in t for k in ["手機號碼", "原帳號"]) and any(k in t for k in ["變更", "更改", "修改"]):
        return "顧客關係類型", "更換帳號"
    if any(k in t for k in ["更換帳號", "換帳號"]):
        return "顧客關係類型", "更換帳號"
    if any(k in t for k in ["新增站點", "設站", "建議", "許願"]):
        return "顧客關係類型", "許願新增站點/設站建議"
    if any(k in t for k in ["回收規則", "材質", "可回收"]):
        return "顧客關係類型", "回收物使用規則"
    if any(k in t for k in ["活動規則", "活動疑問", "相關活動"]):
        return "顧客關係類型", "相關活動規則疑問"

    # APP帳號設定 / 登入
    if any(k in t for k in ["驗證碼", "認證碼", "otp", "簡訊"]):
        if "忘記密碼" in t:
            return "APP帳號設定問題類型", "忘記密碼/無法重設密碼"
        return "APP帳號設定問題類型", "無法接收簡訊驗證碼"
    if any(k in t for k in ["修改", "更改", "更換"]) and any(k in t for k in ["帳號", "手機", "電話", "號碼"]):
        return "APP帳號設定問題類型", "帳號資訊修改/設定"

    # 登入問題（獨立分類）
    if any(k in t for k in ["登入", "登不進去"]) and any(k in t for k in ["螢幕", "機台", "黑掉"]) and any(k in t for k in ["無法", "不能", "失敗", "不了"]):
        return "機台問題類型", "機台操作畫面無法登入"
    if any(k in t for k in ["無法登入", "不能登入", "登不進去", "登入失敗", "登入不了"]):
        return "APP帳號設定問題類型", "APP無法登入"

    # APP使用
    if "可投數量" in t or ("app" in t and "顯示" in t and "0" not in t):
        return "APP使用問題類型", "APP畫面顯示與機台狀態不符"
    if "顯示" in t and "不符" in t:
        return "APP使用問題類型", "APP畫面顯示與機台狀態不符"
    if any(k in t for k in ["app異常", "閃退", "轉圈", "更新"]):
        return "APP使用問題類型", "APP多重異常狀況"
    if any(k in t for k in ["商家頁面", "頁面空白"]):
        return "APP使用問題類型", "APP商家頁面空白"
    if any(k in t for k in ["點數顯示異常", "點數顯示錯誤"]):
        return "APP使用問題類型", "APP點數顯示異常"

    # ── 重量：只要提到就一律歸「重量偵測異常」──────────────
    # 必須排在辨識／顯示／點數那幾條之前。實測「無法辨識重量」會被
    # 辨識規則攔成「辨識/重量判定異常」、「投很多次重量都不對」會被
    # 點數規則攔成「點數未入帳號」、「一直顯示重量錯誤」會被螢幕規則攔走。
    # 滿艙排在更前面（公司政策細項），兩者同時出現時仍以滿艙為準。
    # 滿艙是公司政策細項，排在更後面但優先權更高，
    # 所以這裡要先排除掉滿艙客訴（例：「寶特瓶滿了重量也不對」仍歸寶特瓶滿艙）。
    if any(k in t for k in WEIGHT_KEYWORDS) and not _is_full_complaint(t):
        return "機台問題類型", "重量偵測異常"

    # 回收點數
    if "點數" in t and any(k in t for k in ["未累積", "未增加", "沒有入帳", "未入帳"]):
        return "回收點數問題類型", "點數未入帳號"
    if ("點數" in t or "沒入點" in t or "計點" in t) and any(k in t for k in ["未入", "沒入", "不見", "沒記", "沒收到"]):
        return "回收點數問題類型", "點數未入帳號"
    if "點數" in t and any(k in t for k in ["重複", "多給", "多入"]):
        return "回收點數問題類型", "點數重複入點"
    if any(k in t for k in ["投入後沒點", "未獲點數", "未記錄"]):
        return "回收點數問題類型", "投入後未獲點數/點數未記錄"

    # 優惠券
    if any(k in t for k in ["優惠券", "兌換券", "折價", "序號", "抵用", "對換券", "票卷", "票夾", "條碼", "換這個"]):
        if any(k in t for k in ["提前按下", "操作錯誤", "系統還沒更新", "已更換", "限制", "期限", "規則"]):
            return "優惠券問題類型", "使用規則/限制條件說明"
        if any(k in t for k in ["過期", "還原", "點到", "沒按出條碼"]):
            return "優惠券問題類型", "無法進行兌換操作"
        if any(k in t for k in ["已使用", "失敗", "錯誤", "不能用", "刷不過", "沒有跑出條碼", "這怎麼一回事"]):
            return "優惠券問題類型", "兌換失敗/顯示錯誤"
        if any(k in t for k in ["查詢", "紀錄", "找不到", "在哪"]):
            return "優惠券問題類型", "查詢優惠券序號紀錄"
        return "優惠券問題類型", "無法進行兌換操作"

    # 機台問題（核心狀態）
    # ── 滿艙：依客訴內容提到的回收物區分三種（順序＝由具體到概括）──
    if _is_full_complaint(t):
        return "機台問題類型", full_bin_detail(t)
    if any(k in t for k in ["處理中", "卡住"]) and "暫停不動" in t:
        return "機台問題類型", "機台當機/無回應"
    if "寶特瓶卡住" in t or "卡在" in t or ("黑色門" in t and "卡住" in t) or "卡瓶" in t:
        return "機台問題類型", "投入物卡住_瓶罐/電池"
    if any(k in t for k in ["投很多次", "無法辨識", "一直顯示", "不顯示綠燈", "辨識失敗", "辨識異常"]):
        return "機台問題類型", "辨識失敗異常或錯誤"
    if "綠燈" in t and any(k in t for k in ["不能", "拒收", "不收"]):
        return "機台問題類型", "投口綠燈拒收容器"
    if any(k in t for k in ["顯示0都沒有更新", "通報維修", "維護", "維修", "需維修", "故障提醒"]):
        return "機台問題類型", "機台需維護/故障提醒"
    if any(k in t for k in ["關閉", "設備不動", "不能使用", "撤機", "故障快", "沒開", "未開啟", "關機"]):
        return "機台問題類型", "機台關閉/無法啟動"
    if any(k in t for k in ["髒污不收", "清潔", "髒污"]):
        return "機台問題類型", "機台髒污/需要清潔"
    if any(k in t for k in ["當機", "故障訊息", "沒反應", "lag", "機台異常"]):
        return "機台問題類型", "機台當機/無回應"
    if "運轉不會停止" in t:
        return "機台問題類型", "操作流程異常/無法正常操作"
    if any(k in t for k in ["黑色分選門", "分選門異常", "卡瓶堵塞"]):
        return "機台問題類型", "黑色分選門異常或卡瓶堵塞"

    # 機台問題（設備/環境）
    if any(k in t for k in ["履帶", "輸送帶", "傳送帶"]) and any(k in t for k in ["不動", "不轉", "異常"]):
        return "機台問題類型", "履帶未作動或異常抖動"
    if "西曬" in t or ("反光" in t and "螢幕" in t):
        return "機台問題類型", "螢幕西曬導致黑屏或反光"
    if any(k in t for k in ["黑屏", "黑畫面", "螢幕異常", "畫面異常", "黑掉"]):
        return "機台問題類型", "螢幕異常顯示/畫面異常"
    if any(k in t for k in ["網路連線失敗", "連不上", "連線失敗"]) and "機台" in t:
        return "機台問題類型", "機台網路連線失敗"
    if any(k in t for k in ["網路不穩", "網路中斷"]):
        return "機台問題類型", "網路中斷或不穩定"
    if any(k in t for k in ["無法操作", "流程異常", "不能操作"]):
        return "機台問題類型", "操作流程異常/無法正常操作"
    if any(k in t for k in ["中斷", "重啟", "重開機"]):
        return "機台問題類型", "機台運作中斷/重啟"
    if any(k in t for k in ["艙門", "門沒關", "回收艙門"]):
        return "機台問題類型", "回收艙門開啟"

    return "顧客關係類型", "其他建議"


# ---- valid type set for fast lookup (all keys + known variant spellings from template) ----
_VALID_TYPES = set(TOPIC_DETAIL_MAP.keys())

# All valid details (flattened from TOPIC_DETAIL_MAP) for quick check
_VALID_DETAILS_FLAT: set[str] = {d for lst in TOPIC_DETAIL_MAP.values() for d in lst}


def _is_valid_pair(t: str, d: str) -> bool:
    """Return True if both type and detail are non-empty and the detail belongs to the type."""
    t, d = t.strip(), d.strip()
    if not t or not d:
        return False
    # Accept if type is valid AND detail is in that type's list
    if t in TOPIC_DETAIL_MAP and d in TOPIC_DETAIL_MAP[t]:
        return True
    # Also accept if type exists but detail is in the FULL detail pool (legacy data)
    if t in _VALID_TYPES and d in _VALID_DETAILS_FLAT:
        return True
    return False
