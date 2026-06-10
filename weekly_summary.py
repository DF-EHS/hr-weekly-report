#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
weekly_summary.py - 週工作彙整報告產生器

彙整週期：每週一（含）到週日（含），共 7 天。
功能：
  1. 自動計算當前彙整週期
  2. 讀取 Outlook 本地資料夾，取得週期內每位成員的所有日誌信件
  3. 將同成員的日誌合併後呼叫 Gemini API，提煉：
     ・本週核心工作總結（3 個 Bullet points）
     ・週平均達成率（0-100 整數）
     ・本週綜合狀態燈號（RAG）
  4. 套用 weekly_template.html 範本，輸出 weekly_dashboard.html

使用方式：
  python weekly_summary.py          # 以今天日期推算當前週期
  python weekly_summary.py 2026-05-21   # 指定週期內的任意一天
"""

import json
import re
import sys
import io
from datetime import datetime, timedelta, date
from pathlib import Path

# 排程執行時強制 stdout/stderr 使用 UTF-8
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

# ─────────────────────────────────────────────
# 套件載入檢查
# ─────────────────────────────────────────────
try:
    import win32com.client
except ImportError:
    print("❌ 缺少套件：請先執行  pip install pywin32")
    sys.exit(1)

try:
    import requests
except ImportError:
    print("❌ 缺少套件：請先執行  pip install requests")
    sys.exit(1)

# ─────────────────────────────────────────────
# 路徑 / 常數設定
# ─────────────────────────────────────────────
SCRIPT_DIR    = Path(__file__).parent
CONFIG_PATH   = SCRIPT_DIR / "config.json"
TEMPLATE_PATH = SCRIPT_DIR / "weekly_template.html"
OUTPUT_PATH   = SCRIPT_DIR / "weekly_dashboard.html"

GEMINI_MODEL = "gemini-2.5-flash"
GEMINI_URL   = (
    "https://generativelanguage.googleapis.com/v1beta/models/"
    f"{GEMINI_MODEL}:generateContent"
)

# 中文星期對照（weekday() 回傳 0=Mon）
_WEEKDAY_ZH = ["一", "二", "三", "四", "五", "六", "日"]

# 成員標籤顏色
_MEMBER_COLORS: dict[str, str] = {
    "bella":   "#7b5ea7",
    "an":      "#2a7dd9",
    "gloom":   "#e07b39",
    "陳信佳":  "#1a8c6e",
    "賴柏羽":  "#c0392b",
    "chiehyi": "#6c757d",
    "phebe":   "#d6336c",
}

# RAG 顏色與文字
_RAG_META: dict[str, tuple[str, str]] = {
    "green":  ("#28a745", "🟢 達成"),
    "yellow": ("#e6a800", "🟡 部分達成"),
    "red":    ("#dc3545", "🔴 待改善"),
}


# ─────────────────────────────────────────────
# 設定檔
# ─────────────────────────────────────────────
def load_config() -> dict:
    """載入 config.json"""
    if not CONFIG_PATH.exists():
        print(f"❌ 找不到設定檔：{CONFIG_PATH}")
        sys.exit(1)
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


# ─────────────────────────────────────────────
# 週期計算
# ─────────────────────────────────────────────
def get_week_range(ref: date | None = None) -> tuple[date, date]:
    """
    計算「週一 → 週日」週期（共 7 天）。
    weekday(): Mon=0, Tue=1, Wed=2, Thu=3, Fri=4, Sat=5, Sun=6

    ・手動指定 ref：回傳 ref 所在的週一~週日週期。
      範例（ref = 2026-06-03 週三）→ week_start=2026-06-01、week_end=2026-06-07。
    ・ref 為 None（自動排程）：回傳「最近一個已結束」的週一~週日週期，
      以上一個已過完的週日作為週期結尾（例：於 6/10 週三執行 → 06/01~06/07）。
    """
    if ref is None:
        ref = date.today()
        # 自動排程：取最近一個已結束的週日作為週期結尾
        last_sunday = ref - timedelta(days=ref.weekday() + 1)
        return last_sunday - timedelta(days=6), last_sunday
    # 手動指定日期：回傳該日所在的週一~週日
    week_start = ref - timedelta(days=ref.weekday())
    return week_start, week_start + timedelta(days=6)


def format_date_zh(d: date) -> str:
    """格式化日期為 YYYY/MM/DD（週N）"""
    return f"{d.strftime('%Y/%m/%d')}（{_WEEKDAY_ZH[d.weekday()]}）"


# ─────────────────────────────────────────────
# 發信人辨識
# ─────────────────────────────────────────────
def identify_sender(sender_name: str, sender_email: str, mapping: dict) -> str:
    """根據 mapping 關鍵字辨識成員代稱（不分大小寫）"""
    combined = f"{sender_name} {sender_email}".lower()
    for keyword, member in mapping.items():
        if keyword.lower() in combined:
            return member
    return sender_name.strip() or "未知"


# ─────────────────────────────────────────────
# Outlook 信件讀取
# ─────────────────────────────────────────────
def _find_folder(parent, name: str):
    """在 parent 的子資料夾中尋找指定名稱（不分大小寫）"""
    for folder in parent.Folders:
        if folder.Name.strip().lower() == name.strip().lower():
            return folder
    return None


def get_week_emails(
    subjects: list[str],
    week_start: date,
    week_end: date,
    folder_path: str = "",
) -> list[dict]:
    """
    讀取 Outlook，回傳 week_start ~ week_end（含）期間內，
    主旨含指定關鍵字的信件清單。
    """
    try:
        outlook   = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        base      = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
    except Exception as exc:
        print(f"❌ 無法連接 Outlook：{exc}")
        sys.exit(1)

    # 導航至目標子資料夾
    target = base
    if folder_path:
        parts = [p.strip() for p in folder_path.split("/") if p.strip()]
        for part in parts:
            found = _find_folder(target, part)
            if found is None:
                print(f"⚠️  找不到子資料夾「{part}」，改用收件匣")
                target = base
                break
            target = found
        print(f"   📂 讀取資料夾：{target.Name}")

    emails: list[dict] = []
    for item in target.Items:
        try:
            if item.Class != 43:   # 43 = olMail
                continue

            received = item.ReceivedTime
            if not hasattr(received, "replace"):
                continue
            received_dt   = received.replace(tzinfo=None)
            received_date = received_dt.date()

            # 篩選週期內的信件
            if not (week_start <= received_date <= week_end):
                continue

            subject = item.Subject or ""
            if not any(kw in subject for kw in subjects):
                continue

            emails.append({
                "subject":      subject,
                "sender_name":  item.SenderName or "",
                "sender_email": item.SenderEmailAddress or "",
                "body":         item.Body or "",
                "date":         received_dt.strftime("%Y-%m-%d"),
            })
        except Exception:
            continue

    return emails


# ─────────────────────────────────────────────
# Gemini API 週彙整分析
# ─────────────────────────────────────────────
def call_gemini_weekly(
    api_key: str,
    member: str,
    combined_logs: str,
    n_days: int,
) -> dict:
    """
    呼叫 Gemini，針對成員本週所有日誌進行彙整分析。
    回傳：{
        bullets: [str, str, str],
        avg_completion_rate: int,
        rag_status: str
    }
    """
    prompt = (
        f"以下是「{member}」本週共 {n_days} 天的工作日誌，"
        "請進行週彙整分析，並以 JSON 格式回傳三個欄位：\n"
        "1. bullets：本週核心工作總結，陣列包含「恰好 3 個」繁體中文條列要點（每點 40 字以內）\n"
        "2. avg_completion_rate：週平均達成率，0 到 100 之間的整數"
        "（綜合評估這幾天的工作完成程度）\n"
        "3. rag_status：依平均達成率判定燈號"
        "（green = 80 以上、yellow = 60-79、red = 59 以下）\n\n"
        "只回傳 JSON 物件，不要加任何說明文字或 markdown 標記。\n"
        '範例格式：{"bullets":["完成...","推進...","協助..."],'
        '"avg_completion_rate":85,"rag_status":"green"}\n\n'
        f"---\n{combined_logs[:8000]}"
    )

    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {
            "temperature": 0.2,
            "maxOutputTokens": 1024,
            "thinkingConfig": {"thinkingBudget": 0},   # 關閉思考模式，確保回傳純 JSON
        },
    }

    try:
        resp = requests.post(
            f"{GEMINI_URL}?key={api_key}",
            json=payload,
            timeout=60,
        )
        resp.raise_for_status()
        data = resp.json()

        # 收集非思考模式的 parts
        parts = data["candidates"][0]["content"]["parts"]
        text  = ""
        for part in parts:
            if not part.get("thought", False):
                text += part.get("text", "")
        text = text.strip()

        # 移除 markdown code block 包裹
        text = re.sub(r"^```(?:json)?\s*", "", text)
        text = re.sub(r"\s*```$", "", text).strip()

        # 從文字中抓取第一個完整 JSON 物件
        json_match = re.search(r"\{[\s\S]*\}", text)
        if json_match:
            text = json_match.group(0)

        result  = json.loads(text)
        bullets = result.get("bullets", [])

        # 確保 bullets 為 list，且恰好 3 個
        if not isinstance(bullets, list):
            bullets = [str(bullets), "—", "—"]
        while len(bullets) < 3:
            bullets.append("（無資料）")
        bullets = [str(b) for b in bullets[:3]]

        return {
            "bullets":             bullets,
            "avg_completion_rate": int(result.get("avg_completion_rate", 0)),
            "rag_status":          str(result.get("rag_status", "red")).lower(),
        }

    except Exception as exc:
        print(f"    ⚠️  Gemini API 呼叫失敗：{exc}")
        return {
            "bullets":             ["（API 錯誤，無法分析）", "—", "—"],
            "avg_completion_rate": 0,
            "rag_status":          "red",
        }


# ─────────────────────────────────────────────
# HTML 組裝
# ─────────────────────────────────────────────
def _member_color(member: str) -> str:
    """取得成員標籤顏色"""
    return _MEMBER_COLORS.get(member.lower(), "#607d8b")


def build_member_card(member: str, result: dict, n_logs: int) -> str:
    """產生有資料的成員週報卡片 HTML"""
    rate = result["avg_completion_rate"]
    rag  = result["rag_status"]
    rag_color, rag_label = _RAG_META.get(rag, ("#6c757d", "⚪ 未知"))
    mc        = _member_color(member)
    bar_color = "#28a745" if rate >= 80 else "#e6a800" if rate >= 60 else "#dc3545"

    bullets_html = "\n".join(
        f"        <li>{b}</li>" for b in result["bullets"]
    )

    return (
        f'      <div class="member-card rag-{rag}">\n'
        f'        <div class="card-header">\n'
        f'          <span class="member-tag" style="background:{mc}">{member}</span>\n'
        f'          <span class="rag-pill" style="background:{rag_color}">{rag_label}</span>\n'
        f'        </div>\n'
        f'        <div class="rate-row">\n'
        f'          <div class="rate-num" style="color:{bar_color}">{rate}%</div>\n'
        f'          <div class="rate-meta">\n'
        f'            <div class="rate-label">週平均達成率</div>\n'
        f'            <div class="rate-days">{n_logs} 天日誌</div>\n'
        f'          </div>\n'
        f'        </div>\n'
        f'        <div class="progress-track">\n'
        f'          <div class="progress-fill" style="width:{rate}%;background:{bar_color}"></div>\n'
        f'        </div>\n'
        f'        <div class="divider"></div>\n'
        f'        <div class="bullets-label">📌 本週核心工作</div>\n'
        f'        <ul class="bullet-list">\n'
        f'{bullets_html}\n'
        f'        </ul>\n'
        f'      </div>'
    )


def build_no_data_card(member: str) -> str:
    """產生無資料的成員週報卡片 HTML"""
    mc = _member_color(member)
    return (
        f'      <div class="member-card rag-none">\n'
        f'        <div class="card-header">\n'
        f'          <span class="member-tag" style="background:{mc}">{member}</span>\n'
        f'          <span class="rag-pill" style="background:#adb5bd">⚪ 無資料</span>\n'
        f'        </div>\n'
        f'        <div class="no-data-msg">本週無日誌記錄</div>\n'
        f'      </div>'
    )


# ─────────────────────────────────────────────
# 主程式
# ─────────────────────────────────────────────
def main() -> None:
    print("=" * 52)
    print("  週工作彙整報告產生器")
    print("=" * 52)
    print()

    # ── 1. 解析命令列參數（可指定週期內任一天）──
    ref_date: date | None = None
    if len(sys.argv) >= 2:
        try:
            ref_date = date.fromisoformat(sys.argv[1])
            print(f"📌 使用指定參考日期：{sys.argv[1]}")
        except ValueError:
            print(f"⚠️  日期格式錯誤（應為 YYYY-MM-DD），改用今天")

    # ── 2. 載入設定 ──
    config  = load_config()
    api_key = config.get("gemini_api_key", "")
    if not api_key or api_key == "YOUR_GEMINI_API_KEY":
        print("❌ 請先在 config.json 中填入有效的 gemini_api_key")
        sys.exit(1)

    members: list[str] = config.get("members", [])

    # ── 3. 計算週期 ──
    week_start, week_end = get_week_range(ref_date)
    week_label = f"{format_date_zh(week_start)} ~ {format_date_zh(week_end)}"
    print(f"📅 彙整週期：{week_label}")
    print()

    # ── 4. 讀取 Outlook 信件 ──
    folder_path: str = config.get("outlook_folder_path", "")
    print("📬 讀取 Outlook（本週信件）...")
    emails = get_week_emails(
        config["email_subjects"], week_start, week_end, folder_path
    )
    print(f"   找到 {len(emails)} 封本週信件\n")

    # ── 5. 辨識發信人並分組 ──
    member_logs: dict[str, list[str]] = {m: [] for m in members}
    for email in emails:
        member = identify_sender(
            email["sender_name"], email["sender_email"], config["sender_mapping"]
        )
        if member not in members:
            continue
        log_text = f"=== {email['date']} ===\n{email['body'].strip()}"
        member_logs[member].append(log_text)

    # ── 6. 對每位成員呼叫 Gemini 彙整 ──
    results: dict[str, tuple[dict, int] | None] = {}
    for member in members:
        logs = member_logs.get(member, [])
        n    = len(logs)
        print(f"  📝 [{member}] 共 {n} 天日誌", end="")

        if n == 0:
            print(" → 略過（無資料）")
            results[member] = None
            continue

        print(" → 呼叫 Gemini 彙整...")
        combined = "\n\n".join(logs)
        result   = call_gemini_weekly(api_key, member, combined, n)
        results[member] = (result, n)

        rate = result["avg_completion_rate"]
        rag  = result["rag_status"]
        print(f"    ✅ 達成率 {rate}%，狀態：{rag}")

    print()

    # ── 7. 計算團隊統計 ──
    valid: list[tuple[dict, int]] = [
        val for val in results.values() if val is not None
    ]
    team_avg   = (
        round(sum(r["avg_completion_rate"] for r, _ in valid) / len(valid))
        if valid else 0
    )
    green_cnt  = sum(1 for r, _ in valid if r["rag_status"] == "green")
    yellow_cnt = sum(1 for r, _ in valid if r["rag_status"] == "yellow")
    red_cnt    = sum(1 for r, _ in valid if r["rag_status"] == "red")
    none_cnt   = len(members) - len(valid)

    team_avg_color = (
        "#28a745" if team_avg >= 80 else
        "#e6a800" if team_avg >= 60 else
        "#dc3545"
    )

    # ── 8. 組裝成員卡片 HTML ──
    cards_html: list[str] = []
    for member in members:
        val = results.get(member)
        if val is None:
            cards_html.append(build_no_data_card(member))
        else:
            r, n = val
            cards_html.append(build_member_card(member, r, n))

    # ── 9. 讀取範本並替換佔位符 ──
    if not TEMPLATE_PATH.exists():
        print(f"❌ 找不到範本：{TEMPLATE_PATH}")
        sys.exit(1)

    with open(TEMPLATE_PATH, "r", encoding="utf-8") as f:
        html = f.read()

    html = html.replace("{{WEEK_LABEL}}",    week_label)
    html = html.replace("{{GENERATED_AT}}",  datetime.now().strftime("%Y-%m-%d %H:%M"))
    html = html.replace("{{TEAM_AVG}}",      str(team_avg))
    html = html.replace("{{TEAM_AVG_COLOR}}", team_avg_color)
    html = html.replace("{{GREEN_COUNT}}",   str(green_cnt))
    html = html.replace("{{YELLOW_COUNT}}",  str(yellow_cnt))
    html = html.replace("{{RED_COUNT}}",     str(red_cnt))
    html = html.replace("{{NONE_COUNT}}",    str(none_cnt))
    html = html.replace("{{MEMBER_COUNT}}",  str(len(members)))
    html = html.replace("{{MEMBER_CARDS}}",  "\n".join(cards_html))

    # ── 10. 輸出 ──
    with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"✅ 完成！已產生：{OUTPUT_PATH}")
    print(
        f"   成員：{len(members)} 人 ｜"
        f" 有資料：{len(valid)} 人 ｜"
        f" 無資料：{none_cnt} 人 ｜"
        f" 團隊平均：{team_avg}%"
    )


if __name__ == "__main__":
    main()
