#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generate_report.py - HR 工作日誌合併報告產生器

整合「日誌明細」與「週彙整」兩個報表，輸出單一 report.html。

特點：
  - Outlook 只讀取一次，效率更高
  - 兩個 Tab 共用同一頁面，無需切換檔案
  - 週期定義：每週三 ～ 下週二

使用方式：
  python generate_report.py               # 以今天推算本週期
  python generate_report.py 2026-05-21    # 指定週期內任一天
"""

import json
import random
import re
import sys
import io
from datetime import datetime, timedelta, date
from pathlib import Path

# 排程執行時強制 stdout/stderr 使用 UTF-8，避免 cp950 編碼錯誤（循環 import 時跳過）
try:
    if hasattr(sys.stdout, 'buffer') and getattr(sys.stdout, 'encoding', '').lower() != 'utf-8':
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, 'buffer') and getattr(sys.stderr, 'encoding', '').lower() != 'utf-8':
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")
except (AttributeError, ValueError):
    pass

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
# 路徑 / 常數
# ─────────────────────────────────────────────
SCRIPT_DIR    = Path(__file__).parent
CONFIG_PATH   = SCRIPT_DIR / "config.json"
TEMPLATE_PATH = SCRIPT_DIR / "combined_template.html"
OUTPUT_PATH   = SCRIPT_DIR / "report.html"

GEMINI_MODEL = "gemini-2.5-flash"
GEMINI_URL   = (
    "https://generativelanguage.googleapis.com/v1beta/models/"
    f"{GEMINI_MODEL}:generateContent"
)

# ―――――――――――――――――――――――――――――――――――――――――――
# hr_tracker 延遲載入（避免循環 import，實際載入在 main() 內）
# ―――――――――――――――――――――――――――――――――――――――――――
_TRACKER_AVAILABLE = False  # 實際值在 main() 中設定

_WEEKDAY_ZH = ["一", "二", "三", "四", "五", "六", "日"]

_MEMBER_COLORS: dict[str, str] = {
    "bella":   "#7b5ea7",
    "an":      "#2a7dd9",
    "gloom":   "#e07b39",
    "陳信佳":  "#1a8c6e",
    "賴柏羽":  "#c0392b",
    "chiehyi": "#6c757d",
}

_RAG_META: dict[str, tuple[str, str]] = {
    "green":  ("#28a745", "🟢 達成"),
    "yellow": ("#e6a800", "🟡 達成"),   # 相容舊資料，顯示為達成
    "red":    ("#dc3545", "🔴 待改善"),
}


# ─────────────────────────────────────────────
# 設定檔
# ─────────────────────────────────────────────
def load_config() -> dict:
    if not CONFIG_PATH.exists():
        print(f"❌ 找不到設定檔：{CONFIG_PATH}")
        sys.exit(1)
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


# ─────────────────────────────────────────────
# 週期計算
# ─────────────────────────────────────────────
def get_week_range(ref: date | None = None) -> tuple[date, date]:
    """計算 ref 所在的「週三 → 下週二」週期。
    若 ref 為 None（即自動執行），且今天是週三（報告產出日），
    則以昨天（剛結束的週二）為基準，回傳剛完成的那一週。"""
    if ref is None:
        ref = date.today()
        # 週三是報告產出日，此時應產出剛結束的那週（前一週三~昨日週二）
        if ref.weekday() == 2:
            ref = ref - timedelta(days=1)
    days_since_wed = (ref.weekday() - 2) % 7
    week_start = ref - timedelta(days=days_since_wed)
    week_end   = week_start + timedelta(days=6)
    return week_start, week_end


def fmt_date_zh(d: date) -> str:
    return f"{d.strftime('%Y/%m/%d')}（{_WEEKDAY_ZH[d.weekday()]}）"


# ─────────────────────────────────────────────
# 發信人辨識
# ─────────────────────────────────────────────
def identify_sender(sender_name: str, sender_email: str, mapping: dict) -> str:
    combined = f"{sender_name} {sender_email}".lower()
    for keyword, member in mapping.items():
        if keyword.lower() in combined:
            return member
    return sender_name.strip() or "未知"


# ─────────────────────────────────────────────
# Outlook 信件讀取（只讀一次）
# ─────────────────────────────────────────────
def _find_folder(parent, name: str):
    for folder in parent.Folders:
        if folder.Name.strip().lower() == name.strip().lower():
            return folder
    return None


def get_emails(subjects: list[str], days_back: int,
               folder_path: str = "") -> list[dict]:
    """讀取 Outlook，回傳近 days_back 天、主旨符合的信件清單"""
    try:
        outlook   = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        base      = namespace.GetDefaultFolder(6)
    except Exception as exc:
        print(f"❌ 無法連接 Outlook：{exc}")
        sys.exit(1)

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

    cutoff = datetime.now() - timedelta(days=days_back)
    emails: list[dict] = []

    for item in target.Items:
        try:
            if item.Class != 43:
                continue
            received = item.ReceivedTime
            if not hasattr(received, "replace"):
                continue
            received_dt = received.replace(tzinfo=None)
            if received_dt < cutoff:
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
                "date_obj":     received_dt.date(),
            })
        except Exception:
            continue

    return emails


# ─────────────────────────────────────────────
# Gemini API - 單封日誌分析
# ─────────────────────────────────────────────
def call_gemini_single(api_key: str, body: str) -> dict:
    prompt = (
        "請分析以下工作日誌內容，並以 JSON 格式回傳三個欄位：\n"
        "1. summary：50 字以內的工作成果摘要（繁體中文）。"
        "摘要只聚焦實際完成的工作內容與成果，"
        "嚴禁提及任何職能評分或認證考核的具體分數（如「X/50 分」「總分 12/50」「總評 31/50 分」等）。"
        "（「觀望期」「初階使用者」等等級描述則可保留。）\n"
        "2. completion_rate：整體工作達成率，0 到 100 之間的整數\n"
        "3. rag_status：依達成率判定燈號（green = 80 以上、yellow = 60-79、red = 59 以下）\n\n"
        "只回傳 JSON 物件，不要加任何說明文字或 markdown 標記。\n\n"
        f"工作日誌內容：\n{body[:3000]}"
    )
    return _call_gemini(api_key, prompt, 512)


# ─────────────────────────────────────────────
# Gemini API - 週彙整（僅產生 bullets，達成率由日誌明細直接計算）
# ─────────────────────────────────────────────
def call_gemini_weekly(api_key: str, member: str,
                       combined_logs: str, n_days: int) -> dict:
    prompt = (
        f"以下是「{member}」本週共 {n_days} 天的工作日誌，"
        "請提煉本週核心工作，以 JSON 格式回傳一個欄位：\n"
        "1. bullets：本週核心工作總結，陣列包含「恰好 3 個」繁體中文條列要點（每點 40 字以內）\n\n"
        "只回傳 JSON 物件，不要加任何說明文字或 markdown 標記。\n"
        '範例：{"bullets":["完成...","推進...","協助..."]}\n\n'
        f"---\n{combined_logs[:8000]}"
    )
    raw = _call_gemini(api_key, prompt, 512)

    # 確保 bullets 欄位格式正確
    bullets = raw.get("bullets", [])
    if not isinstance(bullets, list):
        bullets = [str(bullets), "—", "—"]
    while len(bullets) < 3:
        bullets.append("（無資料）")
    raw["bullets"] = [str(b) for b in bullets[:3]]
    return raw


# ─────────────────────────────────────────────
# Gemini 共用呼叫函式
# ─────────────────────────────────────────────
def _call_gemini(api_key: str, prompt: str, max_tokens: int) -> dict:
    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {
            "temperature": 0.2,
            "maxOutputTokens": max_tokens,
            "thinkingConfig": {"thinkingBudget": 0},   # 關閉思考模式
        },
    }
    try:
        resp = requests.post(
            f"{GEMINI_URL}?key={api_key}", json=payload, timeout=60
        )
        resp.raise_for_status()
        data = resp.json()

        parts = data["candidates"][0]["content"]["parts"]
        text  = "".join(p.get("text", "") for p in parts if not p.get("thought", False))
        text  = re.sub(r"^```(?:json)?\s*", "", text.strip())
        text  = re.sub(r"\s*```$", "", text).strip()
        m     = re.search(r"\{[\s\S]*\}", text)
        if m:
            text = m.group(0)
        return json.loads(text)
    except Exception as exc:
        print(f"    ⚠️  Gemini API 呼叫失敗：{exc}")
        return {
            "summary": "（API 錯誤）", "completion_rate": 0, "rag_status": "red",
            "bullets": ["（API 錯誤，無法分析）", "—", "—"],
            "avg_completion_rate": 0,
        }


# ─────────────────────────────────────────────
# HTML 組裝工具
# ─────────────────────────────────────────────
def _mc(member: str) -> str:
    """取得成員標籤顏色"""
    return _MEMBER_COLORS.get(member.lower(), "#607d8b")


def _rag_badge(status: str, label_map=None) -> str:
    """產生 RAG 燈號 HTML badge（用於表格）"""
    label_map = label_map or {"green": "🟢 已達成", "yellow": "� 已達成", "red": "🔴 待改善"}
    color, _ = _RAG_META.get(status, ("#6c757d", ""))
    label = label_map.get(status, "⚪ 未知")
    return f'<span class="rag-badge" style="background:{color}">{label}</span>'


# 日誌明細：表格列
def build_table_rows(records: list) -> str:
    if not records:
        return '<tr><td colspan="5" class="no-data">目前尚無符合條件的日誌資料</td></tr>'
    rows = []
    for idx, r in enumerate(records, 1):
        rate = r["completion_rate"]
        bc   = "#28a745" if rate >= 60 else "#dc3545"
        rows.append(
            f'          <tr>\n'
            f'            <td>{idx}</td>\n'
            f'            <td><span class="member-tag member-{r["member"].lower()}">'
            f'{r["member"]}</span></td>\n'
            f'            <td>{r["date"]}</td>\n'
            f'            <td class="summary-cell">{r["summary"]}</td>\n'
            f'            <td>\n'
            f'              <div class="progress-bar-wrap">\n'
            f'                <div class="progress-bar-fill" style="width:{rate}%;background:{bc}"></div>\n'
            f'                <span class="progress-label">{rate}%</span>\n'
            f'              </div>\n'
            f'            </td>\n'
            f'          </tr>'
        )
    return "\n".join(rows)


# 日誌明細：成員統計卡片
def build_stats_cards(records: list, members: list) -> str:
    cards = []
    for member in members:
        mrs = [r for r in records if r["member"] == member]
        if not mrs:
            avg, desc, cls = 0, "尚無資料", "card-empty"
        else:
            avg  = round(sum(r["completion_rate"] for r in mrs) / len(mrs))
            desc = f"{len(mrs)} 筆日誌"
            cls  = "card-green" if avg >= 80 else "card-yellow" if avg >= 60 else "card-red"
        cards.append(
            f'        <div class="stat-card {cls}">\n'
            f'          <div class="stat-name">{member}</div>\n'
            f'          <div class="stat-rate">{avg}%</div>\n'
            f'          <div class="stat-desc">{desc}</div>\n'
            f'        </div>'
        )
    return "\n".join(cards)


# 週彙整：有資料的成員卡片
def build_member_card(member: str, result: dict, n: int) -> str:
    rate = int(result.get("avg_completion_rate", 0))
    rag  = str(result.get("rag_status", "red")).lower()
    rag_color, rag_label = _RAG_META.get(rag, ("#6c757d", "⚪ 未知"))
    bc   = "#28a745" if rate >= 60 else "#dc3545"
    buls = "\n".join(f'        <li>{b}</li>' for b in result.get("bullets", []))
    return (
        f'      <div class="member-card rag-{rag}">\n'
        f'        <div class="card-header">\n'
        f'          <span class="member-tag" style="background:{_mc(member)}">{member}</span>\n'
        f'          <span class="rag-pill" style="background:{rag_color}">{rag_label}</span>\n'
        f'        </div>\n'
        f'        <div class="rate-row">\n'
        f'          <div class="rate-num" style="color:{bc}">{rate}%</div>\n'
        f'          <div class="rate-meta">\n'
        f'            <div class="rate-label">週平均達成率</div>\n'
        f'            <div class="rate-days">{n} 天日誌</div>\n'
        f'          </div>\n'
        f'        </div>\n'
        f'        <div class="progress-track">\n'
        f'          <div class="progress-fill" style="width:{rate}%;background:{bc}"></div>\n'
        f'        </div>\n'
        f'        <div class="divider"></div>\n'
        f'        <div class="bullets-label">📌 本週核心工作</div>\n'
        f'        <ul class="bullet-list">\n{buls}\n        </ul>\n'
        f'      </div>'
    )


# 週彙整：無資料的成員卡片
def build_no_data_card(member: str) -> str:
    return (
        f'      <div class="member-card rag-none">\n'
        f'        <div class="card-header">\n'
        f'          <span class="member-tag" style="background:{_mc(member)}">{member}</span>\n'
        f'          <span class="rag-pill" style="background:#adb5bd">⚪ 無資料</span>\n'
        f'        </div>\n'
        f'        <div class="no-data-msg">本週無日誌記錄</div>\n'
        f'      </div>'
    )


# ─────────────────────────────────────────────
# 分數覆蓋輔助
# ─────────────────────────────────────────────
def _apply_override(override, ai_score: int) -> int:
    """
    override 可以是：
      - None          → 使用 AI 分數
      - int           → 固定值（如 70）
      - [min, max]    → 在範圍內隨機取整數（如 [69, 85]）
    """
    if override is None:
        return ai_score
    if isinstance(override, list) and len(override) == 2:
        return random.randint(int(override[0]), int(override[1]))
    return int(override)


# ─────────────────────────────────────────────
# 主程式
# ─────────────────────────────────────────────
def main() -> None:
    print("=" * 54)
    print("  HR 工作日誌合併報告產生器")
    print("=" * 54)
    print()

    # 1. 命令列參數
    ref_date: date | None = None
    if len(sys.argv) >= 2:
        try:
            ref_date = date.fromisoformat(sys.argv[1])
            print(f"📌 使用指定參考日期：{sys.argv[1]}")
        except ValueError:
            print("⚠️  日期格式錯誤（應為 YYYY-MM-DD），改用今天")

    # 2. 載入設定
    config  = load_config()
    api_key = config.get("gemini_api_key", "")
    if not api_key or api_key == "YOUR_GEMINI_API_KEY":
        print("❌ 請先在 config.json 中填入有效的 gemini_api_key")
        sys.exit(1)

    members: list[str]        = config.get("members", [])
    days_back: int             = int(config.get("days_to_look_back", 30))
    score_overrides: dict[str, int] = config.get("score_overrides", {})

    # 3. 計算週期
    week_start, week_end = get_week_range(ref_date)
    week_label = f"{fmt_date_zh(week_start)} ~ {fmt_date_zh(week_end)}"
    print(f"📅 本週週期：{week_label}")
    print()

    # 4. 讀取 Outlook（只讀一次）
    folder_path: str = config.get("outlook_folder_path", "")
    print(f"📬 讀取 Outlook（近 {days_back} 天）...")
    all_emails = get_emails(config["email_subjects"], days_back, folder_path)
    print(f"   找到 {len(all_emails)} 封符合條件的信件\n")

    # 5. 辨識發信人
    for e in all_emails:
        e["member"] = identify_sender(
            e["sender_name"], e["sender_email"], config["sender_mapping"]
        )

    # 6. 日誌明細只保留週期開始日以後的信件（週期開始 = week_start）
    all_emails = [e for e in all_emails if e["date_obj"] >= week_start]
    print(f"   筛選後：{len(all_emails)} 封（{week_start.strftime('%m/%d')} 起）\n")

    # ── 日誌明細：分析所有信件 ──────────────
    print("━" * 54)
    print("  [日誌明細] 逐封分析...")
    print("━" * 54)
    daily_records: list[dict] = []
    for e in all_emails:
        member = e["member"]
        if member not in members:
            print(f"  ⏭️  略過（非追蹤成員：{member}）")
            continue
        short = e["subject"][:28] + ("..." if len(e["subject"]) > 28 else "")
        print(f"  🔍 [{e['date']}] {member} — {short}")
        r = call_gemini_single(api_key, e["body"])
        final_rate = _apply_override(score_overrides.get(member), int(r.get("completion_rate", 0)))
        daily_records.append({
            "member":          member,
            "date":            e["date"],
            "summary":         str(r.get("summary", "（無法解析）")),
            "completion_rate": final_rate,
            "rag_status":      "green" if final_rate >= 60 else "red",
        })

    print()

    # ── 週彙整：分析本週信件 ──────────────
    print("━" * 54)
    print("  [週彙整] 本週信件彙整分析...")
    print("━" * 54)
    week_logs: dict[str, list[str]] = {m: [] for m in members}
    for e in all_emails:
        if e["member"] in members and week_start <= e["date_obj"] <= week_end:
            week_logs[e["member"]].append(
                f"=== {e['date']} ===\n{e['body'].strip()}"
            )

    weekly_results: dict[str, tuple[dict, int] | None] = {}
    for member in members:
        logs = week_logs[member]
        n    = len(logs)
        print(f"  📝 [{member}] 共 {n} 天日誌", end="")
        if n == 0:
            print(" → 略過（無資料）")
            weekly_results[member] = None
            continue

        # 達成率由日誌明細直接計算平均，不交由 Gemini 判斷（避免每次結果不一致）
        week_daily = [
            r for r in daily_records
            if r["member"] == member
            and week_start <= date.fromisoformat(r["date"]) <= week_end
        ]
        avg_rate = round(sum(r["completion_rate"] for r in week_daily) / len(week_daily)) if week_daily else 0
        rag      = "green" if avg_rate >= 60 else "red"

        print(f" → 呼叫 Gemini 提煉摘要...（達成率由日誌平均計算：{avg_rate}%）")
        r = call_gemini_weekly(api_key, member, "\n\n".join(logs), n)
        # 將穩定的達成率寫入，覆蓋 Gemini 的評估
        r["avg_completion_rate"] = avg_rate
        r["rag_status"]          = rag
        weekly_results[member] = (r, n)
        print(f"    ✅ 達成率 {avg_rate}%，狀態：{rag}")

    print()

    # ── 統計 ──────────────
    # 日誌明細統計
    green_d  = sum(1 for r in daily_records if r["rag_status"] == "green")
    yellow_d = sum(1 for r in daily_records if r["rag_status"] == "yellow")
    red_d    = sum(1 for r in daily_records if r["rag_status"] == "red")

    # 週彙整統計
    valid_w    = [(r, n) for r, n in (v for v in weekly_results.values() if v)]
    team_avg   = round(sum(r.get("avg_completion_rate", 0) for r, _ in valid_w) / len(valid_w)) if valid_w else 0
    green_w    = sum(1 for r, _ in valid_w if r.get("rag_status") == "green")
    yellow_w   = 0   # 週彙整僅分達成/待改善兩段，無 yellow
    red_w      = sum(1 for r, _ in valid_w if r.get("rag_status") == "red")
    none_cnt   = len(members) - len(valid_w)
    team_color = "#28a745" if team_avg >= 60 else "#dc3545"

    # ── 週追蹤報告：分析項目目標 ────────────────────────────────
    tracker_rows_html = ""
    tracker_stats: dict = {"total": 0, "done": 0, "green": 0, "yellow": 0, "red": 0, "none": 0}

    # 延遲載入 hr_tracker（避免模組層級循環 import）
    global _TRACKER_AVAILABLE
    try:
        from hr_tracker import (
            read_excel_projects,
            call_gemini_tracker,
            merge_ai_results,
            build_table_rows as build_tracker_rows,
            build_summary_stats as build_tracker_stats,
            export_excel as tracker_export_excel,
            EXCEL_PATH as TRACKER_EXCEL_PATH,
            EXCEL_TO_OUTLOOK,
            RAG_META as TRACKER_RAG_META,
            OWNER_COLORS,
            _owner_badge,
            _rag_badge as _tracker_rag_badge,
            _rate_bar as _tracker_rate_bar,
        )
        _TRACKER_AVAILABLE = True
    except ImportError:
        _TRACKER_AVAILABLE = False

    if _TRACKER_AVAILABLE and TRACKER_EXCEL_PATH.exists():
        print()
        print("━" * 54)
        print("  [週追蹤報告] 對應項目目標分析...")
        print("━" * 54)

        all_tracker_projects = read_excel_projects(TRACKER_EXCEL_PATH)

        # 使用本週信件，依成員分組日誌
        tracker_member_logs: dict[str, list[str]] = {}
        for e in all_emails:
            m = e["member"]
            if m not in EXCEL_TO_OUTLOOK.values():
                continue
            if e["date_obj"] < week_start or e["date_obj"] > week_end:
                continue
            if m not in tracker_member_logs:
                tracker_member_logs[m] = []
            header = f"=== {e['date']} | {e['subject']} ===\n"
            tracker_member_logs[m].append(header + e["body"])

        for outlook_name in EXCEL_TO_OUTLOOK.values():
            member_projects = [p for p in all_tracker_projects if p["outlook_name"] == outlook_name]
            if not member_projects:
                continue
            display_name = next((k for k, v in EXCEL_TO_OUTLOOK.items() if v == outlook_name), outlook_name)
            logs_combined = "\n\n".join(tracker_member_logs.get(outlook_name, []))

            if not logs_combined.strip():
                print(f"  ⏭️  [{display_name}] 本週無日誌，維持 Excel 基準值")
                for p in member_projects:
                    p["weekly_result"] = "本週無日誌提交"
                    p["ai_rate"]       = p["base_rate"]
                    p["rag"]           = "本週無資料"
                    p["blocker"]       = ""
                continue

            print(f"  🤖 [{display_name}] {len(tracker_member_logs.get(outlook_name, []))} 封日誌"
                  f" → 分析 {len(member_projects)} 個項目...")
            ai_results = call_gemini_tracker(api_key, display_name, logs_combined, member_projects)
            merge_ai_results(member_projects, ai_results)

            for p in member_projects:
                rag_icon = {"已完成": "✅", "正常推進": "🟢", "待改善": "🟡", "卡關": "🔴"}.get(p["rag"], "⚪")
                print(f"    {rag_icon}  {p['goal'][:28]}  → {p['ai_rate']}%  {p['rag']}")

        tracker_stats     = build_tracker_stats(all_tracker_projects)
        tracker_rows_html = build_tracker_rows(all_tracker_projects)

        # 匯出 Excel
        try:
            excel_out_path, excel_b64 = tracker_export_excel(
                all_tracker_projects, ref_date, week_start, week_end
            )
            excel_filename = excel_out_path.name
            print(f"  📊 已匯出 Excel：{excel_out_path}")
        except Exception as _exc:
            print(f"  ⚠️  Excel 匯出失敗：{_exc}")
            excel_b64, excel_filename = "", ""
    else:
        tracker_rows_html = '<tr><td colspan="9" class="tbl-no-data">尚未設定 Excel 或 hr_tracker.py 無法載入</td></tr>'
        excel_b64, excel_filename = "", ""

    # ── 組裝 HTML ──────────────
    member_cards = "\n".join(
        build_member_card(m, r, n) if (v := weekly_results.get(m)) and (r := v[0]) and (n := v[1]) else build_no_data_card(m)
        for m in members
    )

    if not TEMPLATE_PATH.exists():
        print(f"❌ 找不到範本：{TEMPLATE_PATH}")
        sys.exit(1)

    with open(TEMPLATE_PATH, "r", encoding="utf-8") as f:
        html = f.read()

    replacements = {
        "{{WEEK_LABEL}}":     week_label,
        "{{HTML_FILENAME}}": f"人資部週追蹤報告({week_start.strftime('%m%d')}~{week_end.strftime('%m%d')}).html",
        "{{GENERATED_AT}}":   datetime.now().strftime("%Y-%m-%d %H:%M"),
        "{{TEAM_AVG}}":       str(team_avg),
        "{{TEAM_AVG_COLOR}}": team_color,
        "{{GREEN_COUNT_W}}":  str(green_w),
        "{{YELLOW_COUNT_W}}": str(yellow_w),
        "{{RED_COUNT_W}}":    str(red_w),
        "{{NONE_COUNT}}":     str(none_cnt),
        "{{MEMBER_COUNT}}":   str(len(members)),
        "{{MEMBER_CARDS}}":   member_cards,
        "{{DAYS_BACK}}":      week_start.strftime("%m/%d"),
        "{{TOTAL_COUNT}}":    str(len(daily_records)),
        "{{GREEN_COUNT_D}}":  str(green_d),
        "{{YELLOW_COUNT_D}}": str(yellow_d),
        "{{RED_COUNT_D}}":    str(red_d),
        "{{STATS_CARDS}}":    build_stats_cards(daily_records, members),
        "{{TABLE_ROWS}}":     build_table_rows(daily_records),
        # 週追蹤報告
        "{{TRACKER_TOTAL}}":  str(tracker_stats["total"]),
        "{{TRACKER_DONE}}":   str(tracker_stats["done"]),
        "{{TRACKER_GREEN}}":  str(tracker_stats["green"]),
        "{{TRACKER_YELLOW}}": str(tracker_stats["yellow"]),
        "{{TRACKER_RED}}":    str(tracker_stats["red"]),
        "{{TRACKER_NONE}}":   str(tracker_stats["none"]),
        "{{TRACKER_ROWS}}":   tracker_rows_html,
        # Excel 匯出
        "{{EXCEL_FILENAME}}": excel_filename,
        "{{EXCEL_BASE64}}":   excel_b64,
        "{{EXCEL_DOWNLOAD_BTN}}": (
            f'<a class="export-btn" '
            f'href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{excel_b64}" '
            f'download="{excel_filename}">&#128229; 匯出本週 Excel 報告</a>'
            if excel_b64 else
            '<span class="export-btn export-btn-disabled">&#128229; Excel 未產生</span>'
        ),
    }
    for k, v in replacements.items():
        html = html.replace(k, v)

    with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
        f.write(html)

    # 複製 HTML 到「資料」資料夾，與 Excel 集中存放
    html_filename = f"人資部週追蹤報告({week_start.strftime('%m%d')}~{week_end.strftime('%m%d')}).html"
    html_out_path = SCRIPT_DIR / "資料" / html_filename
    import shutil as _shutil
    _shutil.copy(str(OUTPUT_PATH), str(html_out_path))
    print(f"  📄 已複製 HTML：{html_out_path}")

    print(f"✅ 完成！已產生：{OUTPUT_PATH}")
    print(
        f"   日誌明細：{len(daily_records)} 筆 ｜"
        f" 🟢 {green_d} / 🟡 {yellow_d} / 🔴 {red_d}"
    )
    print(
        f"   週彙整：有資料 {len(valid_w)} 人 ｜"
        f" 無資料 {none_cnt} 人 ｜ 團隊平均 {team_avg}%"
    )


if __name__ == "__main__":
    main()
