#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
update_logs.py - 週工作日誌追蹤報告產生器

功能：
  1. 讀取 Outlook 本地收件夾，篩選主旨含「工作日誌」或「工作報告」的信件
  2. 辨識發信人（Bella / An / Gloom）
  3. 呼叫 Gemini 2.5 Flash API 分析日誌內容
  4. 套用 dashboard_template.html 範本產生 index.html 進度追蹤網頁

使用前準備：
  pip install pywin32 requests
  並在 config.json 中填入有效的 gemini_api_key
"""

import json
import os
import re
import sys
from datetime import datetime, timedelta
from pathlib import Path

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
# 路徑設定
# ─────────────────────────────────────────────
SCRIPT_DIR    = Path(__file__).parent
CONFIG_PATH   = SCRIPT_DIR / "config.json"
TEMPLATE_PATH = SCRIPT_DIR / "dashboard_template.html"
OUTPUT_PATH   = SCRIPT_DIR / "index.html"

# Gemini 2.5 Flash 端點
GEMINI_MODEL  = "gemini-2.5-flash"
GEMINI_URL    = (
    "https://generativelanguage.googleapis.com/v1beta/models/"
    f"{GEMINI_MODEL}:generateContent"
)


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
# 發信人辨識
# ─────────────────────────────────────────────
def identify_sender(sender_name: str, sender_email: str, mapping: dict) -> str:
    """
    根據 mapping 中的關鍵字辨識成員代稱。
    mapping key 為辨識關鍵字（不分大小寫），value 為代稱（Bella / An / Gloom）。
    """
    combined = f"{sender_name} {sender_email}".lower()
    for keyword, member in mapping.items():
        if keyword.lower() in combined:
            return member
    # 無法辨識時回傳原始名稱
    return sender_name.strip() or "未知"


# ─────────────────────────────────────────────
# Outlook 信件讀取
# ─────────────────────────────────────────────
def _find_folder(parent, name: str):
    """在 parent 的子資料夾中尋找指定名稱（不分大小寫），找不到回傳 None。"""
    for folder in parent.Folders:
        if folder.Name.strip().lower() == name.strip().lower():
            return folder
    return None


def get_outlook_emails(subjects: list, days_back: int,
                       folder_path: str = "") -> list:
    """
    連接 Outlook，讀取指定資料夾中符合主旨關鍵字的信件。
    folder_path：以「/」分隔的子資料夾路徑，例如 "人資部/每日工作報告"。
    留空則讀取預設收件匣。
    回傳 list of dict：subject / sender_name / sender_email / body / date
    """
    try:
        outlook   = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        base      = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
    except Exception as exc:
        print(f"❌ 無法連接 Outlook：{exc}")
        sys.exit(1)

    # 依 folder_path 逐層導航至目標子資料夾
    target = base
    if folder_path:
        parts = [p.strip() for p in folder_path.split("/") if p.strip()]
        for part in parts:
            found = _find_folder(target, part)
            if found is None:
                print(f"⚠️  找不到 Outlook 子資料夾「{part}」，改用收件匣讀取")
                target = base
                break
            target = found
        print(f"   📂 讀取資料夾：{target.Name}")

    cutoff = datetime.now() - timedelta(days=days_back)
    emails = []

    for item in target.Items:
        try:
            if item.Class != 43:   # 43 = olMail，略過非信件項目
                continue

            # 取得接收時間（去除時區資訊以便比較）
            received = item.ReceivedTime
            if not hasattr(received, "replace"):
                continue
            received_dt = received.replace(tzinfo=None)
            if received_dt < cutoff:
                continue

            # 主旨關鍵字篩選
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
            # 略過無法讀取的單封信件，繼續處理其他信件
            continue

    return emails


# ─────────────────────────────────────────────
# Gemini API 分析
# ─────────────────────────────────────────────
def call_gemini(api_key: str, body: str) -> dict:
    """
    呼叫 Gemini 2.5 Flash API 分析日誌內容，回傳：
      { summary: str, completion_rate: int, rag_status: str }
    """
    prompt = (
        "請分析以下工作日誌內容，並以 JSON 格式回傳三個欄位：\n"
        "1. summary：50 字以內的工作成果摘要（繁體中文）\n"
        "2. completion_rate：整體工作達成率，0 到 100 之間的整數\n"
        "3. rag_status：依達成率判定燈號（green = 80 以上、yellow = 60-79、red = 59 以下）\n\n"
        "只回傳 JSON 物件，不要加任何說明文字或 markdown 標記。\n\n"
        f"工作日誌內容：\n{body[:3000]}"
    )

    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {
            "temperature": 0.2,
            "maxOutputTokens": 512,
            "thinkingConfig": {"thinkingBudget": 0},  # 關閉思考模式，確保回傳純 JSON
        },
    }

    try:
        resp = requests.post(
            f"{GEMINI_URL}?key={api_key}",
            json=payload,
            timeout=30,
        )
        resp.raise_for_status()
        data = resp.json()

        # 收集所有 parts 的 text（思考內容在 thought=True 的 part 中）
        parts = data["candidates"][0]["content"]["parts"]
        text = ""
        for part in parts:
            if not part.get("thought", False):
                text += part.get("text", "")
        text = text.strip()

        # 移除 Gemini 可能回傳的 markdown code block 包裹
        text = re.sub(r"^```(?:json)?\s*", "", text)
        text = re.sub(r"\s*```$", "", text).strip()

        # 從文字中抓取第一個 {...} JSON 區塊（防止前後有多餘說明文字）
        json_match = re.search(r"\{[^{}]*\}", text, re.DOTALL)
        if json_match:
            text = json_match.group(0)

        result = json.loads(text)
        return {
            "summary":         str(result.get("summary", "（無法解析）")),
            "completion_rate": int(result.get("completion_rate", 0)),
            "rag_status":      str(result.get("rag_status", "red")).lower(),
        }
    except Exception as exc:
        print(f"    ⚠️  Gemini API 呼叫失敗：{exc}")
        return {
            "summary":         "（API 錯誤，無法分析）",
            "completion_rate": 0,
            "rag_status":      "red",
        }


# ─────────────────────────────────────────────
# HTML 組裝工具
# ─────────────────────────────────────────────
_RAG_COLORS = {
    "green":  ("#28a745", "🟢 已達成"),
    "yellow": ("#e6a800", "🟡 部分達成"),
    "red":    ("#dc3545", "🔴 待改善"),
}


def _rag_badge(status: str) -> str:
    """產生 RAG 燈號 HTML badge"""
    color, label = _RAG_COLORS.get(status, ("#6c757d", "⚪ 未知"))
    return f'<span class="rag-badge" style="background:{color}">{label}</span>'


def build_table_rows(records: list) -> str:
    """將紀錄清單轉為 HTML 表格橫列字串"""
    if not records:
        return '<tr><td colspan="6" class="no-data">目前尚無符合條件的日誌資料</td></tr>'

    rows = []
    for idx, r in enumerate(records, 1):
        rate = r["completion_rate"]
        # 依達成率決定進度條顏色
        bar_color = "#28a745" if rate >= 80 else "#e6a800" if rate >= 60 else "#dc3545"

        rows.append(
            f"          <tr>\n"
            f"            <td>{idx}</td>\n"
            f"            <td>"
            f'<span class="member-tag member-{r["member"].lower()}">'
            f'{r["member"]}</span></td>\n'
            f"            <td>{r['date']}</td>\n"
            f"            <td class=\"summary-cell\">{r['summary']}</td>\n"
            f"            <td>\n"
            f"              <div class=\"progress-bar-wrap\">\n"
            f"                <div class=\"progress-bar-fill\" "
            f'style="width:{rate}%;background:{bar_color}"></div>\n'
            f"                <span class=\"progress-label\">{rate}%</span>\n"
            f"              </div>\n"
            f"            </td>\n"
            f"            <td>{_rag_badge(r['rag_status'])}</td>\n"
            f"          </tr>"
        )

    return "\n".join(rows)


def generate_stats_cards(records: list, members: list) -> str:
    """產生成員平均達成率統計卡片"""
    cards = []
    for member in members:
        member_records = [r for r in records if r["member"] == member]

        if not member_records:
            avg_rate   = 0
            desc_text  = "尚無資料"
            card_class = "card-empty"
        else:
            avg_rate   = round(
                sum(r["completion_rate"] for r in member_records) / len(member_records)
            )
            desc_text  = f"{len(member_records)} 筆日誌"
            card_class = (
                "card-green"  if avg_rate >= 80 else
                "card-yellow" if avg_rate >= 60 else
                "card-red"
            )

        cards.append(
            f'        <div class="stat-card {card_class}">\n'
            f'          <div class="stat-name">{member}</div>\n'
            f'          <div class="stat-rate">{avg_rate}%</div>\n'
            f'          <div class="stat-desc">{desc_text}</div>\n'
            f'        </div>'
        )

    return "\n".join(cards)


# ─────────────────────────────────────────────
# 主程式
# ─────────────────────────────────────────────
def main():
    print("=" * 45)
    print("  週工作日誌追蹤報告產生器")
    print("=" * 45)
    print()

    # 1. 載入設定
    config  = load_config()
    api_key = config.get("gemini_api_key", "")
    if not api_key or api_key == "YOUR_GEMINI_API_KEY":
        print("❌ 請先在 config.json 中填入有效的 gemini_api_key")
        sys.exit(1)

    days_back = int(config.get("days_to_look_back", 30))
    members   = config.get("members", ["Bella", "An", "Gloom"])

    # 2. 讀取 Outlook 信件
    folder_path = config.get("outlook_folder_path", "")
    print(f"📬 讀取 Outlook（近 {days_back} 天）...")
    emails = get_outlook_emails(config["email_subjects"], days_back, folder_path)
    print(f"   找到 {len(emails)} 封符合條件的信件\n")

    if not emails:
        print("ℹ️  沒有找到任何日誌信件，仍將產生空白報表。\n")

    # 3. 逐一分析
    records = []
    for email in emails:
        member = identify_sender(
            email["sender_name"],
            email["sender_email"],
            config["sender_mapping"],
        )
        short_subject = email["subject"][:30] + ("..." if len(email["subject"]) > 30 else "")
        print(f"  🔍 [{email['date']}] {member} — {short_subject}")

        # 過濾掉不在成員名單的寄件人（例如主管自己的報告）
        if member not in members:
            print(f"    ⏭️  略過（非追蹤成員：{member}）")
            continue

        analysis = call_gemini(api_key, email["body"])
        records.append({
            "member": member,
            "date":   email["date"],
            **analysis,
        })

    # 按日期降序排列（最新在上）
    records.sort(key=lambda x: x["date"], reverse=True)

    # 4. 讀取 HTML 範本
    if not TEMPLATE_PATH.exists():
        print(f"\n❌ 找不到範本檔：{TEMPLATE_PATH}")
        sys.exit(1)

    with open(TEMPLATE_PATH, "r", encoding="utf-8") as f:
        template = f.read()

    # 5. 計算統計數字
    now_str      = datetime.now().strftime("%Y年%m月%d日 %H:%M")
    total_count  = len(records)
    green_count  = sum(1 for r in records if r["rag_status"] == "green")
    yellow_count = sum(1 for r in records if r["rag_status"] == "yellow")
    red_count    = sum(1 for r in records if r["rag_status"] == "red")

    # 6. 填入範本佔位符
    html = (
        template
        .replace("{{GENERATED_AT}}",  now_str)
        .replace("{{DAYS_BACK}}",      str(days_back))
        .replace("{{TOTAL_COUNT}}",    str(total_count))
        .replace("{{GREEN_COUNT}}",    str(green_count))
        .replace("{{YELLOW_COUNT}}",   str(yellow_count))
        .replace("{{RED_COUNT}}",      str(red_count))
        .replace("{{STATS_CARDS}}",    generate_stats_cards(records, members))
        .replace("{{TABLE_ROWS}}",     build_table_rows(records))
    )

    # 7. 寫出 index.html
    with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
        f.write(html)

    print()
    print(f"✅ 完成！已產生：{OUTPUT_PATH}")
    print(
        f"   共處理 {total_count} 筆日誌"
        f" — 🟢 {green_count} 已達成"
        f" / 🟡 {yellow_count} 部分達成"
        f" / 🔴 {red_count} 待改善"
    )


if __name__ == "__main__":
    main()
