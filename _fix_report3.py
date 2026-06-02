"""
修正 report.html 週彙整：
- 移除「部分達成」，改為只有「達成(>=60%)」和「待改善(<60%)」
- 更新概覽卡片統計
"""
import re

with open("report.html", encoding="utf-8") as f:
    html = f.read()

# ── 1. 成員卡片：把「部分達成」badge 改為「達成」，顏色改綠 ──────────────
# rag-pill 的 badge，黃色部分達成 → 綠色達成
html = html.replace(
    'class="rag-pill" style="background:#e6a800">🟡 部分達成</span>',
    'class="rag-pill" style="background:#28a745">🟢 達成</span>'
)

# 成員卡片外框 class：rag-yellow → rag-green
html = html.replace('class="member-card rag-yellow"', 'class="member-card rag-green"')

# 進度條與數字顏色：黃色 #e6a800 → 綠色 #28a745（只改成員卡片內的）
# 用更精確的方式：找到 rate-num 與 progress-fill 的黃色
html = re.sub(
    r'(<div class="rate-num" style="color:)#e6a800(">)',
    r'\g<1>#28a745\2',
    html
)
html = re.sub(
    r'(<div class="progress-fill" style="width:\d+%;background:)#e6a800("></div>)',
    r'\g<1>#28a745\2',
    html
)

# ── 2. 概覽卡片：移除「部分達成」那格 ──────────────────────────────────
# 找到 ov-yellow 整個 div 移除
html = re.sub(
    r'\s*<div class="overview-card ov-yellow">.*?</div>',
    '',
    html,
    flags=re.DOTALL
)

# ── 3. 重算概覽統計數字 ──────────────────────────────────────────────────
# 從成員卡片中統計 green / red
green_w = len(re.findall(r'class="member-card rag-green"', html))
red_w   = len(re.findall(r'class="member-card rag-red"', html))

def replace_overview_num(html, css_class, new_val):
    pattern = rf'(<div class="overview-card {re.escape(css_class)}">\s*<div class="num">)\d+(</div>)'
    result, n = re.subn(pattern, rf'\g<1>{new_val}\2', html)
    if n:
        print(f"  更新 [{css_class}] → {new_val}")
    return result

html = replace_overview_num(html, "ov-green", green_w)
html = replace_overview_num(html, "ov-red",   red_w)

# ── 4. 週彙整概覽文字「部分達成」→「待改善」（如有殘留）──────────────────
html = html.replace('🟡 部分達成', '🟢 達成')

with open("report.html", "w", encoding="utf-8") as f:
    f.write(html)

print(f"完成：達成 {green_w} 人，待改善 {red_w} 人")
