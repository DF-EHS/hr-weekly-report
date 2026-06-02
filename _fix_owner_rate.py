import re

with open("report.html", encoding="utf-8") as f:
    html = f.read()

# 1. 負責人：去掉 owner-badge span，只留名字文字
html = re.sub(r'<span class="owner-badge"[^>]*>([^<]+)</span>', r'\1', html)

# 2. 達成率：rate-input 改為深色文字
html = html.replace(
    "font-size: .78rem; font-weight: 700; color: #fff;\n      text-shadow: 0 1px 2px rgba(0,0,0,.4);\n      text-align: right; outline: none; cursor: text; z-index: 1;",
    "font-size: .78rem; font-weight: 700; color: #333;\n      text-align: right; outline: none; cursor: text; z-index: 1;"
)

with open("report.html", "w", encoding="utf-8") as f:
    f.write(html)
print("完成")
