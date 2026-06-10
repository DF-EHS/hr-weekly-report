#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
sharepoint_archive.py - 把最新的「人資部週追蹤報告」複製到 OneDrive 同步資料夾，
由 OneDrive 自動上傳至 SharePoint（降本增效週追蹤會議 / 人資部）。

機制：此 SharePoint 文件庫已本機同步，複製進同步夾即等同上傳，免登入/API。
用法：python sharepoint_archive.py
"""
import sys
import io
import shutil
from pathlib import Path

# 排程執行時強制 UTF-8 輸出
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")
except (AttributeError, ValueError):
    pass

SCRIPT_DIR = Path(__file__).parent
DATA_DIR   = SCRIPT_DIR / "資料"

# OneDrive 同步的 SharePoint「降本增效週追蹤會議 / 人資部」資料夾
SP_DIR = (
    Path.home()
    / "OneDrive - 大豐環保科技股份有限公司"
    / "總經理室 - 降本增效週追蹤會議"
    / "人資部"
)

# 只上傳這兩種附檔（排除範本與備份檔，靠檔名的「(週期)」括號過濾）
EXTS = ("html", "xlsx")


def latest_report(ext: str) -> Path | None:
    """回傳 資料/ 中最新（依修改時間）的『人資部週追蹤報告(週期).ext』。"""
    candidates = sorted(
        DATA_DIR.glob(f"人資部週追蹤報告(*).{ext}"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    return candidates[0] if candidates else None


def main() -> None:
    print("=" * 50)
    print("  SharePoint 自動存檔（人資部週追蹤報告）")
    print("=" * 50)

    if not SP_DIR.exists():
        print(f"❌ 找不到 OneDrive 同步資料夾：\n   {SP_DIR}")
        print("   請確認該 SharePoint 文件庫仍在本機同步。")
        sys.exit(1)

    copied = 0
    for ext in EXTS:
        src = latest_report(ext)
        if src is None:
            print(f"⚠️  資料夾中找不到 .{ext} 報告，略過")
            continue
        dst = SP_DIR / src.name
        shutil.copy2(str(src), str(dst))
        print(f"✅ {src.name}  →  SharePoint 同步夾")
        copied += 1

    print(f"\n完成：複製 {copied} 個檔案（OneDrive 將自動上傳至 SharePoint）")
    if copied == 0:
        sys.exit(1)


if __name__ == "__main__":
    main()
