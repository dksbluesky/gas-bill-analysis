#!/usr/bin/env python3
"""
update_dashboard.py
-------------------
讀取瓦斯用量 Excel (.xlsm) 的 Daily_Summary sheet，更新 index.html 的每日用量資料，
然後 git commit + push 到 GitHub。

實際帳單（KNOWN_BILLS）無法從 Excel 自動讀取，需手動維護於本檔的 KNOWN_BILLS 區塊
（資料來源：gas.e-letter.com.tw 歷史帳單）。

使用方式：
  python update_dashboard.py
  python update_dashboard.py --dry-run
"""

import argparse, json, re, subprocess, sys
from pathlib import Path

try:
    import openpyxl
except ImportError:
    print("缺少 openpyxl，請執行：pip install openpyxl")
    sys.exit(1)

DEFAULT_EXCEL = r"C:\Users\dk098\Documents\House related\Gas usage.xlsm"
HTML_FILE     = "index.html"
SHEET         = "Daily_Summary"

# 實際帳單記錄 — 從 gas.e-letter.com.tw 歷史帳單手動抄錄，新帳單請手動新增
KNOWN_BILLS = [
    {"deg": 1, "rate": 13.14, "base": 300, "total": 313},
    {"deg": 1, "rate": 13.14, "base": 300, "total": 313},
    {"deg": 2, "rate": 13.28, "base": 300, "total": 327},
    {"deg": 1, "rate": 13.41, "base": 300, "total": 313},
]


def read_daily(path: Path) -> list:
    print(f"讀取 {path.name} ...")
    wb = openpyxl.load_workbook(path, data_only=True)
    if SHEET not in wb.sheetnames:
        print(f"找不到工作表「{SHEET}」"); sys.exit(1)
    ws = wb[SHEET]
    rows = []
    for date, year, quarter, month, day, usage in ws.iter_rows(min_row=2, values_only=True):
        if date is None:
            continue
        rows.append([date.strftime("%Y-%m-%d"), round(float(usage), 3) if usage is not None else 0])
    wb.close()
    print(f"讀取完成，共 {len(rows)} 天記錄")
    return rows


def update_html(html_path: Path, daily: list, bills: list) -> bool:
    content = html_path.read_text(encoding="utf-8")
    daily_json = json.dumps(daily, ensure_ascii=False, separators=(",", ":"))
    bills_json = json.dumps(bills, ensure_ascii=False)
    latest_rate = bills[-1]["rate"] if bills else 13.0

    pattern = r"const DAILY = .*?;\n// 實際帳單記錄.*?\nconst KNOWN_BILLS = .*?;"
    replacement = (
        f"const DAILY = {daily_json};\n"
        f"// 實際帳單記錄（手動從 gas.e-letter.com.tw 抄錄，無法自動同步）\n"
        f"const KNOWN_BILLS = {bills_json};"
    )
    new_content, count = re.subn(pattern, replacement, content, flags=re.DOTALL)
    if count == 0:
        print("找不到 DATA BLOCK 標記"); return False

    new_content = new_content.replace("{{LATEST_RATE}}", str(latest_rate))

    html_path.write_text(new_content, encoding="utf-8")
    print(f"index.html 已更新（{len(daily)} 天 / {len(bills)} 筆帳單）")
    return True


def git_push(message: str, files: list):
    for f in files:
        subprocess.run(["git", "add", f], check=True)
    result = subprocess.run(["git", "commit", "-m", message], capture_output=True, text=True)
    if "nothing to commit" in result.stdout + result.stderr:
        print("資料未變動，不需要 push"); return
    if result.returncode != 0:
        print(f"commit 失敗：{result.stderr}"); sys.exit(1)
    subprocess.run(["git", "push"], check=True)
    print("已成功 push 到 GitHub！")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--excel", default=DEFAULT_EXCEL)
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--message", default="")
    args = parser.parse_args()

    repo_root = Path(__file__).parent
    excel_path = Path(args.excel) if Path(args.excel).is_absolute() else repo_root / args.excel
    html_path = repo_root / HTML_FILE

    if not excel_path.exists():
        print(f"找不到 Excel：{excel_path}"); sys.exit(1)

    daily = read_daily(excel_path)
    if not daily:
        print("沒有讀到資料"); sys.exit(1)

    if not update_html(html_path, daily, KNOWN_BILLS):
        sys.exit(1)

    if args.dry_run:
        print("dry-run，跳過 push"); return

    latest = daily[-1][0]
    msg = args.message or f"update dashboard: 最新資料至 {latest}"
    git_push(msg, [HTML_FILE])
    print(f"完成！Dashboard 已更新至 {latest}")


if __name__ == "__main__":
    main()
