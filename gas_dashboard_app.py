#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gas_dashboard_app.py
---------------------
簡易桌面小工具：一個按鈕同步瓦斯用量 Dashboard，並可選擇性新增帳單記錄。
雙擊 啟動瓦斯Dashboard工具.bat 開啟。
"""

import json
import re
import subprocess
import sys
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, scrolledtext

try:
    import openpyxl
except ImportError:
    openpyxl = None

REPO_ROOT = Path(__file__).parent
DEFAULT_EXCEL = r"D:\AI application code\gas-bill-analysis\Gas usage.xlsm"
HTML_FILE = REPO_ROOT / "index.html"
PY_FILE = REPO_ROOT / "update_dashboard.py"
SHEET = "Daily_Summary"


def load_known_bills():
    content = PY_FILE.read_text(encoding="utf-8")
    m = re.search(r"KNOWN_BILLS = (\[.*?\])", content, flags=re.DOTALL)
    if not m:
        return []
    # KNOWN_BILLS is written with Python-style dict literals (double-quoted keys), valid JSON
    text = m.group(1).rstrip()
    text = re.sub(r",\s*\]$", "]", text)  # tolerate trailing comma
    return json.loads(text)


def save_known_bills(bills):
    content = PY_FILE.read_text(encoding="utf-8")
    bills_src = "[\n" + "\n".join(
        f'    {{"deg": {b["deg"]}, "rate": {b["rate"]}, "base": {b["base"]}, "total": {b["total"]}}},'
        for b in bills
    ) + "\n]"
    new_content, count = re.subn(r"KNOWN_BILLS = \[.*?\]", f"KNOWN_BILLS = {bills_src}",
                                  content, count=1, flags=re.DOTALL)
    if count == 0:
        raise RuntimeError("找不到 KNOWN_BILLS 區塊")
    PY_FILE.write_text(new_content, encoding="utf-8")


def read_daily(excel_path):
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb[SHEET]
    rows = []
    for date, year, quarter, month, day, usage in ws.iter_rows(min_row=2, values_only=True):
        if date is None:
            continue
        rows.append([date.strftime("%Y-%m-%d"), round(float(usage), 3) if usage is not None else 0])
    wb.close()
    return rows


def update_html(daily, bills):
    content = HTML_FILE.read_text(encoding="utf-8")
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
        raise RuntimeError("找不到 DATA BLOCK 標記")
    new_content = new_content.replace("{{LATEST_RATE}}", str(latest_rate))
    HTML_FILE.write_text(new_content, encoding="utf-8")
    return len(daily)


def git_push(message, files):
    for f in files:
        subprocess.run(["git", "add", f], check=True, cwd=REPO_ROOT)
    result = subprocess.run(["git", "commit", "-m", message], capture_output=True, text=True, cwd=REPO_ROOT)
    out = result.stdout + result.stderr
    if "nothing to commit" in out:
        return "資料未變動，不需要 push"
    if result.returncode != 0:
        raise RuntimeError(f"commit 失敗：{result.stderr}")
    subprocess.run(["git", "push"], check=True, cwd=REPO_ROOT)
    return "已成功 push 到 GitHub！"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("瓦斯用量 Dashboard 工具")
        self.geometry("520x420")
        self.configure(bg="#FFF7ED")

        tk.Label(self, text="🔥 瓦斯用量 Dashboard 工具", font=("Microsoft JhengHei", 14, "bold"),
                 bg="#FFF7ED", fg="#9A3412").pack(pady=(16, 4))
        tk.Label(self, text="同步用量資料，並可選擇新增一筆實際帳單", font=("Microsoft JhengHei", 9),
                 bg="#FFF7ED", fg="#6B7280").pack(pady=(0, 12))

        btn_frame = tk.Frame(self, bg="#FFF7ED")
        btn_frame.pack(pady=4)

        tk.Button(btn_frame, text="📥 只同步用量", font=("Microsoft JhengHei", 11),
                  bg="#2563EB", fg="white", padx=14, pady=8, relief="flat",
                  command=self.sync_only).grid(row=0, column=0, padx=6)

        tk.Button(btn_frame, text="🧾 新增帳單並同步", font=("Microsoft JhengHei", 11),
                  bg="#EA580C", fg="white", padx=14, pady=8, relief="flat",
                  command=self.add_bill_and_sync).grid(row=0, column=1, padx=6)

        self.log = scrolledtext.ScrolledText(self, height=14, font=("Consolas", 9), bg="white")
        self.log.pack(fill="both", expand=True, padx=16, pady=12)

    def write_log(self, text):
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.update_idletasks()

    def sync_only(self):
        self._run_sync(new_bill=None)

    def add_bill_and_sync(self):
        bill = self.prompt_bill_dialog()
        if bill is None:
            return
        self._run_sync(new_bill=bill)

    def prompt_bill_dialog(self):
        result = {}

        dlg = tk.Toplevel(self)
        dlg.title("新增帳單記錄")
        dlg.configure(bg="#FFF7ED")
        dlg.geometry("320x260")
        dlg.grab_set()

        fields = [("計費氣量（度數）", "deg", ""), ("單價", "rate", ""),
                  ("基本費", "base", "300"), ("合計金額（可留空自動試算）", "total", "")]
        entries = {}
        for label, key, default in fields:
            tk.Label(dlg, text=label, bg="#FFF7ED", font=("Microsoft JhengHei", 9)).pack(anchor="w", padx=16, pady=(10, 0))
            e = tk.Entry(dlg, font=("Microsoft JhengHei", 10))
            e.insert(0, default)
            e.pack(fill="x", padx=16)
            entries[key] = e

        def submit():
            try:
                deg = float(entries["deg"].get().strip())
                rate = float(entries["rate"].get().strip())
                base = float(entries["base"].get().strip() or "300")
                total_raw = entries["total"].get().strip()
                total = float(total_raw) if total_raw else round(deg * rate + base)
                result["bill"] = {"deg": deg, "rate": rate, "base": base, "total": total}
                dlg.destroy()
            except ValueError:
                messagebox.showerror("輸入錯誤", "請輸入正確的數字")

        tk.Button(dlg, text="確認新增", bg="#EA580C", fg="white", relief="flat",
                  command=submit).pack(pady=16)

        self.wait_window(dlg)
        return result.get("bill")

    def _run_sync(self, new_bill):
        self.log.delete("1.0", "end")
        try:
            if openpyxl is None:
                self.write_log("缺少 openpyxl，請執行：pip install openpyxl")
                return

            excel_path = Path(DEFAULT_EXCEL)
            if not excel_path.exists():
                self.write_log(f"找不到 Excel：{excel_path}")
                return

            self.write_log(f"讀取 {excel_path.name} ...")
            daily = read_daily(excel_path)
            self.write_log(f"讀取完成，共 {len(daily)} 天記錄")

            bills = load_known_bills()
            files_changed = ["index.html"]
            if new_bill:
                bills = bills + [new_bill]
                save_known_bills(bills)
                self.write_log(f"已新增帳單：{new_bill}")
                files_changed.append("update_dashboard.py")

            n = update_html(daily, bills)
            self.write_log(f"index.html 已更新（{n} 天 / {len(bills)} 筆帳單）")

            latest = daily[-1][0]
            msg = f"update dashboard: 最新資料至 {latest}"
            result = git_push(msg, files_changed)
            self.write_log(result)
            self.write_log(f"完成！Dashboard 已更新至 {latest}")
        except Exception as e:
            self.write_log(f"發生錯誤：{e}")


if __name__ == "__main__":
    App().mainloop()
