# 瓦斯用量分析 Dashboard — Claude Code 操作指南

## 專案說明
追蹤家用瓦斯每日用量（無雙月帳單資料，僅有 Domotz/感應器記錄的每日用量），費用為估算值。

## 檔案結構
```
./
├── index.html              ← GitHub Pages dashboard
├── update_dashboard.py     ← 同步腳本（從 Excel 更新 HTML 並 push）
├── CLAUDE.md               ← 本說明檔
└── .gitignore              ← 排除 .xlsm / .xlsx
```

> ⚠️ Excel 檔（Gas usage.xlsm）放在本機 `C:\Users\dk098\Documents\House related\`，不進 git repo

## 常用指令

### 更新 dashboard
```bash
python update_dashboard.py
```
這個指令會：
1. 讀取 .xlsm 的 Daily_Summary sheet（每日用量）
2. 更新 index.html 的 DATA BLOCK 區塊
3. git commit + push 到 GitHub

### 只更新 HTML，不 push
```bash
python update_dashboard.py --dry-run
```

## 實際帳單記錄（KNOWN_BILLS）
無法從 Excel 自動讀取，需從 gas.e-letter.com.tw 歷史帳單手動抄錄後，
更新 `update_dashboard.py` 內的 `KNOWN_BILLS` 清單（deg / rate / base / total）。
最新一筆帳單的單價會被用作所有月份的估算費率。

## 資料來源差異（與電費 app 不同）
- 無雙月帳單，僅有逐日用量記錄（Daily_Summary sheet）
- 瓦斯實際為**月結**（非雙月），基本費固定 $300/月
- 單價隨中油牌價每週調整，目前僅有 4 筆實際帳單可參考

## GitHub Pages 網址
- `https://dksbluesky.github.io/gas-bill-analysis/`
