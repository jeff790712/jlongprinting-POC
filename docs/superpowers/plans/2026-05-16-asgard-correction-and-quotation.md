# Asgard 架構敘事修正 + POC 報價單產出 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 修正 4 份既有 POC 文件中對 Asgard 三大引擎（Odin/Sindri/Mimir）的錯誤歸屬，重新定位為 Phase 1（Odin + 客製化開發 + 基本 DB/AP Server）+ Phase 2（Sindri/Mimir 未來擴展），並產出含公式可填寫的 Excel 報價單。

**Architecture:** 純文件 + 靜態 HTML/CSS 修改，加上一支獨立 Python 腳本（openpyxl）產生 .xlsx。所有變動都是離散的、可獨立驗證的；不涉及 demo 行為邏輯。

**Tech Stack:** Markdown / HTML / CSS（無 build system）；Python 3 + openpyxl 3.1.5；zh-TW 文字內容。

---

## Spec Reference

設計規格：`docs/superpowers/specs/2026-05-16-asgard-correction-and-quotation-design.md`

## File Structure

| File | Status | Responsibility |
|------|--------|----------------|
| `scripts/build_quotation.py` | Create | Python/openpyxl 腳本，產生 Excel 報價單 |
| `POC_報價單_杰隆印刷.xlsx` | Create (by script) | 單一工作表 Excel 報價單，含公式 |
| `POC_Proposal_杰隆印刷_AI接單系統.md` | Modify | § 1 摘要 / § 3 解決方案 / § 5 系統架構 / § 6 核心功能 / § 7 工具映射 / § 10 時程；新增 Phase 2 段 |
| `flow.html` | Modify | Hero 副文、「三大 AI 引擎」section、流程步驟 badge、Legend；末段加 Phase 2 卡片 |
| `index.html` | Verify only | Grep 確認無遺留三引擎描述（已知無，但需 final check） |
| `CLAUDE.md` | Modify | 對「三大引擎」段落的描述更新為 Phase 1 + Phase 2 框架 |

---

## Task 1: 建立 scripts/build_quotation.py 並產生 .xlsx

**Files:**
- Create: `scripts/build_quotation.py`
- Create (via script): `POC_報價單_杰隆印刷.xlsx`

- [ ] **Step 1: Create scripts/ directory**

Run:
```bash
mkdir -p scripts && ls scripts
```
Expected: directory exists, empty.

- [ ] **Step 2: Write the script (full content)**

Create `scripts/build_quotation.py` with the following content. (Path constants resolve relative to the repository root so the script can be invoked from any CWD.)

```python
#!/usr/bin/env python3
"""Generate POC_報價單_杰隆印刷.xlsx for the Jlong Printing POC.

Single-sheet quotation with:
  [A] Asgard 平台訂閱（靜態文字，不入加總）
  [B] 一次性 POC 導入服務 — 三階段，人天 × 單價 自動加總
  [C] Phase 2 進階模組（另行報價，不入加總）
  [D] 加總區 — 一次性導入小計 + 5% 營業稅 + 含稅總價
  [E] 備註 + 簽核欄

Run from anywhere:
  python3 scripts/build_quotation.py
"""

from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

# Brand colours (mirrors index.html / flow.html)
NAVY = "0F3460"
LIGHT_GOLD = "F7EFD3"
SINDRI_PURPLE = "5B21B6"
MIMIR_GREEN = "065F46"
GREY_TEXT = "64748B"
HEADER_ROW_FILL = "EEF2FF"

FONT_NAME = "微軟正黑體"
DAILY_RATE = 20000


def _section_header(ws, row, text):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
    c = ws.cell(row=row, column=1, value=text)
    c.font = Font(name=FONT_NAME, size=12, bold=True, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor=NAVY)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[row].height = 24


def _phase_header(ws, row, text):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
    c = ws.cell(row=row, column=1, value=text)
    c.font = Font(name=FONT_NAME, size=11, bold=True, color=NAVY)
    c.fill = PatternFill("solid", fgColor=LIGHT_GOLD)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[row].height = 22


def _money_format(cell):
    cell.number_format = '"NT$" #,##0'


def _write_item_row(ws, row, name, desc):
    """Write one B-section item: name, desc, 人天 (blank), 單價 (預填), 小計 (公式)."""
    ws.cell(row=row, column=1, value=name).font = Font(name=FONT_NAME, size=10)
    ws.cell(row=row, column=2, value=desc).font = Font(name=FONT_NAME, size=10, color=GREY_TEXT)
    # C: 人天 留空
    ws.cell(row=row, column=3, value=None).alignment = Alignment(horizontal="center")
    # D: 單價 預填
    d = ws.cell(row=row, column=4, value=DAILY_RATE)
    _money_format(d)
    d.alignment = Alignment(horizontal="right")
    # E: 小計 公式
    e = ws.cell(row=row, column=5, value=f"=C{row}*D{row}")
    _money_format(e)
    e.alignment = Alignment(horizontal="right")


def _phase_subtotal(ws, row, label, start_row, end_row):
    ws.cell(row=row, column=1, value=label).font = Font(name=FONT_NAME, size=10, bold=True)
    e = ws.cell(row=row, column=5, value=f"=SUM(E{start_row}:E{end_row})")
    _money_format(e)
    e.font = Font(name=FONT_NAME, size=10, bold=True)
    e.alignment = Alignment(horizontal="right")


def build_workbook(out_path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "POC 報價單"

    # Column widths
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 50
    ws.column_dimensions["C"].width = 10
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 16

    # Row 1 — Title (merged A1:E1)
    ws.merge_cells("A1:E1")
    t = ws["A1"]
    t.value = "Asgard AI × 杰隆印刷　POC 報價單"
    t.font = Font(name=FONT_NAME, size=18, bold=True, color="FFFFFF")
    t.fill = PatternFill("solid", fgColor=NAVY)
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 38

    # Row 2 — Meta
    meta = [
        ("A", f"報價日期：{date.today().strftime('%Y-%m-%d')}"),
        ("B", "有效期：30 日（自報價日起）"),
        ("C", "案件編號：________"),
        ("E", "聯絡人：________"),
    ]
    for col, val in meta:
        c = ws[f"{col}2"]
        c.value = val
        c.font = Font(name=FONT_NAME, size=10, color=GREY_TEXT)

    # ── [A] Asgard 平台訂閱 ──
    _section_header(ws, 4, "[A] Asgard 平台訂閱費（持續性月費，與下方一次性費用分開）")

    ws["A5"] = "方案"
    ws["B5"] = "Basic-Lite（依正式上線使用量可升級至 Regular NT$ 50,000/月 或 Plus NT$ 100,000/月）"
    ws["A6"] = "月費"
    ws["B6"] = "NT$ 20,000 / 月"
    for r in (5, 6):
        ws[f"A{r}"].font = Font(name=FONT_NAME, size=10, bold=True)
        ws[f"B{r}"].font = Font(name=FONT_NAME, size=10)

    ws["A7"] = "付款方式（擇一）"
    ws["A7"].font = Font(name=FONT_NAME, size=10, bold=True)

    ws["A8"] = "  ☐ 信用卡月扣"
    ws["B8"] = "NT$ 20,000 / 月（每月自動扣款）"
    ws["A9"] = "  ☐ 年付開立發票"
    ws["B9"] = "NT$ 240,000 / 年（一次開立，預付一年）"
    for r in (8, 9):
        ws[f"A{r}"].font = Font(name=FONT_NAME, size=10)
        ws[f"B{r}"].font = Font(name=FONT_NAME, size=10, color=GREY_TEXT)

    # ── [B] 一次性 POC 導入服務 ──
    _section_header(ws, 11, "[B] 一次性 POC 導入服務（4-6 週，單價 NT$ 20,000 / 人天 = 8 小時）")

    # Column headers row 12
    headers = [("A", "項目"), ("B", "工作內容"), ("C", "人天"), ("D", "單價"), ("E", "小計")]
    for col, val in headers:
        c = ws[f"{col}12"]
        c.value = val
        c.font = Font(name=FONT_NAME, size=10, bold=True, color=NAVY)
        c.fill = PatternFill("solid", fgColor=HEADER_ROW_FILL)
        c.alignment = Alignment(horizontal="center")
    ws.row_dimensions[12].height = 22

    # Phase A: rows 13 (header), 14-17 (items), 18 (subtotal)
    _phase_header(ws, 13, "──── W1-W2  Phase A：基礎架構 + DB Schema + AP Server 部署 ────")
    phase_a = [
        ("雲端環境規劃與建置", "GCP / AWS 帳號設定、VPC、IAM、Secret 管理"),
        ("DB Schema 設計與建置", "訂單草稿 / 客戶 / 對話紀錄 / 知識庫 schema"),
        ("AP Server 部署", "後端服務容器化、反向代理、TLS 設定"),
        ("CI/CD 與基本監控初始化", "GitHub Actions、Log/Error 監控初始化"),
    ]
    for i, (name, desc) in enumerate(phase_a):
        _write_item_row(ws, 14 + i, name, desc)
    _phase_subtotal(ws, 18, "Phase A 階段小計", 14, 17)

    # Phase B: rows 19 (header), 20-23 (items), 24 (subtotal)
    _phase_header(ws, 19, "──── W3-W4  Phase B：核心對話流 + Odin 工作流建置 ────")
    phase_b = [
        ("Odin 工作流設計", "接單主流程、條件分支、轉接邏輯（於 Odin Studio 內設計）"),
        ("對話 Prompt 與欄位偵測邏輯", "LLM Prompt 工程、欄位缺口偵測、白話追問生成"),
        ("Chatbot 前端 UI 開發", "嵌入式對話框、訊息流、模擬圖預覽元件"),
        ("Session 管理 + 訂單草稿 API", "Session 持久化、訂單 JSON schema、後端 API"),
    ]
    for i, (name, desc) in enumerate(phase_b):
        _write_item_row(ws, 20 + i, name, desc)
    _phase_subtotal(ws, 24, "Phase B 階段小計", 20, 23)

    # Phase C: rows 25 (header), 26-31 (items), 32 (subtotal)
    _phase_header(ws, 25, "──── W5-W6  Phase C：進階功能 + UAT + 教育訓練 ────")
    phase_c = [
        ("圖檔上傳解析", "LLM Vision 串接、設計稿色數與尺寸建議"),
        ("AI 模擬圖生成", "Image Gen API 串接、貼標情境圖輸出"),
        ("舊客識別查詢", "Email / 電話比對、歷史訂單帶入流程"),
        ("情緒偵測 + 真人轉接通知", "情緒分類、Email / LINE Webhook 通知"),
        ("UAT 支援 + 上線部署", "Bug 修正、上線切換、煙霧測試"),
        ("教育訓練與文件交付", "業務操作訓練（1 小時）、技術交接文件"),
    ]
    for i, (name, desc) in enumerate(phase_c):
        _write_item_row(ws, 26 + i, name, desc)
    _phase_subtotal(ws, 32, "Phase C 階段小計", 26, 31)

    # ── [C] Phase 2 ──
    _section_header(ws, 34, "[C] Phase 2 進階模組（另行報價，不計入本次總價）")

    ws["A35"] = "Sindri Agent Hub"
    ws["B35"] = "多 Agent 編排（取代 Phase 1 由 Odin + 客製化程式組成的單一 LLM 對話）"
    ws["E35"] = "另行報價"
    ws["A35"].font = Font(name=FONT_NAME, size=10, bold=True, color=SINDRI_PURPLE)
    ws["B35"].font = Font(name=FONT_NAME, size=10, color=GREY_TEXT)
    ws["E35"].font = Font(name=FONT_NAME, size=10, italic=True, color=GREY_TEXT)
    ws["E35"].alignment = Alignment(horizontal="right")

    ws["A36"] = "Mimir Data Insight"
    ws["B36"] = "數據洞察儀表板（取代 Phase 1 的基本 DB 報表）"
    ws["E36"] = "另行報價"
    ws["A36"].font = Font(name=FONT_NAME, size=10, bold=True, color=MIMIR_GREEN)
    ws["B36"].font = Font(name=FONT_NAME, size=10, color=GREY_TEXT)
    ws["E36"].font = Font(name=FONT_NAME, size=10, italic=True, color=GREY_TEXT)
    ws["E36"].alignment = Alignment(horizontal="right")

    # ── [D] 加總區 ──
    _section_header(ws, 38, "[D] 一次性 POC 導入服務 加總")

    ws["A39"] = "一次性導入服務小計（未稅）"
    ws["E39"] = "=E18+E24+E32"
    ws["A39"].font = Font(name=FONT_NAME, size=11, bold=True)
    ws["E39"].font = Font(name=FONT_NAME, size=11, bold=True)
    _money_format(ws["E39"])
    ws["E39"].alignment = Alignment(horizontal="right")

    ws["A40"] = "營業稅 5%"
    ws["E40"] = "=ROUND(E39*0.05,0)"
    ws["A40"].font = Font(name=FONT_NAME, size=11, bold=True)
    ws["E40"].font = Font(name=FONT_NAME, size=11, bold=True)
    _money_format(ws["E40"])
    ws["E40"].alignment = Alignment(horizontal="right")

    ws["A41"] = "含稅總價"
    ws["E41"] = "=E39+E40"
    ws["A41"].font = Font(name=FONT_NAME, size=13, bold=True, color=NAVY)
    ws["A41"].fill = PatternFill("solid", fgColor=LIGHT_GOLD)
    ws["E41"].font = Font(name=FONT_NAME, size=13, bold=True, color=NAVY)
    ws["E41"].fill = PatternFill("solid", fgColor=LIGHT_GOLD)
    _money_format(ws["E41"])
    ws["E41"].alignment = Alignment(horizontal="right")
    ws.row_dimensions[41].height = 28

    ws.merge_cells("A42:E42")
    ws["A42"] = "★ 本份總價僅為 POC 一次性導入服務費用；Asgard 平台訂閱費另計（見 [A] 區）"
    ws["A42"].font = Font(name=FONT_NAME, size=10, italic=True, color=GREY_TEXT)

    # ── [E] 備註 ──
    _section_header(ws, 44, "[E] 備註")

    notes = [
        "假設與排除：",
        "  1. 不含杰隆印刷既有 ERP / 生產系統的整合與對接",
        "  2. 不含 LINE 官方帳號 / FB Messenger 等通道串接（Phase 2 規劃）",
        "  3. 不含自動報價計算邏輯（需杰隆提供定價模型）",
        "  4. 雲端基礎設施費用（GCP / AWS）由客戶以實際使用量支付",
        "",
        "付款條件建議：簽約 30% / 中期驗收 40% / 最終驗收 30%",
        "報價有效期：自報價日起 30 日",
    ]
    for i, text in enumerate(notes):
        c = ws.cell(row=45 + i, column=1, value=text)
        c.font = Font(name=FONT_NAME, size=10, color=("000000" if text and not text.startswith(" ") else GREY_TEXT))

    # Signature
    ws["A54"] = "客戶簽章：______________________"
    ws["C54"] = "日期：______________"
    ws["A55"] = "Asgard AI：______________________"
    ws["C55"] = "日期：______________"
    for r in (54, 55):
        ws[f"A{r}"].font = Font(name=FONT_NAME, size=10)
        ws[f"C{r}"].font = Font(name=FONT_NAME, size=10)

    # Freeze panes — keep title row visible
    ws.freeze_panes = "A3"

    wb.save(out_path)


def main() -> None:
    repo_root = Path(__file__).resolve().parent.parent
    out = repo_root / "POC_報價單_杰隆印刷.xlsx"
    build_workbook(out)
    print(f"Generated: {out}")


if __name__ == "__main__":
    main()
```

- [ ] **Step 3: Run the script**

Run from repo root:
```bash
python3 scripts/build_quotation.py
```
Expected output:
```
Generated: /Users/williamwang/Documents/executing/cases/Asgard/jlongprinting-POC/POC_報價單_杰隆印刷.xlsx
```

- [ ] **Step 4: Verify formula cells via openpyxl readback**

Run this one-liner from repo root:
```bash
python3 -c "
from openpyxl import load_workbook
wb = load_workbook('POC_報價單_杰隆印刷.xlsx')
ws = wb.active
print('E14 formula:', ws['E14'].value)
print('E18 subtotal:', ws['E18'].value)
print('E24 subtotal:', ws['E24'].value)
print('E32 subtotal:', ws['E32'].value)
print('E39 untaxed total:', ws['E39'].value)
print('E40 tax:', ws['E40'].value)
print('E41 with-tax total:', ws['E41'].value)
print('E35 Sindri price:', ws['E35'].value)
print('E36 Mimir price:', ws['E36'].value)
"
```
Expected:
```
E14 formula: =C14*D14
E18 subtotal: =SUM(E14:E17)
E24 subtotal: =SUM(E20:E23)
E32 subtotal: =SUM(E26:E31)
E39 untaxed total: =E18+E24+E32
E40 tax: =ROUND(E39*0.05,0)
E41 with-tax total: =E39+E40
E35 Sindri price: 另行報價
E36 Mimir price: 另行報價
```

- [ ] **Step 5: Spot-check calculation with sample 人天 values**

Run from repo root:
```bash
python3 -c "
from openpyxl import load_workbook
wb = load_workbook('POC_報價單_杰隆印刷.xlsx')
ws = wb.active
# Write 1 人天 to every B-section item (rows 14-17, 20-23, 26-31)
for r in list(range(14, 18)) + list(range(20, 24)) + list(range(26, 32)):
    ws[f'C{r}'] = 1
wb.save('POC_報價單_杰隆印刷_test.xlsx')
print('Wrote test copy with 人天=1 for all 14 items.')
print('Expected on open: 14 items × NT$20,000 = NT$280,000 未稅, NT$14,000 稅, NT$294,000 含稅')
"
```
Then open `POC_報價單_杰隆印刷_test.xlsx` in Excel / LibreOffice / Numbers and confirm:
- 各小計顯示 NT$ 20,000
- Phase A 小計 NT$ 80,000、Phase B 小計 NT$ 80,000、Phase C 小計 NT$ 120,000
- 未稅 NT$ 280,000、稅 NT$ 14,000、含稅 NT$ 294,000

Then delete the test file:
```bash
rm POC_報價單_杰隆印刷_test.xlsx
```

- [ ] **Step 6: Commit**

```bash
git add scripts/build_quotation.py POC_報價單_杰隆印刷.xlsx
git commit -m "feat: add Excel POC quotation generator and output

scripts/build_quotation.py creates POC_報價單_杰隆印刷.xlsx with:
- [A] Asgard subscription as static text (NT\$20k/月 credit card OR NT\$240k/年 invoice)
- [B] One-time POC implementation, three phases (A/B/C), 人天 × 單價 auto-sum
- [C] Phase 2 (Sindri / Mimir) as future-quote line items
- [D] Totals: untaxed subtotal, 5% VAT, with-tax total — only sums [B]
- [E] Assumptions, payment terms, signature block

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

---

## Task 2: 修正 POC_Proposal § 1 + § 3（執行摘要 + 解決方案）

**Files:**
- Modify: `POC_Proposal_杰隆印刷_AI接單系統.md:32` (執行摘要)
- Modify: `POC_Proposal_杰隆印刷_AI接單系統.md:74-100` (§ 3.1 整體思路 + § 3.2 工具對應表)

- [ ] **Step 1: Read affected section to confirm current state**

Read `POC_Proposal_杰隆印刷_AI接單系統.md` lines 28–105 with the Read tool to confirm the text below still matches.

- [ ] **Step 2: Update § 1 執行摘要 (line 32)**

Find and replace:
```
本提案旨在以 **Asgard AI 平台**（Odin 工作流引擎 + Sindri 多 Agent 執行 + Mimir 資料整合）為技術底層，打造一套 **AI 自動接單 Chatbot**，
```
With:
```
本提案旨在以 **Asgard AI 平台的 Odin 工作流引擎** 為核心，搭配客製化前後端開發與基本 DB / AP Server，打造一套 **AI 自動接單 Chatbot**，
```

- [ ] **Step 3: Update § 3.1 architecture flow (lines 74–86)**

Replace the ASCII block:
```
以 **Asgard AI 平台** 的三層工具，對應業主從「接單」到「訂單成立」的完整流程：

```
客戶白話詢問
     ↓
[Sindri Agent] — 意圖識別 + 舊客比對 + 欄位缺口追蹤
     ↓
[Odin Workflow] — 引導對話 → 圖檔解析 → 模擬圖生成 → 報價試算
     ↓
[Mimir Data] — 客戶資料庫 / 訂單追蹤表 / 知識庫查詢
     ↓
訂單草稿 → 人工確認 → 正式成立
```
```
With:
```
以 **Asgard Odin 工作流引擎** 為核心，搭配客製化前後端與基本 DB，對應業主從「接單」到「訂單成立」的完整流程：

```
客戶白話詢問
     ↓
[Odin Workflow] — 對話編排 + 意圖識別節點 + 欄位缺口追蹤
     ↓
[客製化後端] — 圖檔解析 / 模擬圖生成 / 舊客比對 / 情緒偵測通知
     ↓
[基本 DB] — 訂單草稿 / 客戶資料 / 對話紀錄
     ↓
訂單草稿 → 人工確認 → 正式成立
```
```

- [ ] **Step 4: Update § 3.2 工具對應表 (lines 88–100)**

Replace the table:
```
| 業主需求 | Asgard AI 工具 | 說明 |
|---------|---------------|------|
| AI Chatbot 對話引導 | **Sindri** Multi-Agent | 意圖識別、欄位缺口偵測、話術生成 |
| 零程式碼工作流設計 | **Odin** Workflow | 設計詢價流程、條件分支、轉接邏輯 |
| 訂單資料管理 | **Mimir** Data Hub | 整合 Excel 追蹤表、客戶資料庫 |
| 圖檔上傳解析 | Odin + LLM Vision | 解析設計稿色數、特殊工藝 |
| AI 模擬圖生成 | Odin + Image Gen | 貼標情境模擬圖輸出 |
| 舊客識別與帶入 | Mimir + Sindri | 電話/Email 比對歷史訂單 |
| 情緒偵測轉真人 | Sindri Agent | 負面情緒觸發 Email/LINE 人工接管 |
| 多語言支援 | LLM（Claude/GPT） | 英文詢價自動翻譯處理 |
| LINE / 網頁雙管道 | Odin Webhook | 統一入口接收各管道訊息 |
```
With:
```
| 業主需求 | 本次採用 | 說明 |
|---------|---------------|------|
| 零程式碼工作流設計 | **Odin** Workflow | 設計詢價流程、條件分支、轉接邏輯 |
| AI Chatbot 對話引導 | **Odin 工作流節點 + 客製化程式** | LLM 節點負責意圖識別、客製化程式做欄位缺口偵測與話術生成 |
| 訂單資料管理 | **客製化 DB + AP Server** | 訂單草稿、客戶資料、對話紀錄持久化 |
| 圖檔上傳解析 | Odin + LLM Vision | 解析設計稿色數、特殊工藝 |
| AI 模擬圖生成 | Odin + Image Gen | 貼標情境模擬圖輸出 |
| 舊客識別與帶入 | 客製化後端 + DB | 電話 / Email 比對歷史訂單 |
| 情緒偵測轉真人 | Odin LLM 節點 + 客製化通知 | 負面情緒觸發 Email / LINE 人工接管 |
| 多語言支援 | LLM（Claude / GPT） | 英文詢價自動翻譯處理 |
| LINE / 網頁雙管道 | Odin Webhook | 統一入口接收各管道訊息 |
```

- [ ] **Step 5: Commit**

```bash
git add POC_Proposal_杰隆印刷_AI接單系統.md
git commit -m "docs(proposal): correct Asgard scope in §1 and §3

Rewrite executive summary, architecture flow, and tool mapping
table to reflect actual Phase-1 scope: Odin workflow engine plus
custom front/back-end and basic DB/AP Server. Removes incorrect
attribution of intent detection, field-gap tracking, and customer
data management to Sindri/Mimir (those products are not in the
POC scope).

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

---

## Task 3: 修正 POC_Proposal § 5 + § 6（系統架構 + 核心功能）

**Files:**
- Modify: `POC_Proposal_杰隆印刷_AI接單系統.md:173–195` (§ 5.2 映射表 + § 6.1 引擎歸屬)
- Modify: `POC_Proposal_杰隆印刷_AI接單系統.md:300–360` (§ 6 其他 Sindri/Mimir 提及)

- [ ] **Step 1: Read § 5.2 mapping table and § 6.1 description**

Read `POC_Proposal_杰隆印刷_AI接單系統.md` lines 173–200 with the Read tool to confirm current state.

- [ ] **Step 2: Update § 5.2 Asgard AI 工具映射 (lines 175–181)**

Replace:
```
| 架構層 | Asgard AI 工具 | 功能 |
|--------|--------------|------|
| 使用者層 | React SDK / Web Embed | 嵌入式對話框 |
| 對話層 | Odin Workflow + Router | 意圖路由、Session 管理 |
| AI 核心 | Sindri Multi-Agent | 欄位偵測、知識庫查詢、模擬圖 |
| 資料層 | Mimir Data Hub | 訂單資料、客戶識別 |
| 通知機制 | Odin HTTP Node | Email / LINE Webhook |
```
With:
```
| 架構層 | 本次採用 | 功能 |
|--------|--------------|------|
| 使用者層 | 客製化前端（React） | 嵌入式對話框、AI 模擬圖預覽 |
| 對話層 | Odin Workflow + Router | 意圖路由、Session 管理 |
| AI 核心 | Odin LLM 節點 + 客製化程式 | 欄位偵測、知識庫 RAG、模擬圖串接 |
| 資料層 | 客製化 DB + AP Server | 訂單資料、客戶識別、對話紀錄 |
| 通知機制 | Odin HTTP Node | Email / LINE Webhook |
```

- [ ] **Step 3: Update § 6.1 Field Gap Detection attribution (line 189)**

Replace:
```
這是系統的最核心元件，由 **Sindri Agent** 驅動，負責在每輪對話後：
```
With:
```
這是系統的最核心元件，由 **Odin 工作流節點 + 客製化程式** 驅動，負責在每輪對話後：
```

- [ ] **Step 4: Read § 6 後段 to locate other mentions**

Read lines 295–360 to locate the references at lines 305 (`Mimir 客戶資料庫比對`) and 355 (`寫入 Mimir DB + 通知業務 Email`).

- [ ] **Step 5: Update § 6 後段提及 (line 305 area)**

Find:
```
Mimir 客戶資料庫比對
```
Replace with:
```
客製化 DB 客戶資料比對
```

- [ ] **Step 6: Update § 6 後段提及 (line 355 area)**

Find:
```
[訂單 JSON 輸出] → 寫入 Mimir DB + 通知業務 Email
```
Replace with:
```
[訂單 JSON 輸出] → 寫入客製化 DB + 通知業務 Email
```

- [ ] **Step 7: Commit**

```bash
git add POC_Proposal_杰隆印刷_AI接單系統.md
git commit -m "docs(proposal): correct §5 architecture map and §6 attributions

Reframe AI core and data layer as 'Odin + custom dev' and 'custom
DB + AP Server' respectively. Field Gap Engine is driven by Odin
LLM nodes plus custom code, not by Sindri. Remove 'Mimir DB'
references in §6 in favor of 客製化 DB.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

---

## Task 4: 修正 POC_Proposal § 7（Asgard 工具映射）+ § 10（時程）+ 新增 Phase 2 段

**Files:**
- Modify: `POC_Proposal_杰隆印刷_AI接單系統.md` § 7.1 / § 7.2 / § 7.3 + § 7 表格
- Modify: `POC_Proposal_杰隆印刷_AI接單系統.md` § 10 Week 2 description
- Add: 新章節「Phase 2 進階擴展（未來規劃）」插在 § 7 之後或 § 11 之前

- [ ] **Step 1: Read § 7 in full**

Read `POC_Proposal_杰隆印刷_AI接單系統.md` lines 360–400 to confirm current state of § 7.1 Odin / § 7.2 Sindri / § 7.3 Mimir headers and the 工具負責項目 table.

- [ ] **Step 2: Update § 7 header to focus on Phase 1**

Find the § 7 section header (search for `## 7. Asgard AI 工具對應` or similar). Replace the section title and intro with:
```
## 7. 本次 POC 採用的工具與分工（Phase 1）

本次 POC 採用 **Asgard Odin 工作流引擎** 為核心，搭配客製化前後端與基本 DB / AP Server。Asgard 平台的其他模組（Sindri Agent Hub、Mimir Data Insight）為 Phase 2 進階擴展，本次不導入。
```

- [ ] **Step 3: Rewrite § 7.2 Sindri subsection**

Find `### 7.2 Sindri（Multi-Agent 執行引擎）` and replace the entire subsection (header + body until § 7.3) with:
```
### 7.2 客製化開發（前後端 + AI 整合層）

本次 POC 由 Asgard AI 提供客製化開發資源，承擔以下原本可由 Sindri 提供的功能：

- **意圖識別 / 欄位缺口偵測 / 話術生成**：以 Odin LLM 節點 + 客製化 Python/Node 程式實作
- **舊客識別與帶入**：客製化 API 查詢基本 DB，將歷史訂單帶回對話流
- **情緒偵測 + 真人轉接**：Odin LLM 節點分類 + 客製化通知（Email / LINE Webhook）
- **AI 模擬圖串接**：Odin Image Gen 節點 + 客製化前端展示
- **嵌入式 Chatbot 前端**：React 元件，可內嵌至杰隆官網

Phase 2 階段可升級為 **Sindri Agent Hub** 的多 Agent 編排，獲得更強的對話狀態管理與 Agent 協作能力。
```

- [ ] **Step 4: Rewrite § 7.3 Mimir subsection**

Find `### 7.3 Mimir（資料整合）` and replace the entire subsection with:
```
### 7.3 基本 DB + AP Server（資料層）

本次 POC 由 Asgard AI 協助搭建基本資料層，承擔以下原本可由 Mimir 提供的功能：

- **訂單草稿 DB**：結構化儲存對話完成後的訂單需求
- **客戶資料 DB**：Email / 電話索引，支援舊客查詢
- **對話紀錄 DB**：完整對話 log，供後續分析與訓練
- **AP Server**：後端 API、Session 管理、容器化部署、基本監控

Phase 2 階段可升級為 **Mimir Data Insight**，獲得儀表板、跨來源資料整合與決策建議能力。
```

- [ ] **Step 5: Update § 7 工具負責項目表 (lines 385–392 area)**

Read the table around lines 385–392 first. Then replace any rows that reference Mimir / Sindri with 客製化 DB / 客製化後端 equivalents. Specifically:

Find:
```
| 客戶資料庫 | Mimir 建立 | 電話/Email 識別比對 |
```
Replace:
```
| 客戶資料庫 | 客製化 DB | 電話 / Email 識別比對 |
```

Find:
```
| 對話紀錄 | 寫入 Mimir | 分析與訓練用途 |
```
Replace:
```
| 對話紀錄 | 寫入客製化 DB | 分析與訓練用途 |
```

Find:
```
| 訂單草稿 | Mimir DB 寫入 | 業務後台查看 |
```
Replace:
```
| 訂單草稿 | 客製化 DB 寫入 | 業務後台查看 |
```

- [ ] **Step 6: Add new "Phase 2 進階擴展" subsection**

After § 7.3 (end of § 7), insert a new subsection:
```

### 7.4 Phase 2 進階擴展（本次 POC 不含，列為未來規劃）

POC 上線並穩定後，可逐步導入 Asgard 平台的進階模組：

| 模組 | 升級價值 | 觸發時機 |
|------|---------|---------|
| **Sindri Agent Hub** | 多 Agent 並行協作、跨對話記憶、Agent 間任務交接，提升複雜詢價案件的對話品質 | 散客量 > 50/日、需處理多步驟詢價時 |
| **Mimir Data Insight** | 訂單轉換漏斗、客群分析、材質偏好趨勢儀表板，輔助業務決策 | 累積 3 個月以上歷史訂單後 |

Phase 2 採另行報價。
```

- [ ] **Step 7: Update § 10 timeline Week 2 (line 464)**

Read `POC_Proposal_杰隆印刷_AI接單系統.md` lines 460–470 first. Then find:
```
| **Week 2** | Sindri Agent 開發、欄位引導引擎 | 完整詢價對話可運行、舊客識別功能 |
```
Replace with:
```
| **Week 2** | Odin 工作流 + 客製化欄位引導引擎開發 | 完整詢價對話可運行、舊客識別功能 |
```

- [ ] **Step 8: Verify no remaining Sindri/Mimir misattribution in proposal**

Run:
```bash
grep -n -i "sindri\|mimir" POC_Proposal_杰隆印刷_AI接單系統.md
```
Expected: only matches inside the new § 7.4 Phase 2 section (the table mentioning Sindri Agent Hub / Mimir Data Insight as future expansion). No matches in §1, §3, §5, §6, §7.1/.2/.3, §10.

- [ ] **Step 9: Commit**

```bash
git add POC_Proposal_杰隆印刷_AI接單系統.md
git commit -m "docs(proposal): rewrite §7 as Phase 1 + add §7.4 Phase 2 expansion

§7 now describes the Phase-1 stack (Odin workflow + custom dev +
basic DB/AP Server) instead of the three-engine pitch. The
previous §7.2 'Sindri' and §7.3 'Mimir' subsections are rewritten
as '客製化開發' and '基本 DB + AP Server' respectively, listing
the responsibilities each subsystem covers. A new §7.4 Phase 2
subsection lists Sindri Agent Hub and Mimir Data Insight as
future upgrades with separate quote. §10 Week-2 deliverable
relabelled.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

---

## Task 5: 修正 flow.html（Hero + 三大引擎 → 兩柱 + 步驟 badge + Phase 2 卡片）

**Files:**
- Modify: `flow.html:356` (hero subtitle)
- Modify: `flow.html:360–405` (三大引擎 section)
- Modify: `flow.html:408–632` (流程步驟區的 step-badge)
- Modify: `flow.html:560–600` (legend section)
- Add: 新「Phase 2」卡片區塊（在效益區之前或之後）

- [ ] **Step 1: Read flow.html in chunks to confirm current state**

Read `flow.html` lines 340–410 (header + hero + three-engine section).

- [ ] **Step 2: Update hero subtitle (line 356)**

Find:
```html
    <p>以 Asgard AI 三大引擎為核心，將散客詢價流程從「0.5～3 天的來回溝通」壓縮至即時完成，並向後整合訂單追蹤系統。</p>
```
Replace with:
```html
    <p>以 Asgard AI 的 <strong>Odin 工作流引擎</strong> 為核心，搭配客製化前後端與基本資料層，將散客詢價流程從「0.5～3 天的來回溝通」壓縮至即時完成。</p>
```

- [ ] **Step 3: Rewrite 三大引擎 section (lines 360–405)**

Read lines 360–410 first to see the exact existing markup. The section currently has three `<div class="tool-card odin/sindri/mimir">` cards. Replace the entire `<section class="section">` block (from `<div class="section-label">核心技術</div>` to the closing `</section>`) with a two-card layout (`tool-card odin` + new `tool-card custom`):

```html
  <section class="section">
    <div class="section-label">核心技術</div>
    <div class="section-title">Odin 工作流 ＋ 客製化交付</div>
    <div class="section-desc">本次 POC 以 Asgard Odin 工作流引擎為編排核心，搭配客製化前後端與基本 DB / AP Server 完成交付。Sindri Agent Hub、Mimir Data Insight 等進階模組為 Phase 2 未來擴展。</div>

    <div class="tool-grid two-col">
      <div class="tool-card odin">
        <div class="tool-tag">Asgard 平台</div>
        <div class="tool-icon">⚙️</div>
        <div class="tool-name">Odin</div>
        <div class="tool-sub">Asgard Studio · 零程式碼 AI 工作流</div>
        <ul class="tool-points">
          <li>對話流程編排（接單主流程、條件分支、轉接邏輯）</li>
          <li>LLM 節點（意圖識別、欄位缺口偵測、話術生成）</li>
          <li>串接 Vision / Image Gen API（圖檔解析、模擬圖）</li>
          <li>Webhook 入口（網頁、LINE 雙管道）</li>
        </ul>
      </div>

      <div class="tool-card custom">
        <div class="tool-tag">客製化開發</div>
        <div class="tool-icon">🛠️</div>
        <div class="tool-name">前後端 + DB / AP Server</div>
        <div class="tool-sub">由 Asgard AI 團隊客製化交付</div>
        <ul class="tool-points">
          <li>嵌入式 Chatbot 前端（React，可內嵌至官網）</li>
          <li>後端 API（Session、訂單草稿、舊客查詢、通知）</li>
          <li>基本 DB（訂單草稿、客戶資料、對話紀錄）</li>
          <li>AP Server 部署 + 基本監控</li>
        </ul>
      </div>
    </div>
  </section>
```

You will also need to add CSS for `.tool-card.custom` and `.tool-grid.two-col`, plus `.tool-sub` and `.tool-points`. Add this CSS in the `<style>` block, near the existing `.tool-card.odin/sindri/mimir` rules:

```css
    /* Custom-dev card (Phase 1, alongside Odin) */
    :root { --custom: #475569; --custom-t: rgba(71,85,105,.12); }
    .tool-card.custom::before { background: var(--custom); }
    .custom .tool-icon { background: var(--custom-t); }
    .custom .tool-name { color: var(--custom); }
    .custom .tool-tag  { background: var(--custom-t); color: var(--custom); }
    .badge-custom { background: var(--custom-t); color: var(--custom); }

    /* Two-column grid for the Phase-1 cards */
    .tool-grid.two-col { display: grid; grid-template-columns: repeat(2, 1fr); gap: 20px; }
    @media (max-width: 720px) { .tool-grid.two-col { grid-template-columns: 1fr; } }

    .tool-sub { font-size: 12px; color: var(--muted); margin-top: 2px; margin-bottom: 12px; }
    .tool-points { list-style: none; padding: 0; margin: 0; }
    .tool-points li { font-size: 13px; color: var(--text); margin-top: 6px; padding-left: 14px; position: relative; line-height: 1.5; }
    .tool-points li::before { content: "•"; position: absolute; left: 0; color: var(--gold); }
```

- [ ] **Step 4: Remap step badges (lines ~420–540)**

Read lines 415–545 to see all `<span class="step-badge badge-sindri">…</span>` and `<span class="step-badge badge-mimir">…</span>` blocks.

Apply the following replacements (use `replace_all` on each `old_string` since exact strings may appear multiple times):

| Find | Replace |
|------|---------|
| `<span class="step-badge badge-sindri">Sindri · Session 建立</span>` | `<span class="step-badge badge-custom">客製化後端 · Session 建立</span>` |
| `<span class="step-badge badge-sindri">Sindri · Intake Agent</span>` | `<span class="step-badge badge-odin">Odin · Intake 節點</span>` |
| `<span class="step-badge badge-sindri">Sindri · Field Collector Agent</span>` | `<span class="step-badge badge-odin">Odin · 欄位偵測節點</span>` |
| `<span class="step-badge badge-mimir">Mimir · 知識庫 RAG</span>` | `<span class="step-badge badge-custom">客製化 DB · 材質知識</span>` |
| `<span class="step-badge badge-sindri">Sindri · Material Advisor Agent</span>` | `<span class="step-badge badge-odin">Odin · 材質推薦節點</span>` |
| `<span class="step-badge badge-sindri">Sindri · Mockup Generator</span>` | `<span class="step-badge badge-odin">Odin · Image Gen 節點</span>` |
| `<span class="step-badge badge-sindri">Sindri · Emotion Detector</span>` | `<span class="step-badge badge-odin">Odin · 情緒偵測節點</span>` |
| `<span class="step-badge badge-sindri">Sindri · Order Summary Agent</span>` | `<span class="step-badge badge-odin">Odin · 訂單總結節點</span>` |
| `<span class="step-badge badge-mimir">Mimir · 訂單 DB 寫入</span>` | `<span class="step-badge badge-custom">客製化 DB · 訂單寫入</span>` |

Also update the `ft-circle` classes around the same step lines. Find each `<div class="ft-circle c-sindri">…</div>` and `<div class="ft-circle c-mimir">…</div>` (around lines 461, 475, 502, 516, 529) and change `c-sindri` → `c-odin` and `c-mimir` → `c-custom`.

Add the missing `c-custom` rule in the `<style>` block near the existing ft-circle rules:
```css
    .ft-circle.c-custom { background: var(--custom); color: #fff; }
```

- [ ] **Step 5: Update flow step description text where it mentions Mimir**

Find:
```
<div class="ft-desc">訂單草稿自動寫入 Mimir 資料庫，同步觸發 Email / LINE 通知業務人員。
```
Replace with:
```
<div class="ft-desc">訂單草稿自動寫入客製化 DB，同步觸發 Email / LINE 通知業務人員。
```

- [ ] **Step 6: Update legend section (lines 569–589)**

Replace the entire Sindri legend block with the 客製化 entry. Old string (12 lines, exact whitespace as in the file):
```html
            <div class="legend-item">
              <div class="legend-dot dot-sindri"></div>
              <div class="legend-item-text">
                <strong>Sindri — Agent 執行層</strong>
                <span>AI 對話、欄位收集、智能判斷</span>
              </div>
            </div>
            <div class="legend-item">
              <div class="legend-dot dot-mimir"></div>
              <div class="legend-item-text">
                <strong>Mimir — 資料整合層</strong>
                <span>資料庫存取、知識庫查詢</span>
              </div>
            </div>
```
New string:
```html
            <div class="legend-item">
              <div class="legend-dot dot-custom"></div>
              <div class="legend-item-text">
                <strong>客製化開發 — 前後端 + AP Server</strong>
                <span>Chatbot 前端、Session 與訂單 API、舊客查詢、情緒通知</span>
              </div>
            </div>
            <div class="legend-item">
              <div class="legend-dot dot-custom-db"></div>
              <div class="legend-item-text">
                <strong>基本 DB — 資料層</strong>
                <span>訂單草稿、客戶資料、對話紀錄；Phase 2 可升級為 Mimir Data Insight</span>
              </div>
            </div>
```

Add the new dot colours in the `<style>` block, near the existing `.dot-odin/.dot-sindri/.dot-mimir` rules (around line 255):
```css
    .dot-custom    { background: var(--custom); }
    .dot-custom-db { background: var(--gold); }
```

- [ ] **Step 7: Add Phase 2 future-expansion card at the end of the page**

Find the existing CTA / footer block at the bottom of the page (after the 導入效益 section, before the final `</body>` or footer-links). Insert a new section just before the closing wrapper:

```html
  <section class="section">
    <div class="section-label">未來擴展</div>
    <div class="section-title">Phase 2：Asgard 平台進階模組</div>
    <div class="section-desc">POC 上線並穩定後，可逐步導入 Asgard 平台的進階模組，獲得更強的對話品質與決策能力。另行報價。</div>

    <div class="tool-grid two-col">
      <div class="tool-card sindri">
        <div class="tool-tag">Phase 2</div>
        <div class="tool-icon">🤖</div>
        <div class="tool-name">Sindri Agent Hub</div>
        <div class="tool-sub">Asgard Agents Hub · 多 Agent 自動化執行</div>
        <ul class="tool-points">
          <li>多 Agent 並行協作、Agent 間任務交接</li>
          <li>跨對話記憶、長期客戶關係追蹤</li>
          <li>複雜詢價案件的對話品質升級</li>
        </ul>
      </div>

      <div class="tool-card mimir">
        <div class="tool-tag">Phase 2</div>
        <div class="tool-icon">📊</div>
        <div class="tool-name">Mimir Data Insight</div>
        <div class="tool-sub">Asgard Data Insight · 資料洞察與決策</div>
        <ul class="tool-points">
          <li>訂單轉換漏斗、客群分析儀表板</li>
          <li>材質偏好趨勢、季節性需求預測</li>
          <li>跨來源資料整合與決策建議</li>
        </ul>
      </div>
    </div>
  </section>
```

- [ ] **Step 8: Open flow.html in a browser and visually verify**

Run a local server:
```bash
python3 -m http.server 8000 &
```
Open http://localhost:8000/flow.html and check:
- Hero subtitle now mentions Odin + 客製化
- 三大引擎 section is replaced by two cards (Odin blue + 客製化 slate)
- Step badges in 流程 section show Odin blue or custom slate (no purple/teal in Phase-1 area)
- Legend has three rows: Odin / 客製化開發 / 基本 DB
- New Phase 2 section appears near the end with two purple/teal cards

Kill the server:
```bash
kill %1
```

- [ ] **Step 9: Final grep verify**

```bash
grep -c -i "sindri\|mimir" flow.html
```
Expected: a small number (~6) of matches, **all** inside the new Phase 2 section or CSS variable definitions. Visually inspect each match to confirm none describe Phase-1 functions.

- [ ] **Step 10: Commit**

```bash
git add flow.html
git commit -m "feat(flow): reframe as Phase 1 (Odin + custom dev) plus Phase 2 card

- Hero subtitle replaces 三大引擎 framing with Odin + 客製化交付
- 三大 AI 引擎 section replaced by two-card grid: Odin and 客製化
- Step badges remap Sindri/Mimir mentions to Odin · [node] or
  客製化後端 · [function] / 客製化 DB · [function]
- Legend updated to three rows: Odin / 客製化開發 / 基本 DB
- New Phase 2 section at page end shows Sindri Agent Hub and
  Mimir Data Insight as future upgrades (kept original purple
  and teal brand colours for visual continuity)

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

---

## Task 6: 驗證 index.html（預期無變動）

**Files:**
- Verify only: `index.html`

- [ ] **Step 1: Grep for engine references**

```bash
grep -n -i "sindri\|mimir\|三大引擎\|三大 AI" index.html
```
Expected: zero matches.

- [ ] **Step 2: If matches found, apply fix per spec; else mark task complete**

If grep returns any matches, read those lines, apply Edit to replace per the same conventions as Task 2–4. If zero matches, this task is complete with no commit needed — proceed to Task 7.

---

## Task 7: 更新 CLAUDE.md 框架敘述

**Files:**
- Modify: `CLAUDE.md` — "Editing the architecture page (flow.html)" 段落

- [ ] **Step 1: Read CLAUDE.md current state**

Read `CLAUDE.md` in full to locate the section that describes the three engines (currently includes lines like `- Odin (workflow) — \`--odin: #1e3a8a\` (blue)`).

- [ ] **Step 2: Update the three-engines paragraph**

Find:
```
## Editing the architecture page (`flow.html`)

Pure static HTML/CSS, no JS state. The three engines have fixed brand colors defined as CSS variables and must stay consistent across both files:

- Odin (workflow) — `--odin: #1e3a8a` (blue)
- Sindri (multi-agent) — `--sindri: #5b21b6` (purple)
- Mimir (data hub) — `--mimir: #065F46` (teal)

The shared brand palette (`--navy`, `--gold`, etc.) is duplicated between `index.html` and `flow.html`; if you change brand colors, update both.
```
Replace with:
```
## Editing the architecture page (`flow.html`)

Pure static HTML/CSS, no JS state. The page is structured as two phases:

- **Phase 1 (本次 POC)** — Odin workflow engine (Asgard 平台) + 客製化前後端 + 基本 DB / AP Server. The hero, primary tool cards, and flow-step badges should attribute work to **Odin** or **客製化** (slate `--custom: #475569`), not to Sindri/Mimir.
- **Phase 2 (未來擴展)** — Sindri Agent Hub and Mimir Data Insight as future upgrades, shown in the dedicated Phase-2 section near the page end.

CSS variables (kept consistent with `index.html`):
- `--odin: #1e3a8a` (Phase 1, primary)
- `--custom: #475569` (Phase 1, custom dev)
- `--sindri: #5b21b6` (Phase 2 card only)
- `--mimir: #065F46` (Phase 2 card only)

The shared brand palette (`--navy`, `--gold`, etc.) is duplicated between `index.html` and `flow.html`; if you change brand colors, update both.

If you need to add or modify the POC quotation, edit `scripts/build_quotation.py` and re-run it to regenerate `POC_報價單_杰隆印刷.xlsx`; don't hand-edit the .xlsx unless the change is a one-off and won't survive the next script run.
```

- [ ] **Step 3: Commit**

```bash
git add CLAUDE.md
git commit -m "docs(CLAUDE.md): reframe flow.html guidance for Phase 1/2 split

Explain to future Claude instances that flow.html now shows
Phase 1 (Odin + 客製化) primary and Phase 2 (Sindri/Mimir) as
future expansion, and that the POC quotation .xlsx is generated
by scripts/build_quotation.py.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

---

## Task 8: 最終跨檔驗證

**Files:** All four content files

- [ ] **Step 1: Grep across all content files for misattribution**

```bash
grep -n -i "sindri\|mimir" \
  POC_Proposal_杰隆印刷_AI接單系統.md \
  flow.html \
  index.html \
  CLAUDE.md
```
Expected:
- `POC_Proposal_杰隆印刷_AI接單系統.md`: only matches inside the new § 7.4 Phase 2 subsection
- `flow.html`: matches inside CSS variable definitions and the new Phase 2 section only
- `index.html`: zero matches
- `CLAUDE.md`: only inside the explicit Phase 2 listing

If anything else turns up, fix it and recommit.

- [ ] **Step 2: Verify xlsx is still valid**

```bash
python3 -c "
from openpyxl import load_workbook
wb = load_workbook('POC_報價單_杰隆印刷.xlsx')
ws = wb.active
assert ws['E14'].value == '=C14*D14'
assert ws['E39'].value == '=E18+E24+E32'
assert ws['E41'].value == '=E39+E40'
assert ws['E35'].value == '另行報價'
print('xlsx structure OK')
"
```
Expected: `xlsx structure OK`

- [ ] **Step 3: Confirm git status is clean**

```bash
git status --short
```
Expected: empty (no uncommitted changes).

- [ ] **Step 4: Show final git log**

```bash
git log --oneline -10
```
Expected: see the 7+ commits from this implementation plan plus the prior spec commit.

---

## Self-Review Checklist (run after implementation)

- [ ] All four content files (proposal / flow.html / index.html / CLAUDE.md) grep clean for Sindri/Mimir outside Phase-2 contexts
- [ ] `POC_報價單_杰隆印刷.xlsx` opens in Excel/LibreOffice/Numbers, formulas evaluate correctly when 人天 is filled
- [ ] flow.html visual: Phase-1 area uses only Odin blue + slate; Phase-2 area uses purple + teal
- [ ] No new dependencies introduced beyond Python 3 stdlib + openpyxl
- [ ] zh-TW maintained throughout; no English / Simplified Chinese mixed in
- [ ] Spec § 7 驗收標準 all checked
