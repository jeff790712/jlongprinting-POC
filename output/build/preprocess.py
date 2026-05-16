#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Pre-process proposal markdown for pandoc → xelatex → PDF.

- Replace 5 ASCII chart blocks with references to rendered Mermaid PDFs.
- Append Appendix A (demo screenshots) and Appendix B (quotation).
- Prepend YAML frontmatter with CJK font + pandoc options.
"""

from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent.parent
PROPOSAL = ROOT / "POC_Proposal_杰隆印刷_AI接單系統.md"
BUILD_MD = ROOT / "output" / "build" / "proposal.md"

src = PROPOSAL.read_text(encoding="utf-8")

# Glyph substitutions for chars PingFang TC doesn't carry
GLYPH_FIX = {
    "⚠️": "★",          # warning emoji → star prefix
    "⚠": "★",
    "☐": "□",          # ballot box → white square (PingFang TC has this)
    " ": " ",     # EN QUAD space → regular space
}
for old, new in GLYPH_FIX.items():
    src = src.replace(old, new)

# ---- ASCII → image swaps -----------------------------------------------------

REPLACEMENTS = [
    # § 2.2 pain flow
    (
        "```\n散客進線\n   ↓\n等待業務回應（可能超過 3 天）\n   ↓\n多次來回確認規格（術語不熟、資訊不完整）\n   ↓\n客戶等不及 → 流失訂單\n```",
        "![**圖 1：散客詢價痛點流**](../diagrams/01-pain-flow.pdf){ width=55% }",
    ),
    # § 3.1 solution flow
    (
        "```\n客戶白話詢問\n     ↓\n[Odin Workflow] — 對話編排 + 意圖識別節點 + 欄位缺口追蹤\n     ↓\n[客製化後端] — 圖檔解析 / 模擬圖生成 / 舊客比對 / 情緒偵測通知\n     ↓\n[基本 DB] — 訂單草稿 / 客戶資料 / 對話紀錄\n     ↓\n訂單草稿 → 人工確認 → 正式成立\n```",
        "![**圖 2：POC 解決方案流（Phase 1）**](../diagrams/02-solution-flow.pdf){ width=80% }",
    ),
    # § 5.1 architecture layers (multi-line ASCII box drawing)
    (
        "```\n┌─────────────────────────────────────────────────────────┐\n│                     使用者層 (User Layer)                 │\n│  ┌─────────────┐  ┌──────────────┐  ┌──────────────┐    │\n│  │ 網頁對話框   │  │ LINE 頻道     │  │ 管理後台      │    │\n│  │ (POC 主要)  │  │ (Phase 2)    │  │ (數據/接管)   │    │\n│  └─────────────┘  └──────────────┘  └──────────────┘    │\n└──────────────────────────┬──────────────────────────────┘\n                           │\n┌──────────────────────────▼──────────────────────────────┐\n│                    對話層 (Conversation Layer)            │\n│                                                          │\n│  路由器（統一入口）→ Session 管理 → 意圖識別              │\n│       ↓                                                  │\n│  意圖分類：[詢價] [訂單追蹤] [舊客補印] [閒聊] [轉接]    │\n│       ↓                                                  │\n│  轉接處理：情緒偵測 → Email 通知真人接管                  │\n└──────────────────────────┬──────────────────────────────┘\n                           │\n┌──────────────────────────▼──────────────────────────────┐\n│                    AI 核心 (AI Core)                      │\n│                                                          │\n│  ┌─────────────────────────────────────────────────┐    │\n│  │  ★ 訂單資料引導引擎 (Field Gap Detection Engine)  │    │\n│  │  · 已收集欄位追蹤                                  │    │\n│  │  · 缺口清單即時更新                               │    │\n│  │  · 白話追問生成                                   │    │\n│  └─────────────────────────────────────────────────┘    │\n│                                                          │\n│  LLM 語言模型 | 知識庫 RAG | Prompt 引擎 | 圖像分析      │\n└──────────────────────────┬──────────────────────────────┘\n                           │\n┌──────────────────────────▼──────────────────────────────┐\n│                     資料層 (Data Layer)                   │\n│                                                          │\n│  訂單草稿 DB │ 客戶資料庫 │ 知識庫 │ 對話紀錄             │\n│  （欄位映射）│（舊客比對）│（FAQ+規格）│（訓練用）         │\n└─────────────────────────────────────────────────────────┘\n```",
        "![**圖 3：四層架構**](../diagrams/03-architecture-layers.pdf){ width=92% }",
    ),
    # § 7.1 Odin workflow
    (
        "```\n[Webhook 入口] → [Session 建立 / 舊客比對]\n      ↓\n[LLM Router] → 意圖分類（詢價 / 追蹤 / 補印 / 閒聊）\n      ↓\n[Odin LLM 節點] → 欄位缺口偵測 → 白話追問生成\n      ↓\n[圖檔上傳節點] → [LLM Vision 解析] → 回傳色數建議\n      ↓\n[Image Gen 節點] → 貼標模擬圖生成\n      ↓\n[條件分支] → 完成 / 觸發轉接 / 繼續追問\n      ↓\n[訂單 JSON 輸出] → 寫入客製化 DB + 通知業務 Email\n```",
        "![**圖 4：Odin 工作流節點配置**](../diagrams/04-odin-workflow.pdf){ width=82% }",
    ),
    # § 10 timeline ASCII milestones
    (
        "```\nWeek 1 ─────────────────────── Week 2 ─────────────────────── Week 3\n   │                               │                               │\n   ▼                               ▼                               ▼\n知識庫 Ready              欄位引導引擎完成              POC 上線 Demo\n材質對照表建立             舊客識別功能                  全流程驗收\nOdin 工作流草稿            情緒偵測觸發                  業務後台查看\n                           圖檔解析功能\n```",
        "![**圖 5：POC 導入時程 Gantt（4-6 週版本）**](../diagrams/05-timeline.pdf){ width=92% }",
    ),
]

for old, new in REPLACEMENTS:
    if old not in src:
        raise SystemExit(f"FAIL: could not locate ASCII block starting with:\n{old[:100]}…")
    src = src.replace(old, new, 1)

# ---- Appendix A: demo screenshots -------------------------------------------

APPENDIX_A = """

\\newpage

## 附錄 A：Demo 操作截圖

下列截圖摘錄自實際運行的 POC Demo（`index.html`），呈現一次完整的「保養品瓶標」詢價流程。

![**A.1　歡迎畫面**：AI 詢價助理啟動，提供「第一次詢價 / 補印舊款式」兩種入口，右側為訂單欄位追蹤面板（0 / 10）。](../screenshots/01-welcome.png){ width=92% }

![**A.2　對話進行中**：客戶輸入品名後，AI 接續詢問尺寸；建議快選按鈕降低輸入負擔。](../screenshots/02-mid-conversation.png){ width=92% }

![**A.3　材質智能引導**：8 種官網提供之材質卡片排列，含特性說明與標籤（最熱門 / 質感首選 等），客戶可一鍵選擇。](../screenshots/03-material-grid.png){ width=92% }

![**A.4　AI 模擬圖生成**：依客戶選擇的材質 / 尺寸即時渲染貼標於容器之情境模擬圖，協助客戶確認設計效果。](../screenshots/04-ai-mockup.png){ width=92% }

![**A.5　訂單確認單**：對話完成後輸出結構化需求確認，業務人員可直接使用此資訊提供報價。](../screenshots/05-summary-card.png){ width=92% }
"""

# ---- Appendix B: quotation ---------------------------------------------------

APPENDIX_B = """

\\newpage

## 附錄 B：POC 報價單明細

### B.1 Asgard 平台訂閱費（持續性月費，與下方一次性費用分開計算）

| 方案 | 月費 | 規格 |
|------|------|------|
| **Basic-Lite**（本次推薦） | NT$ 20,000 / 月 | 1 Agent Hub user、1 Data Insight user、1,000 Asgard Credits |

**付款方式（擇一）：**

| 選項 | 金額 | 備註 |
|------|------|------|
| □ 信用卡月扣 | NT$ 20,000 / 月 | 每月自動扣款 |
| □ 年付開立發票 | NT$ 240,000 / 年 | 一次開立，預付一年 |

依正式上線使用量可升級至 Regular（NT$ 50,000/月）或 Plus（NT$ 100,000/月）。

### B.2 一次性 POC 導入服務（4-6 週，單價 NT$ 20,000 / 人天）

#### W1-W2　Phase A：基礎架構 + DB Schema + AP Server 部署

| 項目 | 工作內容 | 人天 | 小計 |
|------|---------|:---:|---:|
| 雲端環境規劃與建置 | GCP / AWS 帳號、VPC、IAM、Secret 管理 | 0.5 | NT$ 10,000 |
| DB Schema 設計與建置 | 訂單草稿 / 客戶 / 對話紀錄 / 知識庫 | 1 | NT$ 20,000 |
| AP Server 部署 | 容器化、反向代理、TLS 設定 | 1 | NT$ 20,000 |
| CI/CD 與基本監控初始化 | GitHub Actions、Log/Error 監控 | 0.5 | NT$ 10,000 |
| **Phase A 階段小計** | | **3** | **NT$ 60,000** |

#### W3-W4　Phase B：核心對話流 + Odin 工作流建置

| 項目 | 工作內容 | 人天 | 小計 |
|------|---------|:---:|---:|
| Odin 工作流設計 | 接單主流程、條件分支、轉接邏輯 | 1 | NT$ 20,000 |
| 對話 Prompt 與欄位偵測邏輯 | LLM Prompt、缺口偵測、白話追問 | 1 | NT$ 20,000 |
| Chatbot 前端 UI 開發 | 嵌入式對話框、訊息流、模擬圖預覽 | 1 | NT$ 20,000 |
| Session 管理 + 訂單草稿 API | Session、訂單 JSON schema、後端 API | 1 | NT$ 20,000 |
| **Phase B 階段小計** | | **4** | **NT$ 80,000** |

#### W5-W6　Phase C：進階功能 + UAT + 教育訓練

| 項目 | 工作內容 | 人天 | 小計 |
|------|---------|:---:|---:|
| 圖檔上傳解析 | LLM Vision 串接、色數與尺寸建議 | 0.5 | NT$ 10,000 |
| AI 模擬圖生成 | Image Gen API 串接、情境圖輸出 | 0.5 | NT$ 10,000 |
| 舊客識別查詢 | Email / 電話比對、歷史訂單帶入 | 0.5 | NT$ 10,000 |
| 情緒偵測 + 真人轉接通知 | 情緒分類、Email / LINE Webhook | 0.5 | NT$ 10,000 |
| UAT 支援 + 上線部署 | Bug 修正、上線切換、煙霧測試 | 0.5 | NT$ 10,000 |
| 教育訓練與文件交付 | 業務操作訓練、技術交接文件 | 0.5 | NT$ 10,000 |
| **Phase C 階段小計** | | **3** | **NT$ 60,000** |

### B.3 Phase 2 進階模組（另行報價，不計入本次總價）

| 模組 | 升級價值 | 費用 |
|------|---------|------|
| Sindri Agent Hub | 多 Agent 並行協作、跨對話記憶、複雜詢價對話品質升級 | 另行報價 |
| Mimir Data Insight | 訂單轉換漏斗、客群分析、材質偏好趨勢儀表板 | 另行報價 |

### B.4 一次性 POC 導入服務 加總

| 項目 | 金額 |
|------|---:|
| 一次性導入服務小計（未稅） | NT$ 200,000 |
| 營業稅 5% | NT$ 10,000 |
| **含稅總價** | **NT$ 210,000** |

> ★ 本份總價僅為 POC 一次性導入服務費用；Asgard 平台訂閱費另計（見 B.1 區）。

### B.5 假設與排除

1. 不含杰隆印刷既有 ERP / 生產系統的整合與對接
2. 不含 LINE 官方帳號 / FB Messenger 等通道串接（Phase 2 規劃）
3. 不含自動報價計算邏輯（需杰隆提供定價模型）
4. 雲端基礎設施費用（GCP / AWS）由客戶以實際使用量支付

### B.6 付款條件與有效期

- **付款條件建議**：簽約 30% / 中期驗收 40% / 最終驗收 30%
- **報價有效期**：自報價日起 30 日
- **本次報價日期**：{today}

---

| 客戶簽章 | 日期 | Asgard AI 簽章 | 日期 |
|----------|------|----------------|------|
| ____________________ | __________ | ____________________ | __________ |
""".format(today=date.today().strftime("%Y-%m-%d"))

# ---- YAML frontmatter --------------------------------------------------------

FRONTMATTER = """---
title: |
  | 杰隆印刷 × Asgard AI
  | AI 自動接單系統 — POC 提案與技術規劃
author: "Asgard AI"
date: "2026-05-16"
lang: zh-Hant
mainfont: "PingFang TC"
CJKmainfont: "PingFang TC"
monofont: "Menlo"
geometry:
  - a4paper
  - margin=2cm
linkcolor: "NavyBlue"
toc: true
toc-depth: 2
numbersections: false
header-includes:
  - \\usepackage{xeCJK}
  - \\setCJKmainfont{PingFang TC}
  - \\usepackage{fancyhdr}
  - \\pagestyle{fancy}
  - \\fancyhf{}
  - \\fancyhead[L]{杰隆印刷 × Asgard AI　POC 提案}
  - \\fancyhead[R]{\\thepage}
  - \\renewcommand{\\headrulewidth}{0.3pt}
---

"""

# Strip the existing top-level title (we provide it via frontmatter)
# Original starts with: "# 杰隆印刷 × Asgard AI\n## AI 自動接單系統 — POC 提案與技術規劃\n..."
# We'll strip those two lines and the "文件版本/日期/提案方/客戶" metadata block,
# leaving the document starting from the first "---" separator.
header_end = src.find("---\n\n## 目錄")
if header_end == -1:
    raise SystemExit("FAIL: could not locate '---' marker before TOC")
body = src[header_end + len("---\n\n"):]   # skip past the separator

# Compose final
final = FRONTMATTER + body + APPENDIX_A + APPENDIX_B

BUILD_MD.write_text(final, encoding="utf-8")
print(f"Wrote: {BUILD_MD}  ({len(final):,} chars)")
