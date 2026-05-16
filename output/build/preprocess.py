#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Pre-process proposal markdown for pandoc → xelatex → PDF.

- Replace 5 ASCII chart blocks with PNG references (single-image,
  no multi-page PDF surprises).
- Substitute glyphs PingFang TC doesn't carry.
- Append Appendix C (demo screenshots) and Appendix D (quotation) —
  the source proposal already has A/B so we continue the lettering.
- The matching template lives at output/build/asgard.latex.
"""

from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent.parent
PROPOSAL = ROOT / "POC_Proposal_杰隆印刷_AI接單系統.md"
BUILD_MD = ROOT / "output" / "build" / "proposal.md"

src = PROPOSAL.read_text(encoding="utf-8")

# ---- Glyph fixes ------------------------------------------------------------

GLYPH_FIX = {
    "⚠️": "★ ",
    "⚠":  "★ ",
    "☑":  "■ ",
    "☐":  "□ ",
    " ":  " ",      # EN QUAD → ordinary space
    " ":  " ",      # NARROW NO-BREAK SPACE
}
for old, new in GLYPH_FIX.items():
    src = src.replace(old, new)

# ---- ASCII → image swaps ----------------------------------------------------
# Captions intentionally have NO leading "圖 N：" because pandoc's
# captionsetup auto-prefixes with "圖 N" via the figure environment.

REPLACEMENTS = [
    (
        "```\n散客進線\n   ↓\n等待業務回應（可能超過 3 天）\n   ↓\n多次來回確認規格（術語不熟、資訊不完整）\n   ↓\n客戶等不及 → 流失訂單\n```",
        "![散客詢價痛點流程](../diagrams/01-pain-flow.png){ width=50% }",
    ),
    (
        "```\n客戶白話詢問\n     ↓\n[Odin Workflow] — 對話編排 + 意圖識別節點 + 欄位缺口追蹤\n     ↓\n[客製化後端] — 圖檔解析 / 模擬圖生成 / 舊客比對 / 情緒偵測通知\n     ↓\n[基本 DB] — 訂單草稿 / 客戶資料 / 對話紀錄\n     ↓\n訂單草稿 → 人工確認 → 正式成立\n```",
        "![POC 解決方案流（Phase 1：Odin + 客製化 + 基本資料層）](../diagrams/02-solution-flow.png){ width=72% }",
    ),
    (
        "```\n┌─────────────────────────────────────────────────────────┐\n│                     使用者層 (User Layer)                 │\n│  ┌─────────────┐  ┌──────────────┐  ┌──────────────┐    │\n│  │ 網頁對話框   │  │ LINE 頻道     │  │ 管理後台      │    │\n│  │ (POC 主要)  │  │ (Phase 2)    │  │ (數據/接管)   │    │\n│  └─────────────┘  └──────────────┘  └──────────────┘    │\n└──────────────────────────┬──────────────────────────────┘\n                           │\n┌──────────────────────────▼──────────────────────────────┐\n│                    對話層 (Conversation Layer)            │\n│                                                          │\n│  路由器（統一入口）→ Session 管理 → 意圖識別              │\n│       ↓                                                  │\n│  意圖分類：[詢價] [訂單追蹤] [舊客補印] [閒聊] [轉接]    │\n│       ↓                                                  │\n│  轉接處理：情緒偵測 → Email 通知真人接管                  │\n└──────────────────────────┬──────────────────────────────┘\n                           │\n┌──────────────────────────▼──────────────────────────────┐\n│                    AI 核心 (AI Core)                      │\n│                                                          │\n│  ┌─────────────────────────────────────────────────┐    │\n│  │  ★ 訂單資料引導引擎 (Field Gap Detection Engine)  │    │\n│  │  · 已收集欄位追蹤                                  │    │\n│  │  · 缺口清單即時更新                               │    │\n│  │  · 白話追問生成                                   │    │\n│  └─────────────────────────────────────────────────┘    │\n│                                                          │\n│  LLM 語言模型 | 知識庫 RAG | Prompt 引擎 | 圖像分析      │\n└──────────────────────────┬──────────────────────────────┘\n                           │\n┌──────────────────────────▼──────────────────────────────┐\n│                     資料層 (Data Layer)                   │\n│                                                          │\n│  訂單草稿 DB │ 客戶資料庫 │ 知識庫 │ 對話紀錄             │\n│  （欄位映射）│（舊客比對）│（FAQ+規格）│（訓練用）         │\n└─────────────────────────────────────────────────────────┘\n```",
        "![四層架構：使用者層 → 對話層 → AI 核心 → 資料層](../diagrams/03-architecture-layers.png){ width=88% }",
    ),
    (
        "```\n[Webhook 入口] → [Session 建立 / 舊客比對]\n      ↓\n[LLM Router] → 意圖分類（詢價 / 追蹤 / 補印 / 閒聊）\n      ↓\n[Odin LLM 節點] → 欄位缺口偵測 → 白話追問生成\n      ↓\n[圖檔上傳節點] → [LLM Vision 解析] → 回傳色數建議\n      ↓\n[Image Gen 節點] → 貼標模擬圖生成\n      ↓\n[條件分支] → 完成 / 觸發轉接 / 繼續追問\n      ↓\n[訂單 JSON 輸出] → 寫入客製化 DB + 通知業務 Email\n```",
        "![Odin 工作流節點配置（Phase 1）](../diagrams/04-odin-workflow.png){ width=76% }",
    ),
    (
        "```\nWeek 1 ─────────────────────── Week 2 ─────────────────────── Week 3\n   │                               │                               │\n   ▼                               ▼                               ▼\n知識庫 Ready              欄位引導引擎完成              POC 上線 Demo\n材質對照表建立             舊客識別功能                  全流程驗收\nOdin 工作流草稿            情緒偵測觸發                  業務後台查看\n                           圖檔解析功能\n```",
        "![POC 導入時程 Gantt（4-6 週版本）](../diagrams/05-timeline.png){ width=92% }",
    ),
    # 附錄 A — 對話流程圖（ASCII → Mermaid）
    (
        "```\n[客戶進入] ──→ [意圖識別]\n                    │\n        ┌───────────┼───────────┐\n        ▼           ▼           ▼\n    [詢價]      [訂單追蹤]    [閒聊/其他]\n        │           │\n        ▼           ▼\n  [舊客比對]   [查詢訂單DB]\n        │\n   ┌────┴────┐\n   ▼         ▼\n[舊客]     [新客]\n   │         │\n   ▼         ▼\n[帶入資料] [建立Session]\n   │         │\n   └────┬────┘\n        ▼\n[欄位缺口偵測] ←──────────────┐\n        │                      │\n        ▼                      │\n[生成下一個白話追問]           │\n        │                      │\n        ▼                      │\n[客戶回應] ──→ [解析填入欄位] ─┘\n        │\n        ▼\n  [欄位是否完整?]\n        │\n   ┌────┴────┐\n   │         │\n  [是]      [否] ──→ 繼續追問\n   │\n   ▼\n[觸發模擬圖?] ──→ [生成貼標情境圖]\n   │\n   ▼\n[訂單草稿確認]\n   │\n   ▼\n[輸出 JSON + 通知業務]\n   │\n   ▼\n[等待業務確認是否成立訂單]\n```",
        "![對話流程：從客戶進入到訂單成立](../diagrams/06-dialog-flow.png){ width=80% }",
    ),
]

for old, new in REPLACEMENTS:
    if old not in src:
        raise SystemExit(f"FAIL: could not locate ASCII block starting with:\n{old[:100]}…")
    src = src.replace(old, new, 1)

# ---- Strip the source's redundant top heading + metadata block --------------
# (we provide title via YAML / template cover page)
header_end = src.find("---\n\n## 目錄")
if header_end == -1:
    raise SystemExit("FAIL: could not locate '---' marker before TOC")
body = src[header_end + len("---\n\n"):]

# ---- Strip "N. " from `## N. xxx` and "N.M " from `### N.M xxx` ------------
# (LaTeX will provide auto-numbering via section counter)
import re
body = re.sub(r"^(## )\d+\.\s+", r"\1", body, flags=re.MULTILINE)
body = re.sub(r"^(### )\d+\.\d+\s+", r"\1", body, flags=re.MULTILINE)
body = re.sub(r"^(#### )\d+\.\d+\.\d+\s+", r"\1", body, flags=re.MULTILINE)

# Mark existing appendices in source ('## 附錄 A / B') as unnumbered so they
# don't grab section numbers (13, 14). Pandoc syntax `{-}` → \section*.
body = re.sub(r"^(## 附錄 [A-Z]：[^\n]+)$", r"\1 {-}", body, flags=re.MULTILINE)

# Subheadings inside appendices (`### A.1 ...` / `### B.1 ...` etc., and
# the quotation's `#### W1-W2 Phase A ...`) should be unnumbered too —
# otherwise LaTeX continues the section counter and shows "12.2.3" prefixes.
body = re.sub(r"^(### [A-Z]\.\d+[^\n]*)$", r"\1 {-}", body, flags=re.MULTILINE)
body = re.sub(r"^(#### W\d[^\n]*)$", r"\1 {-}", body, flags=re.MULTILINE)

# Drop the original markdown TOC (lines 11..24 of source) — pandoc will
# generate its own. The original starts with "## 目錄" and ends at "---".
toc_start = body.find("## 目錄")
toc_end = body.find("---", toc_start)
if toc_start != -1 and toc_end != -1:
    body = body[:toc_start] + body[toc_end + len("---") + 1:]

# ---- Appendix C: demo screenshots -------------------------------------------

APPENDIX_C = r"""

\newpage

## 附錄 C：Demo 操作截圖 {-}

下列截圖摘錄自實際運行的 POC Demo（`index.html`），呈現一次完整的「保養品瓶標」詢價流程，從歡迎畫面一路到結構化訂單輸出。

![**歡迎畫面**　　AI 詢價助理啟動，提供「第一次詢價 / 補印舊款式」兩種入口；右側為訂單欄位追蹤面板（0 / 10）。](../screenshots/01-welcome.png){ width=92% }

\vspace{0.6em}

![**對話進行中**　　客戶輸入品名與尺寸後，AI 接續詢問印製數量；建議快選按鈕降低輸入負擔，右側追蹤面板同步更新已收集欄位。](../screenshots/02-mid-conversation.png){ width=92% }

\vspace{0.6em}

![**材質智能引導**　　8 種官網提供之材質卡片排列，含特性說明與標籤（最熱門 / 質感首選 等），客戶以白話描述即可一鍵選擇。](../screenshots/03-material-grid.png){ width=92% }

\vspace{0.6em}

![**AI 模擬圖生成**　　依客戶選擇的材質 / 尺寸即時渲染貼標於容器之情境模擬圖，協助客戶確認設計效果。](../screenshots/04-ai-mockup.png){ width=92% }

\vspace{0.6em}

![**訂單需求確認單**　　對話完成後輸出結構化需求確認，業務人員可直接使用此資訊進行報價，無需再次電話確認規格。](../screenshots/05-summary-card.png){ width=92% }
"""

# ---- Appendix D: quotation (no signature block) -----------------------------

APPENDIX_D = """

\\newpage

## 附錄 D：POC 報價單明細 {-}

### D.1　Asgard 平台訂閱費（持續性月費，與下方一次性費用分開計算） {-}

| 方案 | 月費 | 規格 |
|------|------|------|
| **Basic-Lite**（本次推薦） | NT$ 20,000 / 月 | 1 Agent Hub user、1 Data Insight user、1,000 Asgard Credits |

**付款方式（擇一）：**

| 選項 | 金額 | 備註 |
|------|------|------|
| □　信用卡月扣 | NT$ 20,000 / 月 | 每月自動扣款 |
| □　年付開立發票 | NT$ 240,000 / 年 | 一次開立，預付一年 |

> 依正式上線使用量可升級至 Regular（NT$ 50,000 / 月）或 Plus（NT$ 100,000 / 月）。

### D.2　一次性 POC 導入服務（4-6 週，單價 NT$ 20,000 / 人天） {-}

#### W1-W2　Phase A：基礎架構 + DB Schema + AP Server 部署 {-}

| 項目 | 工作內容 | 人天 | 小計 |
|------|---------|:---:|---:|
| 雲端環境規劃與建置 | GCP / AWS 帳號、VPC、IAM、Secret 管理 | 0.5 | NT$ 10,000 |
| DB Schema 設計與建置 | 訂單草稿 / 客戶 / 對話紀錄 / 知識庫 | 1 | NT$ 20,000 |
| AP Server 部署 | 容器化、反向代理、TLS 設定 | 1 | NT$ 20,000 |
| CI/CD 與基本監控初始化 | GitHub Actions、Log / Error 監控 | 0.5 | NT$ 10,000 |
| **Phase A 階段小計** | | **3** | **NT$ 60,000** |

#### W3-W4　Phase B：核心對話流 + Odin 工作流建置 {-}

| 項目 | 工作內容 | 人天 | 小計 |
|------|---------|:---:|---:|
| Odin 工作流設計 | 接單主流程、條件分支、轉接邏輯 | 1 | NT$ 20,000 |
| 對話 Prompt 與欄位偵測邏輯 | LLM Prompt、缺口偵測、白話追問 | 1 | NT$ 20,000 |
| Chatbot 前端 UI 開發 | 嵌入式對話框、訊息流、模擬圖預覽 | 1 | NT$ 20,000 |
| Session 管理 + 訂單草稿 API | Session、訂單 JSON schema、後端 API | 1 | NT$ 20,000 |
| **Phase B 階段小計** | | **4** | **NT$ 80,000** |

#### W5-W6　Phase C：進階功能 + UAT + 教育訓練 {-}

| 項目 | 工作內容 | 人天 | 小計 |
|------|---------|:---:|---:|
| 圖檔上傳解析 | LLM Vision 串接、色數與尺寸建議 | 0.5 | NT$ 10,000 |
| AI 模擬圖生成 | Image Gen API 串接、情境圖輸出 | 0.5 | NT$ 10,000 |
| 舊客識別查詢 | Email / 電話比對、歷史訂單帶入 | 0.5 | NT$ 10,000 |
| 情緒偵測 + 真人轉接通知 | 情緒分類、Email / LINE Webhook | 0.5 | NT$ 10,000 |
| UAT 支援 + 上線部署 | Bug 修正、上線切換、煙霧測試 | 0.5 | NT$ 10,000 |
| 教育訓練與文件交付 | 業務操作訓練、技術交接文件 | 0.5 | NT$ 10,000 |
| **Phase C 階段小計** | | **3** | **NT$ 60,000** |

### D.3　Phase 2 進階模組（另行報價，不計入本次總價） {-}

| 模組 | 升級價值 | 費用 |
|------|---------|------|
| Sindri Agent Hub | 多 Agent 並行協作、跨對話記憶、複雜詢價對話品質升級 | 另行報價 |
| Mimir Data Insight | 訂單轉換漏斗、客群分析、材質偏好趨勢儀表板 | 另行報價 |

### D.4　一次性 POC 導入服務 加總 {-}

| 項目 | 金額 |
|------|---:|
| 一次性導入服務小計（未稅） | NT$ 200,000 |
| 營業稅 5% | NT$ 10,000 |
| **含稅總價** | **NT$ 210,000** |

> ★ 本份總價僅為 POC 一次性導入服務費用；Asgard 平台訂閱費另計（見 D.1 區）。

### D.5　假設與排除 {-}

1. 不含杰隆印刷既有 ERP / 生產系統的整合與對接
2. 不含 LINE 官方帳號 / FB Messenger 等通道串接（Phase 2 規劃）
3. 不含自動報價計算邏輯（需杰隆提供定價模型）
4. 雲端基礎設施費用（GCP / AWS）由客戶以實際使用量支付

### D.6　付款條件與有效期 {-}

- **付款條件建議**：簽約 30% / 中期驗收 40% / 最終驗收 30%
- **報價有效期**：自報價日起 30 日
- **本次報價日期**：__TODAY__
""".replace("__TODAY__", date.today().strftime("%Y-%m-%d"))

# ---- YAML frontmatter (consumed by pandoc / asgard.latex template) ---------

FRONTMATTER = """---
title: "杰隆印刷 × Asgard AI — POC 提案"
author: "Asgard AI"
date: "2026-05-16"
lang: zh-Hant
---

"""

final = FRONTMATTER + body + APPENDIX_C + APPENDIX_D

BUILD_MD.write_text(final, encoding="utf-8")
print(f"Wrote: {BUILD_MD}  ({len(final):,} chars)")
