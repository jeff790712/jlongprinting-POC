# Asgard 架構敘事修正 + POC 報價單產出 — 設計規格

**日期**：2026-05-16
**狀態**：草稿（待使用者最終確認）
**範圍**：杰隆印刷 POC bundle 的文字敘事修正 + 新增 Excel 報價單

---

## 1. 背景與動機

現行 POC bundle（`POC_Proposal_杰隆印刷_AI接單系統.md`、`flow.html`、`index.html`、`CLAUDE.md`）以 **Odin + Sindri + Mimir 三大引擎** 為架構敘事主軸。但本次 POC 實際只會用到：

- **Odin (Asgard Studio)** — 零程式碼 AI 工作流編排
- **客製化前後端開發** — Chatbot UI、Session 管理、訂單 API、舊客查詢
- **基本 DB + AP Server** — 替代 Mimir 在本階段的資料層角色

Sindri（Agents Hub）與 Mimir（Data Insight）為獨立商業模組，本 POC 不導入；目前的文件描述把它們的工作攤到本次 POC 範圍內，造成銷售敘事與實際交付不符。同時客戶需要一份可填寫人天並自動加總的 Excel 報價單。

## 2. Asgard 平台實際產品線（驗證自 asgard-ai.com / 2026-05-16）

| 產品 | 中文名 | 定位 |
|------|--------|------|
| **Odin** | Asgard Studio | 零程式碼 AI 工作流編排 |
| **Sindri** | Asgard Agents Hub | 多 Agent 自動化執行 |
| **Mimir** | Asgard Data Insight | 資料洞察與智慧決策 |
| **Yggdrasil** | 開源生態系 | 63+ MCP Servers / 277+ SKILLs / 10+ Solution Bundles（MIT） |

**Asgard 標準訂閱定價**（自 asgard-ai.com/pricing）：

| 方案 | 月費 | Agent Hub users | Data Insight users | Credits |
|------|------|----------------|--------------------|---------|
| Basic-Lite | NT$ 20,000 | 1 | 1 | 1,000 |
| Basic-Regular | NT$ 50,000 | 10 | 5 | 5,000 |
| Basic-Plus | NT$ 100,000 | 25 | 10 | 10,000 |
| 專業導入服務 | NT$ 20,000 / 人天（1 人天 = 8 hr） | — | — | — |

## 3. 修正方向（Phase 框架）

### Phase 1 — 本次 POC（4-6 週）
- **Asgard 平台**：Odin Studio（工作流編排、LLM/Vision/Image Gen API 串接的承載層）
- **客製化前端**：Chatbot 嵌入式對話框、AI 模擬圖預覽
- **客製化後端**：Session、訂單草稿 API、舊客查詢 API、情緒偵測通知
- **基本 DB**：訂單草稿、客戶資料、對話紀錄
- **基本 AP Server**：雲端部署、CI/CD、監控

### Phase 2 — 未來擴展（另行報價）
- **Sindri Agent Hub** — 真正的多 Agent 編排（取代 Phase 1 由 Odin 節點 + 客製化程式組成的單一 LLM 對話）
- **Mimir Data Insight** — 數據洞察儀表板（取代 Phase 1 的基本 DB 報表）

## 4. 檔案異動範圍

| # | 檔案 | 動作 | 重點 |
|---|------|------|------|
| 1 | `POC_Proposal_杰隆印刷_AI接單系統.md` | 編輯 | § 3.2 工具對應表、§ 5 系統架構映射改寫；新增「Phase 2 進階擴展」一節 |
| 2 | `flow.html` | 編輯 | 「三大 AI 引擎」section 重構為 **Odin 主柱 + 客製化交付柱**；流程圖負責模組標籤同步；末段加 Phase 2 預告卡片 |
| 3 | `index.html` | 檢查 + 微調 | 掃描 header/footer/comments 中是否有 Sindri/Mimir/三引擎描述 |
| 4 | `CLAUDE.md` | 編輯 | 對應段落（flow.html 三引擎描述）更新為 Phase 1 + Phase 2 框架 |
| 5 | `POC_報價單_杰隆印刷.xlsx` | 新增 | 單一工作表、含公式的 Excel 報價單 |
| 6 | `scripts/build_quotation.py` | 新增 | openpyxl 腳本，產生上述 xlsx |

## 5. Excel 報價單設計

### 5.1 結構（單一工作表，自上而下）

報價單分兩塊：上方 **Asgard 平台訂閱**（固定月費，獨立計算，不進總價）+ 下方 **一次性 POC 導入服務**（人天 × 單價，是本份報價單的主要計算對象）。

```
[抬頭區]
  Asgard AI × 杰隆印刷  POC 報價單
  報價日期 / 有效期 30 日 / 案件編號 / 聯絡人

[A] Asgard 平台訂閱費（持續性月費，與下方一次性費用分開）
  方案：Basic-Lite
  月費：NT$ 20,000 / 月

  付款方式（擇一）：
    ☐ 信用卡月扣        NT$ 20,000 / 月
    ☐ 年付開立發票      NT$ 240,000 / 年（一次開立）

  附註：依正式上線使用量可升級至 Regular（NT$50,000/月）/ Plus（NT$100,000/月）

[B] 一次性 POC 導入服務（4-6 週，分三階段）
  欄位：項目 | 工作內容 | 人天 | 單價 | 小計 (=人天 × 單價)

  W1-W2 Phase A：基礎架構 + DB Schema + AP Server 部署
    - 雲端環境規劃與建置
    - DB Schema 設計（訂單草稿 / 客戶 / 對話紀錄）
    - AP Server 部署
    - CI/CD 與基本監控初始化
  → 階段小計 (SUM)

  W3-W4 Phase B：核心對話流 + Odin 工作流建置
    - Odin 工作流設計（接單主流程）
    - 對話 Prompt 與欄位偵測邏輯
    - Chatbot 前端 UI 開發
    - Session 管理 + 訂單草稿 API
  → 階段小計 (SUM)

  W5-W6 Phase C：進階功能 + UAT + 教育訓練
    - 圖檔上傳解析（LLM Vision 串接）
    - AI 模擬圖生成（Image Gen API 串接）
    - 舊客識別查詢
    - 情緒偵測 + 真人轉接通知
    - UAT 支援 + 上線部署
    - 教育訓練與文件交付
  → 階段小計 (SUM)

[C] Phase 2 進階模組（另行報價，不計入本次總價）
  Sindri Agent Hub（多 Agent 編排）   金額欄位：另行報價
  Mimir Data Insight（數據洞察儀表板） 金額欄位：另行報價

[D] 一次性 POC 導入服務 加總（含公式）
  一次性導入服務小計（未稅）   = SUM(三階段小計)
  營業稅 5%                    = ROUND(未稅 × 0.05, 0)
  含稅總價                     = 未稅 + 營業稅    ← 粗體 + 強調底色

  ★ 本份總價僅為 POC 一次性導入服務費用；Asgard 平台訂閱費另計（見 [A] 區）

[E] 備註區
  假設與排除（範本 4 條）
  付款條件（建議：簽約 30% / 中期 40% / 驗收 30%）
  報價有效期 30 日
  簽核欄（客戶 / Asgard AI）
```

### 5.2 公式設計

- 各項小計 = `=C×D`（人天 × 單價）
- 階段小計 = `=SUM(階段內小計欄)`
- 一次性導入小計（未稅） = `=SUM(三階段小計)`
- 營業稅 = `=ROUND(未稅小計 × 0.05, 0)`
- 含稅總價 = `=未稅小計 + 營業稅`

人天欄位留空、單價已預填 NT$ 20,000，使用者填入人天後所有公式即時更新。

[A] 區 Asgard 平台訂閱費為固定靜態文字（NT$ 20,000/月、NT$ 240,000/年），不參與 [D] 區加總，避免使用者誤把訂閱費當成一次性費用報給客戶。

### 5.3 格式

- 字體：微軟正黑體 / Arial fallback
- 貨幣顯示：`NT$ #,##0`
- 標題列：深藍底（#0F3460）+ 白字（對應 index.html / flow.html 既有品牌色）
- 階段分隔列：金色底（#C9A84C）強調
- 含稅總價列：粗體 + 淺金底
- 欄寬：項目欄寬 30、工作內容欄寬 50、數值欄寬 12-16
- 凍結窗格：抬頭區下方

## 6. 修正原則（不變項）

- **不擴張功能範疇**：原 proposal 承諾的功能（圖檔解析、AI 模擬圖、舊客識別、情緒偵測、多語）全保留；只改「由誰承擔」（Sindri/Mimir → Odin 工作流節點 + 客製化程式 + 基本 DB）
- **zh-TW 維持**：所有修改保持繁體中文
- **品牌色維持**：Odin 藍 #1e3a8a 在 Phase 1 強調；Sindri 紫 #5b21b6、Mimir 綠 #065F46 在 Phase 2 區塊使用
- **demo 行為不動**：`index.html` 的 chatbot 邏輯（FLOW、MATERIALS、ORDER_FIELDS 等）不修改
- **proposal 為正典**：若 demo 與 proposal 衝突仍以 proposal 為準（CLAUDE.md 既有原則）

## 7. 驗收標準

- [ ] 4 份既有檔案內所有 Sindri/Mimir 描述都已修正為「Phase 1 由 Odin + 客製化開發承擔；Sindri/Mimir 列為 Phase 2 未來擴展」
- [ ] `POC_報價單_杰隆印刷.xlsx` 在 Excel/Numbers 中開啟，公式自動計算正確
- [ ] 將 Phase A/B/C 各填 1 人天測試後，加總區公式正確：未稅 NT$ 60,000、營業稅 NT$ 3,000、含稅 NT$ 63,000；[A] 區訂閱費為靜態文字不參與計算
- [ ] `flow.html` 的 Odin 主柱配色與 index.html 既有品牌色一致
- [ ] CLAUDE.md 更新後新進 Claude 實例能正確理解 Phase 1 / Phase 2 框架

## 8. 不在本次範圍

- Asgard 平台實際採購流程、合約用印、技術 POC 環境建置
- 與杰隆印刷實際業務人員的需求訪談
- demo 內 chatbot 邏輯本身的功能調整
- references/ 內客戶提供文件的修改
