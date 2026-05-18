# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this repo is

A **client-facing POC bundle** pitched to 杰隆印刷 (Jlong Printing) by Asgard AI. It is not a product or an app — it is a presentation artifact in three pieces:

- `index.html` — a self-contained, scripted "AI 詢價助理" chatbot demo (zh-TW). Single file, no build, no backend; all state lives in browser memory.
- `flow.html` — a static marketing/architecture page explaining the proposed ordering flow. Phase 1 (本次 POC) uses Asgard's **Odin** workflow engine plus **客製化** front/back-end and basic DB / AP Server; **Sindri** and **Mimir** appear as a Phase 2 future-expansion card near the page end. Linked from `index.html` via the header.
- `POC_Proposal_杰隆印刷_AI接單系統.md` — the written proposal (scope, architecture, acceptance criteria, timeline). This is the source of truth for *what* the system is supposed to do; the HTML files are what the demo currently *shows*.

Source references for the printing domain (materials, order tracking spreadsheets, original requirements doc, drawio flow PDF) live under `references/`. Treat them as read-only inputs.

## How to run / preview

There is no build system, no package manager, and no test suite. Open the HTML files directly in a browser, or serve the directory statically:

```bash
python3 -m http.server 8000   # then visit http://localhost:8000/index.html
```

Fonts load from Google Fonts at runtime, so the demo needs internet access on first view.

## Editing the chatbot demo (`index.html`)

The whole demo is one HTML file with one `<script>` block. The architecture is small but specific:

- **`MATERIALS`** (around line 383): the **only** allowed materials are the 8 listed here, mirroring the items on `jlongprinting.com/label`. The proposal explicitly constrains material choices to the official catalog — do not invent new materials when changing copy or adding flows.
- **`ORDER_FIELDS`** (around line 405): the 10 Category-A fields the demo collects. These drive the right-hand progress tracker via `buildTracker` / `updateTracker` / `setActiveField`. Keep this list and the step→field mapping inside `getCurrentFieldKey` in sync — they are coupled by string key (`品名`, `尺寸`, …) and silently drift if only one side is edited.
- **`FLOW`** (around line 432): the conversation script as an array of step objects. Each step has `aiMsg()`, an `inputType` (`quickreply` | `text` | `material-grid` | etc.), and an `onSelect`/`onSubmit` callback that writes to `state.order` and calls `advance()`. New conversation steps are added by inserting into this array; rendering for each input type is dispatched inside `renderInput`.
- **`state`** is a single module-level object (`{ step, mode, order, done }`). `mode` switches between the `'new'` customer flow and the `'returning'` (補印舊款式) flow, which has its own handlers (`handleReturning`, `showReturningResult`, `handleReturningQuantity`, `handleReturningModify`).
- **Mockup rendering** (`buildMockupSVG`, `getLabelColor`) generates an inline SVG bottle preview from the collected order. Label color comes from a map keyed by the Chinese material name — when adding a material to `MATERIALS`, add a matching entry to `getLabelColor`'s `map` or the preview falls back to plain cream.
- Messages stream into `#messages` via `addAiMsg` / `addUserMsg` / `addElement`, with a fake "typing" delay (`showTyping` / `removeTyping`) of ~700–1100 ms in `advance()` to make pacing feel natural. Don't remove the delay — it's load-bearing for the demo feel.

The demo is in Traditional Chinese (zh-TW). Keep copy in zh-TW; mixing in English or Simplified Chinese breaks the pitch.

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

## Conventions worth knowing

- **The proposal leads, the demo follows.** If the proposal (`POC_Proposal_...md`) and the demo conflict on scope (e.g., which fields are collected, which materials exist, whether quoting is automated), the proposal is canonical — the demo is a simplified slice. Don't expand the demo beyond Category-A fields without checking section 4 ("POC 範疇定義") of the proposal first.
- **No dependencies, keep it that way.** This is shipped as files a non-technical stakeholder can double-click. Don't introduce a bundler, framework, or npm install step unless the user explicitly asks.
- **`references/` is input, not output.** The `.xlsx`, `.docx`, `.pdf`, and `.drawio.pdf` files are client-provided artifacts. Don't modify or regenerate them.
