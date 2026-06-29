# Excel-to-Slide Studio

Turn a budget/financial Excel workbook into a board-ready PowerPoint deck — and keep every deck in sync as the workbook changes.

A fully client-side web app (HTML + vanilla JS). The workbook is parsed in the browser; nothing is uploaded to a server.

## Highlights

- **Generate a deck from Excel** — upload a consolidation workbook (`.xlsx` / `.xlsm`) and get an editable, branded slide deck.
- **AI Deck Copilot** — change slides, charts, data sources, template and branding in plain English (in-browser engine by default; pluggable Azure OpenAI proxy endpoint).
- **Lineage & versioning (the key feature)** — every deck remembers which *version* of the Excel it was built from. Load a newer version and the decks that should change light up:
  - A workspace with a visual **version timeline** (v1 → v2 → v3) and a deck map.
  - Per-deck update policy: **Notify**, **Auto-update**, or **Pinned** (frozen).
  - A role-based **diff** (old → new, %), one-click **refresh** that updates the numbers everywhere, **highlights** changed cells (green up / red down), and inserts a **"What changed"** summary slide.
- **One-click export** to an editable `.pptx` (native tables, chart images).

## Architecture

The AI never draws slides directly. It authors a structured **deck-spec** (a single source of truth); a deterministic engine renders that spec into both the live preview and the PowerPoint. This keeps figures **consistent** (a value like ROE is defined once and reused), **traceable** (each value carries its source), and **governable** (changes are reviewable and reversible).

```
Excel workbook → parse → deck-spec (JSON: theme + slides + data + sources)
        AI Copilot edits the spec  →  deterministic render → preview + .pptx
        New version → role-based diff → refresh decks (notify / auto / pinned)
```

## Run locally

It is a static site — no build step.

```bash
# from the project root
python3 -m http.server 8000
# then open http://localhost:8000
```

Or just open `index.html` in a browser.

## Access gate

The landing screen has a simple shared-password gate. The password is a placeholder in this repo:

```js
const ACCESS_PW = "CHANGE_ME"; // set your own, or remove the gate
```

This is **not** a security mechanism (the value ships in client-side JS). For real access control, put the site behind proper authentication (e.g. Microsoft Entra ID via Azure Static Web Apps).

## Tech

- [SheetJS](https://sheetjs.com/) — Excel parsing
- [Chart.js](https://www.chartjs.org/) — chart previews
- [PptxGenJS](https://gitbrent.github.io/PptxGenJS/) — PowerPoint export

(All loaded via CDN.)

## Notes

This started as a proof of concept for automating a group budget deck. The lineage/versioning prototype stores its ledger in the browser's `localStorage`; a production version would move that to a backend (e.g. SharePoint + Azure Functions) with proper auth and an audit trail.
