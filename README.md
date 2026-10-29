# Lantern

<h3 align="center">Auto-Design PPT Office Automation Platform — One JSON, One Command, One Finished Deck</h3>

<p align="center">
  <img src="https://img.shields.io/badge/Platform-Windows-blue?logo=windows" alt="Platform">
  <img src="https://img.shields.io/badge/Python-3.10+-blue?logo=python" alt="Python">
  <img src="https://img.shields.io/badge/License-MIT-green" alt="License">
  <img src="https://img.shields.io/badge/WPS-Office%2012.0+-red" alt="WPS">
</p>

---

**Lantern** is an office automation platform focused on **auto-designing PPT decks**:
it drives WPS Presentation through Windows COM automation, so **one JSON file → one command → finished PPTX + PDF** —
no manual layout, reproducible, and re-stylable in one run.
It also ships a WPS Spreadsheet data-reporting capability (CSV → XLSX/PDF with formulas, formatting and charts) and
extends to Zotero academic workflow and Photoshop/Illustrator design automation.

> "Lantern" — turn "designing a deck" from tedious layout work into one command,
> so an AI Agent handles the repetitive part and you return to "expressing ideas with content".

## Why this project

Everyday decks — proposals, reports, plans — face three burdens:

- **Error-prone**: manual layout breaks pages, misaligns tables, mixes font sizes; one edit breaks the rest;
- **Not reproducible**: steps vary per person, so output style is inconsistent;
- **Slow**: a proposal deck takes 1–2 hours by hand, and revisions cost even more.

Lantern's answer: **freeze PPT design into a single command**.
WPS exposes a full COM automation interface on Windows (compatible with Microsoft Office VBA).
We build a JSON-driven auto-design engine on top of it: cover, rich text, tables and banner elements are all
declared in data, auto-laid-out, auto-rendered, and auto-verified.

## Core scenario: Auto-design PPT (fully automated workflow)

```bash
cd WPS/ppt-engine/examples
pip install pywin32 Pillow pymupdf
python build_from_json.py --json lantern/example1_hiking_plan_lite.json --out lantern/out --name example1_hiking_plan_lite
```

Output (≈10–20 s, with automatic page-count and overflow verification):

- `lantern/out/徒步团建策划案.pptx` — finished WPS Presentation, ready to play
- `lantern/out/徒步团建策划案.pdf` — one page per slide, compact layout

Demo: **Ningbo Kingfa New Material Co., Ltd. hiking team-building plan** (3 slides: cover / schedule / budget & support).
Change a budget or schedule number in the JSON and re-run — the new deck is generated automatically, fully reproducible.

### Demo output (one command, fully automated, zero manual layout)

The 3 finished slides below are produced by a single `example1_hiking_plan_lite.json` command (screenshots from WPS Presentation):

![Slide 1 · Activity cover](WPS/ppt-engine/examples/lantern/screenshots/ppt_p1.png)

![Slide 2 · Schedule](WPS/ppt-engine/examples/lantern/screenshots/ppt_p2.png)

![Slide 3 · Budget & support](WPS/ppt-engine/examples/lantern/screenshots/ppt_p3.png)

Headers, tables and banner strips are all declared in JSON and auto-laid-out; change any budget/schedule number and re-run to regenerate all three slides, with PPTX and PDF exported together.

## Platform capabilities

| Module | Capability |
|--------|------------|
| `WPS/ppt-engine` (core, new) | **JSON-driven PPT auto-design engine**: cover/rich/gray-box/table/banner elements, multi-column table auto-render, overflow verification, PPTX+PDF dual export |
| `cli_anything/wps` | WPS Writer/Calc/Presentation CLI: 47 commands + data reporting (summary formulas, range formatting, chart images, XLSX/PDF export) |
| `cli_anything/zotero` | Zotero academic agent: 27 skills (search/review/draft/reviewer/journal figures) |
| `illustrator-harness` | Adobe Illustrator programmatic vector creation & editing |
| `photoshop-harness` | Adobe Photoshop bitmap processing: layers/selection/text/filter/export |

## Auto-design PPT capabilities (new core)

- **Declarative JSON design**: one `slides[]` JSON defines the whole deck; element types include `header / rich / gray_box / table / redbar`
- **Multi-column table auto-render**: any column count is auto-composed from JSON (blue header row, zebra rows, emphasized first column), avoiding unstable WPS AddTable borders
- **Overflow verification**: every page is checked for out-of-bound content at export; `[verify] pages=N overflow=0` guarantees a clean deck
- **One-run re-styling**: change numbers in JSON (budget, schedule, headcount) and re-run — both formats update together
- **16 university-recruitment PPT cases**: Tsinghua/Peking/Fudan/USTC etc., with full build scripts and data.json, reusable as templates

## Quick start

```bash
# Auto-design PPT (core)
cd WPS/ppt-engine/examples && python build_from_json.py --json lantern/example1_hiking_plan_lite.json --out lantern/out --name example1_hiking_plan_lite

# Data reporting (bundled capability)
cd report_demo && python build_report.py

# WPS office
cli-anything-wps document new --type impress --name "demo"
cli-anything-wps export render output.pptx -p pptx

# Zotero academic
cli-anything-zotero skills pipeline thesis

# Illustrator / Photoshop design
cli-anything-illustrator project new logo.ai -w 500 -h 500
cli-anything-photoshop project new banner.psd -w 1920 -h 1080
```

## Requirements

- Windows 10/11
- WPS Office 12.0+ (PPT auto-design + data reporting) / Zotero 7+ / Adobe Illustrator 2023+ / Adobe Photoshop 2023+ (optional)
- Python 3.10+ + pywin32 + Pillow + pymupdf + matplotlib + openpyxl

## License

MIT (see [LICENSE](LICENSE))
