# 📊 AI One-Pager Teaser Maker

Generate professional PowerPoint presentations from Excel data with local AI-powered content generation. No API keys required.

## Quick Start

### GitHub Codespaces
App auto-starts on Codespace launch. Check **Ports tab** (bottom panel) for port 8501, then click the globe icon.

**502 Error?** Run `bash restart.sh` and check Ports tab for new URL.

### Local Development
```bash
bash start.sh                # Start app (http://localhost:8501)
bash restart.sh              # Restart if not working
bash keep-alive.sh           # Auto-restart monitoring
pip install -r requirements.txt  # First-time setup
```

## Features

- **Local AI Generation** — 3-paragraph highlights, sector labels, client one-liners (no API costs)
- **Flexible Excel Input** — Column-based, row-based, or multi-sheet (.xlsx/.xlsm)
- **Smart Matching** — Fuzzy label matching (60% similarity threshold)
- **Multiple Placeholder Types** — Direct `[Name]`, tags `:NAME:`, AI prompts `[AI: ...]`
- **Module Auto-Detection** — Maps sponsor/company/project Excel labels to standardized PPT tags

## Usage

1. Upload PPTX template
2. Upload Excel file (auto-detects format)
3. Configure AI tone if using AI prompts
4. Generate and download

## Excel Formats

**Column-Based:** Traditional spreadsheet with headers
**Row-Based:** Label/Value pairs in two columns
**Multi-Sheet:** Organized data across sheets (reads all automatically)

## Template Placeholders

**Direct:** `[Company Name]`, `[Country]`
**Tags:** `:COMPANY_NAME:`, `:SECTOR:`
**AI Prompts:**
- `[AI: Write a 3-paragraph Project Highlight for {Sector}...]`
- `[AI: Generate 1-3 word sector label from "{Sector}"...]`
- `[AI: Write one sentence starting with "The client"...]`

## Module Mappings

System maps department-specific Excel labels to standardized PPT tags:

**Module 1 (Sponsor):** "Sponsor Name" → `:ISSUER:`
**Module 2 (Company):** "Company Legal Name" → `:ISSUER:`
**Module 3 (Project):** "Project Name" → `:ISSUER:`

Standard tags: `:ISSUER:`, `:INDUSTRY:`, `:JURISDICTION:`, `:ISSUANCE_TYPE:`, `:INITIAL_NOTIONAL:`, `:COUPON_RATE:`, `:COUPON_FREQUENCY:`, `:TENOR:`, `:CLIENT_SUMMARY:`, `:PROJECT_HIGHLIGHT:`

## Troubleshooting

**502 Errors (Codespaces):** `bash restart.sh`, check Ports tab
**Still broken:** `pkill -f streamlit && bash restart.sh`
**Check status:** `lsof -i :8501`
**View logs:** `tail -f streamlit.log`
**Dependencies:** `pip install -r requirements.txt`

## Requirements

Python 3.8+, dependencies: `streamlit`, `pandas`, `python-pptx`, `openpyxl`

## Test Files

- `example[1-3]_*.xlsx` — Project examples
- `multi_sheet_example.xlsx` — Multi-sheet demo
- `test_module[1-3]_*.xlsx` — Module format examples
- `test_*_template.pptx` — Template examples
