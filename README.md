# 📊 AI One-Pager Teaser Maker

**Generate professional PowerPoint presentations with AI-powered content — no API key required!**

This Streamlit web app automatically fills PPTX templates with data from Excel files and generates compelling teaser text using smart local AI. Perfect for creating investor one-pagers, deal teasers, company profiles, and marketing materials at scale.

## ✨ Key Features

- **🤖 Smart Local AI** — No API keys, no costs, completely free intelligent text generation
  - 3-paragraph Project Highlights
  - 2-3 word sector labels
  - Anonymous client one-liners
- **📊 Flexible Excel Support**
  - Column-based format (traditional)
  - Row-based Label/Value pairs
  - Multi-sheet workbooks (unified knowledge base)
  - Hybrid mode (reads both formats automatically)
  - .xlsx and .xlsm files (macros disabled for security)
- **🎯 Multiple Placeholder Formats**
  - Direct: `[Title]`, `[Company Name]`
  - Tags: `:COMPANY_NAME:`, `:SECTOR:`
  - AI prompts: `[AI: Write about {Company}...]`
- **🔍 Smart Matching** — Fuzzy label matching handles inconsistent wording
- **🎨 Format Preservation** — Maintains all PowerPoint formatting
- **⚡ Batch Processing** — Generate multiple presentations from rows of data
- **💾 Instant Download** — Get your filled PPTX file immediately

## Quick Start

**To start the app (recommended):**
```bash
bash start.sh
```

**To restart if it's not working:**
```bash
bash restart.sh
```

**For auto-restart monitoring (keeps app running permanently):**
```bash
bash keep-alive.sh
```

**For first-time setup or if dependencies are missing:**
```bash
pip install -r requirements.txt
```

The app will be available at: **http://localhost:8501**

## Troubleshooting

### Port 8501 not working?
- Run `bash restart.sh` - this will kill any stuck process and restart cleanly
- Or manually: `lsof -ti :8501 | xargs kill -9` then `bash start.sh`

### For Codespaces users:
- After starting, click "Open in Browser" when the port forwarding notification appears
- Or go to the "Ports" tab and click the globe icon next to port 8501
- The app auto-starts when you open the Codespace!

## 📋 What You Need

1. **PPTX Template** — PowerPoint with placeholders, tags, and/or AI prompts
2. **Excel File** — Data in any format:
   - Column-based (traditional spreadsheet)
   - Row-based Label/Value pairs
   - Multi-sheet workbooks
   - .xlsx or .xlsm files
3. **That's it!** — No API keys or external services required

## 📖 How to Use

1. **Upload Template** — Upload your PPTX template
2. **Upload Data** — Upload Excel (.xlsx or .xlsm) - auto-detects format
3. **Review Parsed Data** — System shows all fields it found
4. **Configure AI** — Set tone (short/medium/long) if using AI prompts
5. **Generate** — Click "Generate PPTX" 
6. **Download** — Get your filled PowerPoint file

## 🎯 Excel Format Options

### Option 1: Column-Based (Traditional)
```
| Company Name | Sector      | Country   |
|--------------|-------------|-----------|
| ABC Corp     | Solar       | Vietnam   |
```

### Option 2: Row-Based (Label/Value)
```
| Label         | Value      |
|---------------|------------|
| Company name  | ABC Corp   |
| Sector        | Solar      |
| Country       | Vietnam    |
```

### Option 3: Multi-Sheet (Organized)
```
Sheet "Company Info":  Company name | ABC Corp
Sheet "Project":       Sector | Solar
Sheet "Location":      Country | Vietnam
```

**The system reads ALL formats automatically!**

## 🏷️ Template Placeholder Formats

### 1. Direct Placeholders
```
Company: [Company Name]
Location: [Country]
```

### 2. Tag Format (for row-based Excel)
```
Company: :COMPANY_NAME:
Location: :LOCATION_COUNTRY:
Sector: :SECTOR:
```

### 3. AI-Generated Content

**3-Paragraph Project Highlight:**
```
[AI: Write a 3-paragraph Project Highlight section for a {Sector} sector project in {Country} called "{Project_Focus}". 

Paragraph 1: What does the company do? Use: "{Company_Operations}"

Paragraph 2: What is the project scope and partnerships? Use: "{Service_Scope}". Also mention they "{Partnership}" and "{Partnership_Benefit}".

Paragraph 3: What is the investment plan? Use: "The project will commence with an initial investment of {Initial_Investment}, with a projected expansion to {Expansion_Investment} {Expansion_Path}."]
```

**2-3 Word Sector Label:**
```
[AI: Given Sector = "{Sector}" and Project_Type = "{Project_Type}", generate a 1-3 word sector label (no company name), e.g. "Solar Energy", "Wind Energy".]
```

**Anonymous Client One-Liner:**
```
[AI: Using Asset_Description = "{Asset_Description}" and Country = "{Country}", write one anonymous sentence starting with "The client" that briefly states what is being developed/operated and where.]
```

## 🛠️ Technical Details

**Requirements:**
- Python 3.8+
- No external APIs required
- Works offline after initial setup

**Dependencies** (auto-installed):
- `streamlit` — Web interface
- `pandas` — Excel processing
- `python-pptx` — PowerPoint manipulation
- `openpyxl` — Excel file reading

## 🎓 Advanced Features

### Fuzzy Label Matching
Labels don't need to match exactly:
- "Company name", "Company Name", "Name of company" → all work!
- 60% similarity threshold for automatic matching

### Sheet Namespacing (Multi-Sheet)
Optionally prefix fields with sheet names:
- `financials.total_cost`
- `company_info.sector`
- `project.country`

### Supported AI Patterns
1. **Project Highlights** - 3-paragraph structured content
2. **Sector Labels** - "Solar Energy", "Wind Energy", etc.
3. **Client One-Liners** - Anonymous project descriptions
4. **Custom Prompts** - Flexible length and style control

### Security
- .xlsm files supported (macros completely disabled)
- No code execution from Excel
- Safe data-only reading with openpyxl

## 📁 Test Files Included

- `example1_battery_japan.xlsx` - Battery storage project
- `example2_solar_vietnam.xlsx` - Solar farm project
- `example3_wind_india.xlsx` - Wind energy project
- `multi_sheet_example.xlsx` - Multi-sheet demo
- `test_template_examples.pptx` - Template with all prompt types

## 🤝 Contributing

Feel free to open issues or submit PRs to improve the generator!