# 📊 AI One-Pager Teaser Maker

**Generate professional PowerPoint presentations with AI-powered content — no API key required!**

This Streamlit web app automatically fills PPTX templates with data from Excel files and generates compelling teaser text using smart local AI. Perfect for creating investor one-pagers, deal teasers, company profiles, and marketing materials at scale.

## ✨ Key Features

- **🤖 Local AI Generation** — No API keys, no costs, completely free smart text generation
- **📝 Template System** — Use direct placeholders `[Title]` or AI prompts `[AI: Write a teaser about {Company}]`
- **📊 Excel Integration** — Bulk process multiple deals/companies from spreadsheet data
- **🎨 Format Preservation** — Maintains all PowerPoint formatting (fonts, colors, bold, italic)
- **⚡ Batch Processing** — Generate multiple presentations from rows of data
- **🔧 Flexible AI Prompts** — Control tone (short/medium/long), style, and structure
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

1. **PPTX Template** — Your PowerPoint template with placeholders and/or AI prompts
2. **Excel File** — Spreadsheet with data columns matching your placeholders
3. **That's it!** — No API keys or external services required

## 📖 How to Use

1. **Upload Template** — Click "Browse files" to upload your PPTX template
2. **Upload Data** — Upload an Excel (.xlsx) file with your data
3. **Select Row** — Choose which row from your Excel to process
4. **Configure AI** — Set the AI tone (short/medium/long) in the sidebar
5. **Generate** — Click "Generate PPTX" to create your presentation
6. **Download** — Get your filled PowerPoint file instantly

## 🎯 Template Syntax Examples

### Direct Placeholders (Simple Replacement)
In your PowerPoint template:
```
Company Name: [Company]
Industry: [Industry]
Location: [City], [State]
```

In your Excel file, have columns named: `Company`, `Industry`, `City`, `State`

### AI-Generated Content (Smart Text)
In your PowerPoint template:
```
[AI: Write a professional 2-3 sentence teaser about {Company}, a company in the {Industry} sector located in {City}]
```

Or with double braces:
```
[AI: Describe {{ISSUER}}'s competitive advantage in {{JURISDICTION}}]
```

The AI will:
- Replace `{Company}`, `{Industry}`, etc. with data from Excel
- Generate contextual, professional text based on the prompt
- Follow length requirements (e.g., "2-3 sentences", "50 words")
- Match the requested tone and style

### Supported AI Prompt Features
- **Length control**: "Write 2 sentences", "50 words", "1 paragraph"
- **Tone selection**: Choose short/medium/long in the sidebar
- **Style keywords**: "professional", "technical", "formal" in your prompt
- **Structure**: "Follow this structure: Sentence 1: [topic]. Sentence 2: [detail]"

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