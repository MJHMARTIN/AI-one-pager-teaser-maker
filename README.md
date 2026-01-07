# AI One Pager Teaser Maker

This is a Streamlit app that generates PowerPoint presentations (PPTX) with AI-generated teasers using DeepSeek API.

## Features

- Upload a PPTX template with placeholders like `[Title]` or AI prompts like `[AI: Write a teaser about {Company}]`
- Upload Excel data with columns matching the placeholders
- AI generates text using DeepSeek API for prompts
- Preserves text formatting (font, size, color, bold, italic)
- Download the filled PPTX

## Setup

1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

2. Run the app:
   ```bash
   streamlit run app.py
   ```

3. Open the URL shown (usually http://localhost:8501)

## Usage

1. Enter your DeepSeek API key in the sidebar (get one from https://platform.deepseek.com/)
2. Upload a PPTX template file
3. Upload an Excel (.xlsx) file with data
4. Select a row from the Excel data
5. Click "Generate PPTX" to create the filled presentation
6. Download the result

## Template Syntax

- Direct placeholders: `[Column Name]` - replaced with Excel column value
- AI prompts: `[AI: Your prompt here {Column Name}]` - generates text using DeepSeek, with {Column} replaced from Excel

## Requirements

- Python 3.8+
- DeepSeek API key
- PPTX template file
- Excel data file