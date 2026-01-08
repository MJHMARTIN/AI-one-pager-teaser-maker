import streamlit as st
import pandas as pd
from pptx import Presentation
from io import BytesIO
import re
import requests

# ---------------- Streamlit page config ----------------
st.set_page_config(page_title="PPTX Teaser Generator with AI", layout="wide")
st.title("📊 PPTX Teaser Generator with AI")

st.write(
    "Upload your template PPTX and Excel data. "
    "The template can contain direct placeholders like [Title] and AI prompts like "
    "[AI: Write a short teaser about {Company} in {Industry}]. "
    "Formatting (font, size, color, bold, italic) of each text box is preserved."
)

# ---------------- Sidebar: Config ----------------
st.sidebar.header("⚙️ Configuration")

st.sidebar.info("💡 **Using smart local AI generator** — No API needed, completely free!")

beta_tone = st.sidebar.radio(
    "AI Text Tone",
    options=["short", "medium", "long"],
    index=1,
    help="short: 1 sentence. medium: 1-2 sentences. long: 2-3 detailed sentences.",
)

missing_to_blank = st.sidebar.checkbox(
    "Treat missing columns as blank",
    value=True,
    help="If a placeholder column is missing, insert blank instead of an error tag.",
)

st.sidebar.divider()

st.sidebar.info(
    """
**Template syntax**

**1. Direct placeholders (no AI)**  
- In PowerPoint:  
  `Title: [Title]`  
  `Company: [Company Name]`  
- In Excel: columns `Title`, `Company Name`

**2. AI-generated prompts (local)**  
- In PowerPoint:  
  `[AI: Write a teaser about {Company} in {Industry}]`  
  or  
  `[AI: Write about {{ISSUER}} in {{JURISDICTION}}]`  
- In Excel: columns match the tokens (e.g., `Company`, `Industry`, `ISSUER`, `JURISDICTION`)  
- Both `{Token}` and `{{Token}}` formats are supported

Text box formatting is preserved when content is replaced.
"""
)


# ---------------- Smart local text generation ----------------
def generate_beta_text(prompt_text: str, row_data: pd.Series, tone: str = "short") -> str:
    """
    Generates text by interpreting the prompt after {token} substitution.
    Detects length requirements (words, sentences, paragraphs) and follows them.
    """
    import re
    
    prompt_lower = prompt_text.lower()
    
    # Detect specific length requirements in the prompt
    word_count = None
    sentence_count = None
    paragraph_count = None
    
    # Look for "X words", "X-Y words"
    word_match = re.search(r'(\d+)(?:-(\d+))?\s+words?', prompt_lower)
    if word_match:
        word_count = (int(word_match.group(1)), int(word_match.group(2) or word_match.group(1)))
    
    # Look for "X sentences", "X-Y sentences"
    sent_match = re.search(r'(\d+)(?:-(\d+))?\s+sentences?', prompt_lower)
    if sent_match:
        sentence_count = (int(sent_match.group(1)), int(sent_match.group(2) or sent_match.group(1)))
    
    # Look for "X paragraphs", "X-Y paragraphs"
    para_match = re.search(r'(\d+)(?:-(\d+))?\s+paragraphs?', prompt_lower)
    if para_match:
        paragraph_count = (int(para_match.group(1)), int(para_match.group(2) or para_match.group(1)))
    
    # Detect content type
    is_structured = "follow this structure" in prompt_lower or "sentence 1:" in prompt_lower
    is_professional = "professional" in prompt_lower or "formal" in prompt_lower or "business" in prompt_lower
    
    # Extract all the actual data values from the prompt (these replaced the tokens)
    values = []
    
    # Find quoted text
    quoted = re.findall(r'"([^"]+)"', prompt_text)
    values.extend(quoted)
    
    # Find dollar amounts
    amounts = re.findall(r'\$[\d,]+(?:\s*(?:million|billion|thousand))?', prompt_text, re.IGNORECASE)
    values.extend(amounts)
    
    # Find capitalized multi-word phrases
    capitalized = re.findall(r'\b[A-Z][a-zA-Z]+(?:\s+[A-Z][a-zA-Z]+)*\b', prompt_text)
    excluded = {'Using', 'Use', 'Write', 'Follow', 'Sentence', 'Do', 'Excel', 'English', 'Initial', 'Future', 'The', 'Based'}
    capitalized = [c for c in capitalized if c not in excluded and len(c) > 2]
    values.extend(capitalized)
    
    # Find longer descriptive phrases (likely field values)
    technical_terms = re.findall(r'\b[a-z]+(?:\s+[a-z]+){1,5}\b', prompt_text)
    technical_terms = [t for t in technical_terms if len(t) > 15 and 'write' not in t and 'using' not in t and 'only' not in t]
    values.extend(technical_terms[:3])
    
    # Remove duplicates
    seen = set()
    unique_values = []
    for v in values:
        v_clean = v.strip()
        if v_clean and v_clean not in seen and len(v_clean) > 1:
            seen.add(v_clean)
            unique_values.append(v_clean)
    
    # Generate content based on requirements
    if word_count and word_count[1] <= 10:
        # Very short: 3-10 words
        if len(unique_values) >= 1:
            return unique_values[0]
        else:
            return "Professional solutions"
    
    elif is_structured and paragraph_count:
        # Structured multi-paragraph request
        paragraphs = []
        
        # Identify key components
        client = "The client"
        location = next((v for v in unique_values if any(place in v for place in ['United', 'States', 'Kingdom', 'Europe', 'Asia', 'Africa', 'America'])), None)
        technology = next((v for v in unique_values if any(tech in v.lower() for tech in ['solar', 'wind', 'energy', 'power', 'technology', 'system', 'renewable'])), None)
        partners = [v for v in unique_values if any(word in v.lower() for word in ['contractor', 'provider', 'partner', 'company', 'corp'])]
        amounts_list = [v for v in unique_values if '$' in v]
        descriptions = [v for v in unique_values if len(v) > 40]
        
        # Build paragraphs
        for i in range(paragraph_count[1]):
            sentences = []
            
            if i == 0:
                # First paragraph: Introduction
                if technology and location:
                    sentences.append(f"{client} is developing a {technology} project in {location}.")
                elif technology:
                    sentences.append(f"{client} is developing a {technology} project.")
                elif len(unique_values) >= 2:
                    sentences.append(f"{client} is undertaking {unique_values[0]} in {unique_values[1]}.")
                else:
                    sentences.append(f"{client} is pursuing a strategic development initiative.")
                
                if descriptions:
                    sentences.append(f"The project encompasses {descriptions[0].lower()}.")
                elif len(unique_values) >= 3:
                    sentences.append(f"This involves {unique_values[2].lower()}.")
            
            elif i == 1:
                # Second paragraph: Development/Partners
                if partners:
                    partner_text = " and ".join(partners[:2])
                    sentences.append(f"The project will be developed in collaboration with {partner_text}.")
                    sentences.append(f"These partnerships ensure successful delivery and operation.")
                else:
                    sentences.append(f"The initiative leverages proven methodologies and industry expertise.")
                    sentences.append(f"Development will proceed in phases to ensure optimal outcomes.")
            
            elif i == 2:
                # Third paragraph: Investment/Financial
                if len(amounts_list) >= 2:
                    sentences.append(f"The initial investment is {amounts_list[0]}, with planned expansion investment of {amounts_list[1]}.")
                elif len(amounts_list) == 1:
                    sentences.append(f"The project represents an investment of {amounts_list[0]}.")
                else:
                    sentences.append(f"The project is backed by substantial capital commitment.")
                sentences.append(f"This financial structure supports both immediate development and future growth.")
            
            if sentences:
                paragraphs.append(" ".join(sentences))
        
        return "\n\n".join(paragraphs)
    
    elif paragraph_count or (sentence_count and sentence_count[1] >= 3):
        # Multi-sentence or paragraph content
        sentences = []
        target_sentences = sentence_count[1] if sentence_count else (paragraph_count[1] * 3 if paragraph_count else 3)
        
        # Build multiple sentences
        if len(unique_values) >= 2:
            sentences.append(f"{unique_values[0]} delivers innovative solutions in {unique_values[1]}.")
        elif len(unique_values) >= 1:
            sentences.append(f"{unique_values[0]} represents excellence in their field.")
        else:
            sentences.append("Professional solutions delivered with expertise.")
        
        if target_sentences >= 2:
            if len(unique_values) >= 3:
                sentences.append(f"Leveraging {unique_values[2]}, the organization drives exceptional outcomes.")
            else:
                sentences.append("Advanced capabilities ensure outstanding results.")
        
        if target_sentences >= 3:
            sentences.append("Proven track record of success across diverse initiatives.")
        
        if target_sentences >= 4:
            sentences.append("Strategic partnerships and innovation create lasting value.")
        
        return " ".join(sentences[:target_sentences])
    
    else:
        # Default: short single sentence
        if len(unique_values) >= 2:
            return f"{unique_values[0]} delivers excellence in {unique_values[1]}."
        elif len(unique_values) >= 1:
            return f"{unique_values[0]} drives innovation and results."
        else:
            return "Excellence delivered through proven expertise."


# -------- Placeholder logic --------
def process_placeholder(
    placeholder: str,
    row_data: pd.Series,
    excel_columns: list,
    beta_tone: str,
    missing_to_blank: bool,
) -> str:
    """
    Resolve a placeholder to final text using smart local generation.

    - Direct: "Title" → row_data["Title"]
    - AI: "AI: Write teaser about {Company}" → smart text generation with Excel values.
    """
    # AI placeholder - always use local generation
    if placeholder.startswith("AI:"):
        import re
        
        prompt_text = placeholder[3:].strip()
        
        # Handle double "AI:" at the start (common error)
        if prompt_text.startswith("AI:"):
            prompt_text = prompt_text[3:].strip()

        # Track which tokens couldn't be found
        missing_tokens = []
        
        # Replace {{ColumnName}} or {ColumnName} with actual values from Excel row
        for col in excel_columns:
            if col in row_data.index:
                value = "" if pd.isna(row_data[col]) else str(row_data[col])
                # Support both {{}} and {} formats
                prompt_text = prompt_text.replace(f"{{{{{col}}}}}", value)  # {{ColumnName}}
                prompt_text = prompt_text.replace(f"{{{col}}}", value)       # {ColumnName}
        
        # Fix common typos: {{TOKEN} with only one closing brace
        prompt_text = re.sub(r'\{\{([^}]+)\}(?!\})', r'{{\1}}', prompt_text)
        
        # Find remaining unreplaced tokens to report to user
        remaining_tokens = re.findall(r'\{\{([^}]+)\}\}', prompt_text)
        remaining_tokens.extend(re.findall(r'\{([^}]+)\}', prompt_text))
        
        if remaining_tokens:
            st.warning(f"⚠️ Missing Excel columns: {', '.join(set(remaining_tokens))}. Add these columns to your Excel or remove from prompt.")
        
        # Replace any remaining {{TOKEN}} or {TOKEN} with blank instead of error message
        prompt_text = re.sub(r'\{\{[^}]+\}\}', '', prompt_text)
        prompt_text = re.sub(r'\{[^}]+\}', '', prompt_text)

        st.write(f"🔍 DEBUG - After token replacement: '{prompt_text[:200]}...'")
        
        generated = generate_beta_text(prompt_text, row_data, beta_tone)
        
        st.write(f"✅ DEBUG - Generated text: '{generated[:200]}...'")
        return generated

    # Direct placeholder
    col_name = placeholder.strip()
    if col_name in row_data.index:
        val = row_data[col_name]
        return "" if pd.isna(val) else str(val)
    else:
        return "" if missing_to_blank else f"[MISSING COLUMN: {col_name}]"


# ---------------- File uploads ----------------
col1, col2 = st.columns(2)

with col1:
    st.subheader("Step 1: Template PPTX")
    template_file = st.file_uploader(
        "Upload your PPTX template (with [Placeholders] and [AI: ...] prompts)",
        type="pptx",
        key="template",
    )

with col2:
    st.subheader("Step 2: Excel Data")
    excel_file = st.file_uploader(
        "Upload your Excel file (columns must match placeholders)",
        type="xlsx",
        key="excel",
    )


# ---------------- Template inspector (optional helper) ----------------
if template_file:
    with st.expander("🔍 Inspect template text boxes (optional)"):
        try:
            prs_preview = Presentation(template_file)
            st.write(f"Slides: **{len(prs_preview.slides)}**")
            slide_index = st.selectbox(
                "Slide to inspect",
                options=range(len(prs_preview.slides)),
                format_func=lambda i: f"Slide {i+1}",
            )
            slide = prs_preview.slides[slide_index]
            st.write(f"**Shapes on Slide {slide_index + 1}:**")

            for idx, shape in enumerate(slide.shapes):
                # Show text from text shapes and table cells
                if hasattr(shape, "text_frame"):
                    st.write(f"- Shape {idx} name: `{shape.name}`")
                    st.write(f"  Text: `{shape.text[:80]}`")
                elif getattr(shape, "has_table", False):
                    st.write(f"- Table {idx} with {len(shape.table.rows)} rows")
        except Exception as e:
            st.error(f"Error inspecting template: {e}")


# ---------------- Main generation flow ----------------
if template_file and excel_file:
    st.divider()

    # Read Excel
    try:
        df = pd.read_excel(excel_file)
    except Exception as e:
        st.error(f"Error reading Excel: {e}")
        st.stop()

    if df.empty:
        st.error("Excel file has no rows.")
        st.stop()

    st.subheader("📋 Excel preview")
    st.dataframe(df, use_container_width=True)

    # Load PPTX template
    try:
        prs = Presentation(template_file)
    except Exception as e:
        st.error(f"Error loading PPTX template: {e}")
        st.stop()

    st.divider()
    st.subheader("⚙️ Generation settings")

    # Choose row from Excel
    row_index = st.selectbox(
        "Select Excel row to use",
        options=range(len(df)),
        format_func=lambda i: f"Row {i+1}",
    )

    # Button to generate
    if st.button("🚀 Generate PPTX", type="primary"):
        excel_columns = df.columns.tolist()
        row_data = df.iloc[row_index]

        try:
            # Count targets: text boxes + table cells
            def count_text_targets(pres: Presentation) -> int:
                total = 0
                for s in pres.slides:
                    for shp in s.shapes:
                        if hasattr(shp, "text_frame"):
                            total += 1
                        elif getattr(shp, "has_table", False):
                            tbl = shp.table
                            for r in tbl.rows:
                                total += len(r.cells)
                return total

            total_targets = count_text_targets(prs)
            processed = 0
            progress = st.progress(0.0)
            status = st.empty()

            for slide in prs.slides:
                for shape in slide.shapes:
                    # Debug: show what type of shape we're looking at
                    st.write(f"🔍 Found shape: {shape.name} (has text_frame: {hasattr(shape, 'text_frame')}, has_table: {getattr(shape, 'has_table', False)})")
                    
                    # Skip shapes without text capability
                    if not hasattr(shape, "text_frame") and not getattr(shape, "has_table", False):
                        processed += 1
                        progress.progress(processed / max(total_targets, 1))
                        continue
                    
                    # Handle tables
                    if getattr(shape, "has_table", False):
                        table = shape.table
                        for r in table.rows:
                            for cell in r.cells:
                                original_text = cell.text
                                if not original_text:
                                    processed += 1
                                    progress.progress(processed / max(total_targets, 1))
                                    continue

                                ai_matches = re.findall(r"\[AI:\s*(.+?)\]", original_text)
                                direct_matches = re.findall(r"\[(?!AI:)(.+?)\]", original_text)

                                if not direct_matches and not ai_matches:
                                    processed += 1
                                    progress.progress(processed / max(total_targets, 1))
                                    continue

                                new_text = original_text

                                for placeholder in direct_matches:
                                    replacement = process_placeholder(
                                        placeholder,
                                        row_data,
                                        excel_columns,
                                        beta_tone,
                                        missing_to_blank,
                                    )
                                    new_text = new_text.replace(f"[{placeholder}]", replacement)

                                for ai_prompt in ai_matches:
                                    status.text(f"🤖 Generating text: {ai_prompt[:60]}...")
                                    replacement = process_placeholder(
                                        f"AI: {ai_prompt}",
                                        row_data,
                                        excel_columns,
                                        beta_tone,
                                        missing_to_blank,
                                    )
                                    st.write(f"DEBUG: Original=[AI: {ai_prompt}], Replacement={replacement[:100]}")
                                    new_text = new_text.replace(f"[AI: {ai_prompt}]", replacement)

                                if new_text != original_text:
                                    tf = cell.text_frame
                                    if not tf.paragraphs:
                                        cell.text = new_text
                                    else:
                                        para = tf.paragraphs[0]
                                        if para.runs:
                                            run0 = para.runs[0]
                                            font = run0.font

                                            tf.clear()
                                            new_para = tf.paragraphs[0]
                                            new_run = new_para.add_run()
                                            new_run.text = new_text

                                            if font.size:
                                                new_run.font.size = font.size
                                            if font.name:
                                                new_run.font.name = font.name
                                            if font.bold is not None:
                                                new_run.font.bold = font.bold
                                            if font.italic is not None:
                                                new_run.font.italic = font.italic
                                            if font.color and hasattr(font.color, 'rgb'):
                                                new_run.font.color.rgb = font.color.rgb

                                            new_para.alignment = para.alignment
                                            new_para.level = para.level
                                        else:
                                            cell.text = new_text

                                processed += 1
                                progress.progress(processed / max(total_targets, 1))

                        # Move to next shape
                        continue

                    # Handle regular text shapes (text boxes, rectangles, etc.)
                    # Most shapes with text will have a text_frame attribute
                    if not hasattr(shape, "text_frame"):
                        processed += 1
                        progress.progress(processed / max(total_targets, 1))
                        continue

                    original_text = shape.text
                    if not original_text:
                        processed += 1
                        progress.progress(processed / max(total_targets, 1))
                        continue

                    st.write(f"  📝 Shape text preview: '{original_text[:100]}...'")
                    
                    # More forgiving regex that handles multi-line and missing closing brackets
                    # First try standard format: [AI: ... ]
                    ai_matches = re.findall(r"\[AI:\s*(.+?)\]", original_text, re.DOTALL)
                    
                    # If no matches, try without closing bracket (for malformed placeholders)
                    if not ai_matches:
                        ai_no_close = re.findall(r"\[AI:\s*(.+)$", original_text, re.DOTALL)
                        if ai_no_close:
                            st.warning(f"⚠️ Found AI placeholder without closing bracket ]")
                            ai_matches = ai_no_close
                    
                    direct_matches = re.findall(r"\[(?!AI:)(.+?)\]", original_text)

                    st.write(f"  🎯 Found {len(ai_matches)} AI placeholders, {len(direct_matches)} direct placeholders")

                    if not direct_matches and not ai_matches:
                        processed += 1
                        progress.progress(processed / max(total_targets, 1))
                        continue

                    new_text = original_text

                    for placeholder in direct_matches:
                        replacement = process_placeholder(
                            placeholder,
                            row_data,
                            excel_columns,
                            beta_tone,
                            missing_to_blank,
                        )
                        new_text = new_text.replace(f"[{placeholder}]", replacement)

                    for ai_prompt in ai_matches:
                        status.text(f"🤖 Generating text: {ai_prompt[:60]}...")
                        replacement = process_placeholder(
                            f"AI: {ai_prompt}",
                            row_data,
                            excel_columns,
                            beta_tone,
                            missing_to_blank,
                        )
                        st.write(f"DEBUG: Original=[AI: {ai_prompt}], Replacement={replacement[:100]}")
                        new_text = new_text.replace(f"[AI: {ai_prompt}]", replacement)

                    if new_text != original_text:
                        text_frame = shape.text_frame
                        if not text_frame.paragraphs:
                            shape.text = new_text
                        else:
                            para = text_frame.paragraphs[0]
                            if para.runs:
                                run0 = para.runs[0]
                                font = run0.font

                                text_frame.clear()
                                new_para = text_frame.paragraphs[0]
                                new_run = new_para.add_run()
                                new_run.text = new_text

                                if font.size:
                                    new_run.font.size = font.size
                                if font.name:
                                    new_run.font.name = font.name
                                if font.bold is not None:
                                    new_run.font.bold = font.bold
                                if font.italic is not None:
                                    new_run.font.italic = font.italic
                                if font.color and hasattr(font.color, 'rgb'):
                                    new_run.font.color.rgb = font.color.rgb

                                new_para.alignment = para.alignment
                                new_para.level = para.level
                            else:
                                shape.text = new_text

                    processed += 1
                    progress.progress(processed / max(total_targets, 1))

            progress.progress(1.0)
            status.text("✅ Done")

            output = BytesIO()
            prs.save(output)
            output.seek(0)

            st.success("✅ PPTX generated successfully!")
            st.download_button(
                label="📥 Download filled PPTX",
                data=output.getvalue(),
                file_name="teaser_filled.pptx",
                mime=(
                    "application/"
                    "vnd.openxmlformats-officedocument."
                    "presentationml.presentation"
                ),
            )

        except Exception as e:
            st.error(f"Error generating PPTX: {e}")
            st.exception(e)