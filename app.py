import streamlit as st
import pandas as pd
from pptx import Presentation
from io import BytesIO
import re
from difflib import get_close_matches

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
- **Column-based Excel:**  
  In PowerPoint: `Title: [Title]` or `Company: [Company Name]`  
  In Excel: columns `Title`, `Company Name`

- **Row-based Excel:**  
  In PowerPoint: `:COMPANY_NAME:`, `:TITLE:`, `:SPONSOR_NAME:`  
  In Excel: rows with labels like "Company name", "Title", "Sponsor name"

**2. AI-generated prompts (local)**  
- In PowerPoint:  
  `[AI: Write a teaser about {Company} in {Industry}]`  
  or  
  `[AI: Write about {COMPANY_NAME} in {LOCATION_COUNTRY}]`  
- Column-based: tokens match Excel columns  
- Row-based: tokens match canonical tags (COMPANY_NAME, etc.)

**Row-based format** automatically handles label variations:  
- "Company name", "Company Name", "Name of company" → all work!  
- Uses fuzzy matching for flexibility

Text box formatting is preserved when content is replaced.
"""
)


# ---------------- Smart local text generation ----------------
def normalize_label(label: str) -> str:
    """Normalize Excel label for consistent matching."""
    if not isinstance(label, str):
        return ""
    # Lowercase, strip spaces, remove punctuation
    normalized = label.lower().strip()
    # Remove common punctuation but keep underscores
    normalized = re.sub(r'[^\w\s]', '', normalized)
    # Replace multiple spaces with single space
    normalized = re.sub(r'\s+', ' ', normalized)
    return normalized


def detect_excel_format(df: pd.DataFrame) -> str:
    """
    Detect if Excel is column-based or row-based format.
    
    Returns:
        'column-based': Traditional format where each column is a field
        'row-based': Label-Value pair format where rows contain field definitions
    """
    # Check if first two columns look like Label-Value pairs
    if len(df.columns) >= 2:
        first_col = df.columns[0]
        second_col = df.columns[1]
        
        # Common indicators of row-based format
        label_indicators = ['label', 'field', 'name', 'key', 'question', 'item']
        value_indicators = ['value', 'answer', 'data', 'content', 'response']
        
        first_lower = str(first_col).lower()
        second_lower = str(second_col).lower()
        
        # Check column names
        is_label_value = (
            any(ind in first_lower for ind in label_indicators) and
            any(ind in second_lower for ind in value_indicators)
        )
        
        if is_label_value:
            return 'row-based'
        
        # Check content: if first column has question-like text
        if len(df) > 0:
            first_values = df.iloc[:, 0].dropna().astype(str).head(5)
            # Check if values look like labels (contain spaces, question marks, colons)
            label_like_count = sum(
                1 for v in first_values 
                if len(v) > 10 and (' ' in v or ':' in v or '?' in v)
            )
            if label_like_count >= 3:  # If 3+ rows look like labels
                return 'row-based'
    
    # Default to column-based
    return 'column-based'


def parse_row_based_excel(df: pd.DataFrame) -> dict:
    """
    Parse row-based Excel format into a normalized key-value dictionary.
    
    Expects Excel with structure:
    - Column 0: Labels/Questions
    - Column 1: Values/Answers
    - Optional: Column 2+ for categories/sections
    
    Returns:
        dict: {normalized_label: value}
    """
    data_dict = {}
    
    # Use first two columns as label-value
    label_col = df.columns[0]
    value_col = df.columns[1] if len(df.columns) > 1 else df.columns[0]
    
    for idx, row in df.iterrows():
        label = row[label_col]
        value = row[value_col]
        
        # Skip empty labels
        if pd.isna(label) or str(label).strip() == '':
            continue
        
        # Normalize the label
        normalized = normalize_label(str(label))
        
        # Store the value (convert NaN to empty string)
        data_dict[normalized] = "" if pd.isna(value) else str(value)
    
    return data_dict


def fuzzy_match_label(target: str, available_labels: list, threshold: float = 0.6) -> str:
    """
    Find the best matching label using fuzzy string matching.
    
    Args:
        target: The label to match
        available_labels: List of available labels
        threshold: Minimum similarity score (0-1)
    
    Returns:
        Best matching label or None
    """
    normalized_target = normalize_label(target)
    
    # Try exact match first
    if normalized_target in available_labels:
        return normalized_target
    
    # Try fuzzy matching
    matches = get_close_matches(normalized_target, available_labels, n=1, cutoff=threshold)
    
    if matches:
        return matches[0]
    
    return None


def create_placeholder_mapping() -> dict:
    """
    Define the mapping between PPT placeholder tags and canonical field names.
    
    This allows flexible wording in Excel while maintaining consistent PPT templates.
    
    Returns:
        dict: {PPT_TAG: [list of possible Excel label variations]}
    """
    return {
        'COMPANY_NAME': ['company name', 'company', 'name of company', 'issuer', 'issuer name'],
        'SPONSOR_NAME': ['sponsor name', 'sponsor', 'lead sponsor', 'project sponsor'],
        'LOCATION_COUNTRY': ['location country', 'country', 'jurisdiction', 'location'],
        'LOCATION_STATE': ['location state', 'state', 'province', 'region'],
        'ASSET_TYPE': ['asset type', 'type of asset', 'asset class', 'sector'],
        'PROJECT_TYPE': ['project type', 'type of project', 'development type'],
        'UNIT': ['unit', 'capacity', 'size', 'scale'],
        'TECHNOLOGY': ['technology', 'tech', 'technology type', 'technical approach'],
        'INITIAL_INVESTMENT': ['initial investment', 'initial capex', 'phase 1 investment', 'startup cost'],
        'FUTURE_INVESTMENT': ['future investment', 'expansion investment', 'phase 2 investment', 'growth capex'],
        'CONTRACTOR': ['contractor', 'epc contractor', 'construction partner', 'builder'],
        'OFFTAKER': ['offtaker', 'offtake partner', 'purchaser', 'buyer'],
        'PROJECT_STATUS': ['project status', 'status', 'development stage', 'phase'],
        'COD': ['cod', 'commercial operation date', 'operational date', 'completion date'],
        'DESCRIPTION': ['description', 'project description', 'overview', 'summary'],
        'TITLE': ['title', 'project title', 'project name', 'deal name'],
        'INDUSTRY': ['industry', 'sector', 'vertical', 'market'],
    }


# ---------------- Smart local text generation ----------------
def extract_structured_prompt_data(prompt_text: str) -> dict:
    """
    Extract structured data from prompts that follow the pattern:
    'Paragraph 1: ... Use: "{value}"'
    
    Returns dict of extracted values and their context.
    """
    data = {}
    
    # Pattern: Use: "{value}" or use "{value}"
    use_patterns = re.findall(r'[Uu]se:\s*["\']([^"\']+)["\']', prompt_text)
    
    # Map them to paragraph numbers if available
    paragraphs = re.split(r'Paragraph \d+:', prompt_text)
    
    for idx, para in enumerate(paragraphs[1:], 1):  # Skip first split (before any paragraph)
        # Extract all quoted strings in this paragraph
        quoted = re.findall(r'["\']([^"\']{10,})["\']', para)  # At least 10 chars to avoid short labels
        if quoted:
            data[f'paragraph_{idx}'] = quoted
    
    # Also extract all quoted strings globally as fallback
    all_quoted = re.findall(r'["\']([^"\']{10,})["\']', prompt_text)
    data['all_values'] = all_quoted
    
    return data


def generate_beta_text(prompt_text: str, row_data: pd.Series, tone: str = "short") -> str:
    """
    Generates text by interpreting the prompt after {token} substitution.
    Now with sophisticated template matching for structured content.
    """
    
    prompt_lower = prompt_text.lower()
    
    # **NEW: Detect 3-paragraph "Project Highlight" structure**
    # This is the most common pattern in the user's examples
    is_project_highlight = (
        'paragraph 1:' in prompt_lower and 
        'paragraph 2:' in prompt_lower and 
        'paragraph 3:' in prompt_lower
    )
    
    if is_project_highlight:
        # Extract structured data from the prompt
        data = extract_structured_prompt_data(prompt_text)
        paragraphs = []
        
        # **PARAGRAPH 1 FORMULA:**
        # Sentence 1: "The client {Company_Operations}."
        # Sentence 2: "The proposed project is the {Project_Focus}."
        if 'paragraph_1' in data and len(data['paragraph_1']) >= 1:
            company_ops = data['paragraph_1'][0]
            project_focus = data['paragraph_1'][1] if len(data['paragraph_1']) > 1 else ""
            
            # Handle "called" pattern: extract project focus from 'called "{Project_Focus}"'
            if not project_focus and 'called' in prompt_text:
                called_match = re.search(r'called\s+["\']([^"\']+)["\']', prompt_text, re.IGNORECASE)
                if called_match:
                    project_focus = called_match.group(1)
            
            para1_sent1 = f"The client {company_ops}."
            para1_sent2 = f"The proposed project is the {project_focus}." if project_focus else ""
            paragraphs.append(f"{para1_sent1} {para1_sent2}".strip())
        
        # **PARAGRAPH 2 FORMULA:**
        # Sentence 1: "The client {Service_Scope}."
        # Sentence 2: "They {Partnership}." (handle verb forms)
        # Sentence 3: "{Partnership_Benefit}."
        if 'paragraph_2' in data and len(data['paragraph_2']) >= 1:
            service_scope = data['paragraph_2'][0]
            partnership = data['paragraph_2'][1] if len(data['paragraph_2']) > 1 else ""
            partnership_benefit = data['paragraph_2'][2] if len(data['paragraph_2']) > 2 else ""
            
            para2_sent1 = f"The client {service_scope}."
            
            # Handle partnership verb form
            if partnership:
                partnership_clean = partnership.strip()
                if partnership_clean.startswith('partner with') or partnership_clean.startswith('partner in'):
                    para2_sent2 = f"They {partnership_clean}."
                elif partnership_clean.startswith('partnered with') or 'partnered' in partnership_clean:
                    para2_sent2 = f"They are {partnership_clean}."
                else:
                    para2_sent2 = f"They {partnership_clean}."
            else:
                para2_sent2 = ""
            
            para2_sent3 = f"{partnership_benefit}." if partnership_benefit else ""
            
            para2 = f"{para2_sent1} {para2_sent2} {para2_sent3}".strip()
            paragraphs.append(para2)
        
        # **PARAGRAPH 3 FORMULA:**
        # Single sentence: "The project will commence with an initial investment of {Initial}, 
        #                   with a projected expansion to {Expansion} {Path}."
        if 'paragraph_3' in data or 'all_values' in data:
            # Look for the investment sentence pattern directly in prompt
            investment_pattern = re.search(
                r'The project will commence with an initial investment of\s+([^,]+),\s*with a projected expansion to\s+([^.]+)',
                prompt_text,
                re.IGNORECASE
            )
            
            if investment_pattern:
                para3 = f"The project will commence with an initial investment of {investment_pattern.group(1)}, with a projected expansion to {investment_pattern.group(2)}."
                paragraphs.append(para3)
        
        # Join paragraphs with double newline
        result = "\n\n".join(p for p in paragraphs if p)
        if result:
            return result
    
    # **FALLBACK: Original logic for other prompt types**
    
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
def resolve_tag_to_value(tag: str, data_dict: dict, placeholder_mapping: dict) -> str:
    """
    Resolve a :TAG: placeholder to its value using the mapping and data dictionary.
    
    Args:
        tag: The tag name (e.g., 'COMPANY_NAME')
        data_dict: Dictionary of normalized_label -> value
        placeholder_mapping: Dictionary of TAG -> [possible label variations]
    
    Returns:
        The resolved value or empty string if not found
    """
    # Check if this tag is in our mapping
    if tag in placeholder_mapping:
        possible_labels = placeholder_mapping[tag]
        
        # Try each possible label variation
        for label_variation in possible_labels:
            normalized = normalize_label(label_variation)
            if normalized in data_dict:
                return data_dict[normalized]
        
        # If no exact match, try fuzzy matching
        fuzzy_match = fuzzy_match_label(possible_labels[0], list(data_dict.keys()))
        if fuzzy_match and fuzzy_match in data_dict:
            return data_dict[fuzzy_match]
    
    # Try using the tag itself as a label (for backwards compatibility)
    normalized_tag = normalize_label(tag)
    if normalized_tag in data_dict:
        return data_dict[normalized_tag]
    
    # Try fuzzy match on the tag
    fuzzy_match = fuzzy_match_label(tag, list(data_dict.keys()))
    if fuzzy_match and fuzzy_match in data_dict:
        return data_dict[fuzzy_match]
    
    return ""


def process_placeholder(
    placeholder: str,
    row_data: pd.Series,
    excel_columns: list,
    beta_tone: str,
    missing_to_blank: bool,
    data_dict: dict = None,
    placeholder_mapping: dict = None,
    excel_format: str = 'column-based',
) -> str:
    """
    Resolve a placeholder to final text using smart local generation.

    Supports multiple formats:
    - Direct: "Title" → row_data["Title"] (column-based)
    - Tag: ":COMPANY_NAME:" → resolved via mapping (row-based)
    - AI: "AI: Write teaser about {Company}" → smart text generation
    """
    # Handle :TAG: format (new row-based system)
    tag_match = re.match(r'^:([A-Z_]+):$', placeholder.strip())
    if tag_match and data_dict is not None and placeholder_mapping is not None:
        tag = tag_match.group(1)
        return resolve_tag_to_value(tag, data_dict, placeholder_mapping)
    
    # AI placeholder - always use local generation
    if placeholder.startswith("AI:"):
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
        
        # Also support {TAG} format for row-based data
        if data_dict is not None and placeholder_mapping is not None:
            # Find all {TAG} or {{TAG}} patterns
            tokens = re.findall(r'\{\{?([A-Z_]+)\}?\}?', prompt_text)
            for token in tokens:
                value = resolve_tag_to_value(token, data_dict, placeholder_mapping)
                prompt_text = prompt_text.replace(f"{{{{{token}}}}}", value)
                prompt_text = prompt_text.replace(f"{{{token}}}", value)
        
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

    # Detect Excel format
    excel_format = detect_excel_format(df)
    st.info(f"📋 **Detected Excel format:** {excel_format}")
    
    # Allow user to override
    excel_format = st.radio(
        "Excel format",
        options=['column-based', 'row-based'],
        index=0 if excel_format == 'column-based' else 1,
        help=(
            "**Column-based**: Each column is a field (e.g., 'Company Name' column)\n\n"
            "**Row-based**: Each row has Label and Value (e.g., 'Company name' | 'CLimAIte')"
        ),
        horizontal=True,
    )
    
    # Parse data based on format
    if excel_format == 'row-based':
        data_dict = parse_row_based_excel(df)
        placeholder_mapping = create_placeholder_mapping()
        
        st.success(f"✅ Parsed {len(data_dict)} fields from row-based Excel")
        
        with st.expander("🔍 View parsed fields"):
            for key, value in list(data_dict.items())[:10]:
                st.write(f"**{key}**: {value[:100] if value else '(empty)'}...")
            if len(data_dict) > 10:
                st.write(f"... and {len(data_dict) - 10} more fields")
        
        # For row-based, we don't need to select a row
        row_data = None
        excel_columns = []
    else:
        # Column-based: use existing logic
        data_dict = None
        placeholder_mapping = None
        excel_columns = df.columns.tolist()

    # Load PPTX template
    try:
        prs = Presentation(template_file)
    except Exception as e:
        st.error(f"Error loading PPTX template: {e}")
        st.stop()

    st.divider()
    st.subheader("⚙️ Generation settings")

    # Choose row from Excel (only for column-based)
    if excel_format == 'column-based':
        row_index = st.selectbox(
            "Select Excel row to use",
            options=range(len(df)),
            format_func=lambda i: f"Row {i+1}",
        )
        row_data = df.iloc[row_index]
    else:
        row_index = None
        st.info("ℹ️ Row-based format: Using all label-value pairs from Excel")

    # Button to generate
    if st.button("🚀 Generate PPTX", type="primary"):
        # Ensure we have the data we need
        if excel_format == 'column-based':
            excel_columns = df.columns.tolist()
            row_data = df.iloc[row_index]
        else:
            # Already set above
            pass

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
                                # Also find :TAG: format placeholders
                                tag_matches = re.findall(r":([A-Z_]+):", original_text)
                                direct_matches.extend([f":{tag}:" for tag in tag_matches])

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
                                        data_dict,
                                        placeholder_mapping,
                                        excel_format,
                                    )
                                    new_text = new_text.replace(f"[{placeholder}]", replacement)
                                    # Also handle :TAG: format
                                    if placeholder.startswith(':') and placeholder.endswith(':'):
                                        new_text = new_text.replace(placeholder, replacement)

                                for ai_prompt in ai_matches:
                                    status.text(f"🤖 Generating text: {ai_prompt[:60]}...")
                                    replacement = process_placeholder(
                                        f"AI: {ai_prompt}",
                                        row_data,
                                        excel_columns,
                                        beta_tone,
                                        missing_to_blank,
                                        data_dict,
                                        placeholder_mapping,
                                        excel_format,
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
                    # Also find :TAG: format placeholders
                    tag_matches = re.findall(r":([A-Z_]+):", original_text)
                    direct_matches.extend([f":{tag}:" for tag in tag_matches])

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
                            data_dict,
                            placeholder_mapping,
                            excel_format,
                        )
                        new_text = new_text.replace(f"[{placeholder}]", replacement)
                        # Also handle :TAG: format
                        if placeholder.startswith(':') and placeholder.endswith(':'):
                            new_text = new_text.replace(placeholder, replacement)

                    for ai_prompt in ai_matches:
                        status.text(f"🤖 Generating text: {ai_prompt[:60]}...")
                        replacement = process_placeholder(
                            f"AI: {ai_prompt}",
                            row_data,
                            excel_columns,
                            beta_tone,
                            missing_to_blank,
                            data_dict,
                            placeholder_mapping,
                            excel_format,
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