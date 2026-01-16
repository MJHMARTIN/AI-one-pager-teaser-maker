# 📝 AI Template Functions Documentation

## Overview

The AI text generation system now uses **fixed template functions** for common patterns. This ensures consistent, predictable output with proper length control.

## Architecture

### Before (Old System)
- ❌ Inline detection mixed with generation logic
- ❌ Hard to test and maintain
- ❌ Inconsistent output length
- ❌ Duplicated code
- ❌ No quality validation

### After (New System)
- ✅ Dedicated template functions for each pattern
- ✅ Separate post-processing for length control
- ✅ **Built-in validation checks**
- ✅ **Automatic regeneration with stricter prompts**
- ✅ Testable and maintainable
- ✅ Consistent output quality

## Quality Control Features 🛡️

### Validation & Regeneration

The system includes **automatic quality checks** after text generation:

#### 1. Client One-liner Validation
**Checks:**
- ✅ Must start with "The client"
- ✅ Must be exactly one sentence
- ✅ Must maintain anonymity (no company names)

**Action:** If validation fails → automatic regeneration with stricter prompt

#### 2. Anonymity Validation
**Process:**
1. Extracts company names from source data
2. Checks if names appear in generated text
3. Prevents identity leaks

**Action:** If anonymity breached → automatic regeneration

#### 3. Stricter Prompt Generation
**When validation fails:**
1. Adds explicit instructions
2. Emphasizes format requirements
3. Regenerates once with enhanced guidance

**Example:**
```
Original: "Write one sentence about the client"
↓ (fails validation)
Stricter: "Write one sentence about the client
IMPORTANT: Start with exactly 'The client is'
Do NOT include any company names
Keep it to ONE sentence only"
```

---

## Available Templates

### 1. Sector Label Template

**Function:** `generate_sector_label(prompt_text, max_words=3)`

**Purpose:** Generate 2-3 word sector labels

**Trigger Patterns:**
```
- "sector label" + ("1-3 word" OR "2-3 word" OR "short")
- "short label" + word count requirement
```

**Example Input:**
```
Generate a 2-3 word sector label. Sector = "Solar", Project_Type = "PV Plant"
```

**Output:**
```
Solar Energy
```

**Supported Sectors:**
- Solar → "Solar Energy"
- Wind → "Wind Energy"
- Battery/Storage → "Energy Storage"
- Hydro → "Hydro Energy"
- Nuclear → "Nuclear Energy"
- Data Center → "Data Centers"
- Technology → "Technology"
- Healthcare → "Healthcare"
- And more...

**Post-Processing:**
- Trims to max 3 words
- Title case formatting
- Fallback: "New Energy"

---

### 2. Client One-liner Template

**Function:** `generate_client_oneliner(prompt_text)`

**Purpose:** Generate anonymous client description sentences

**Trigger Patterns:**
```
- "the client" + ("one sentence" OR "anonymous" OR "briefly states")
```

**Example Input:**
```
Write one anonymous sentence starting with "The client" that briefly states:
- Asset_Description: "a 500 MW solar photovoltaic power plant with battery storage"
- Country: "Vietnam"
- Action: developing
```

**Output:**
```
The client is developing a 500 MW solar photovoltaic power plant with battery storage in Vietnam.
```

**Formula:**
```
The client is [verb_phrase] [asset_description] in [country].
```

**Verb Options:**
- `developing` (default)
- `operating`
- `developing and operating` (if both mentioned)

**Extraction Logic:**
1. Looks for quoted strings in prompt
2. First long quote (20+ chars) → asset description
3. Second quote → country/location
4. Verb determined by keywords in prompt

**Fallback:**
```
The client is developing renewable energy projects.
```

---

### 3. Project Highlight Template (3 Paragraphs)

**Function:** `generate_project_highlight_3para(prompt_text)`

**Purpose:** Generate structured 3-paragraph project highlights

**Trigger Patterns:**
```
- "paragraph 1:" + "paragraph 2:" + "paragraph 3:"
```

**Structure:**

**Paragraph 1: Company Operations + Project Focus**
```
Sentence 1: "The client {Company_Operations}."
Sentence 2: "The proposed project is the {Project_Focus}."
```

**Paragraph 2: Service Scope + Partnerships**
```
Sentence 1: "The client {Service_Scope}."
Sentence 2: "They {Partnership}."
Sentence 3: "{Partnership_Benefit}."
```

**Paragraph 3: Investment Details**
```
Single sentence: "The project will commence with an initial investment of {Initial}, with a projected expansion to {Expansion}."
```

**Example Input:**
```
Generate 3 paragraphs following this structure:

Paragraph 1: Company operations and project focus.
Use: "specializes in renewable energy development and operates 2 GW of solar assets across Asia"
Use: "expansion of their flagship solar portfolio in Vietnam"

Paragraph 2: Service scope and partnerships.
Use: "provides comprehensive EPC and O&M services for utility-scale solar projects"
Use: "partner with leading international equipment suppliers and local contractors"
Use: "This collaboration ensures timely delivery and optimal performance"

Paragraph 3: Investment details.
The project will commence with an initial investment of "$150 million", with a projected expansion to "$500 million over the next 5 years".
```

**Output:**
```
The client specializes in renewable energy development and operates 2 GW of solar assets across Asia. The proposed project is the expansion of their flagship solar portfolio in Vietnam.

The client provides comprehensive EPC and O&M services for utility-scale solar projects. They partner with leading international equipment suppliers and local contractors. This collaboration ensures timely delivery and optimal performance.

The project will commence with an initial investment of "$150 million", with a projected expansion to "$500 million over the next 5 years".
```

**Post-Processing:**
- Trims to exactly 3 paragraphs
- Double newline separation
- Proper punctuation

---

## Post-Processing Function

**Function:** `post_process_length(text, target_words=None, target_sentences=None, target_paragraphs=None)`

**Purpose:** Adjust generated text to match length requirements

**Parameters:**
- `target_words`: Tuple of (min, max) words
- `target_sentences`: Target number of sentences
- `target_paragraphs`: Target number of paragraphs

**Examples:**

```python
# Word trimming
text = "Solar Energy Wind Power Battery Storage"
result = post_process_length(text, target_words=(2, 3))
# Output: "Solar Energy Wind."

# Sentence trimming
text = "First. Second. Third. Fourth."
result = post_process_length(text, target_sentences=2)
# Output: "First. Second."

# Paragraph trimming
text = "Para 1.\n\nPara 2.\n\nPara 3.\n\nPara 4."
result = post_process_length(text, target_paragraphs=3)
# Output: "Para 1.\n\nPara 2.\n\nPara 3."
```

---

## Validation Functions 🔍

### `extract_company_names(row_data)`

**Purpose:** Extract potential company/client names from source data

**Process:**
1. Looks for columns with name-related keywords
2. Extracts significant words (3+ characters)
3. Returns list of potential identifiers

**Columns Checked:**
- company_name, company, client
- issuer, sponsor, project name
- legal name, etc.

**Example:**
```python
data = pd.Series({
    'Company Name': 'ABC Solar Corp',
    'Sponsor': 'XYZ Energy'
})
names = extract_company_names(data)
# Returns: ['abc', 'solar', 'corp', 'xyz', 'energy', ...]
```

---

### `validate_anonymous_text(text, row_data)`

**Purpose:** Ensure generated text maintains anonymity

**Returns:** `(is_valid: bool, reason: str)`

**Logic:**
1. Extracts company names from data
2. Checks if any appear in text
3. Excludes common words (the, client, company, etc.)

**Example:**
```python
text = "The client ABC Corp is developing projects."
is_valid, reason = validate_anonymous_text(text, data)
# Returns: (False, "Company name 'abc corp' appears in text")
```

---

### `validate_client_oneliner(text, row_data)`

**Purpose:** Validate client one-liner meets all requirements

**Returns:** `(is_valid: bool, reason: str)`

**Checks:**
1. Starts with "The client"
2. Single sentence only
3. Maintains anonymity

**Example:**
```python
text = "The client is developing solar projects."
is_valid, reason = validate_client_oneliner(text, data)
# Returns: (True, "")
```

---

### `make_stricter_prompt(original_prompt, validation_failure)`

**Purpose:** Create enhanced prompt after validation failure

**Adds:**
- Explicit format instructions
- Anonymity reminders
- Stricter formatting rules

**Example:**
```python
original = "Write about the client"
stricter = make_stricter_prompt(original, "Missing prefix")
# Returns: Original + "IMPORTANT: Start with 'The client is'..."
```

---

## Generic Fallback Template

**Function:** `generate_generic_content(prompt_text, tone="short")`

**Purpose:** Handle prompts that don't match fixed templates

**When Used:**
- No template pattern detected
- Custom or unusual prompt structures
- Backward compatibility

**Features:**
- Detects word/sentence/paragraph requirements from prompt
- Extracts quoted values, dollar amounts, capitalized phrases
- Builds content using extracted data
- Respects tone setting (short/medium/long)

**Length Detection:**
- `"5 words"` or `"5-10 words"`
- `"3 sentences"` or `"2-3 sentences"`
- `"2 paragraphs"` or `"3-5 paragraphs"`

---

## Main Generation Function

**Function:** `generate_beta_text(prompt_text, row_data, tone="short")`

**Enhanced Flow with Validation:**
```
1. Check for Sector Label pattern
   ↓ Match → generate_sector_label()
           → post_process_length()
           → Return

2. Check for Client One-liner pattern
   ↓ Match → generate_client_oneliner()
           → post_process_length()
           → validate_client_oneliner() ✨
           ↓ Failed?
           → make_stricter_prompt()
           → generate_client_oneliner() (retry)
           → Force "The client" prefix if needed
           → Return

3. Check for Project Highlight pattern
   ↓ Match → generate_project_highlight_3para()
           → post_process_length()
           → validate_anonymous_text() (if required) ✨
           ↓ Failed?
           → make_stricter_prompt()
           → generate_project_highlight_3para() (retry)
           → Return

4. No match → generate_generic_content() (fallback)
```

**Benefits:**
- ✅ Clean separation of concerns
- ✅ Easy to add new templates
- ✅ Testable components
- ✅ **Quality assurance through validation**
- ✅ **Self-correcting with regeneration**
- ✅ Consistent output quality
- ✅ Maintainable codebase

---

## Testing

### Template Functions Test
```bash
python test_templates.py
```

**Test Coverage:**
1. ✅ Sector label generation (2-3 words)
2. ✅ Client one-liner (single sentence)
3. ✅ Project highlight (3 paragraphs)
4. ✅ Post-processing length control

### Validation & Regeneration Test ✨
```bash
python test_validation.py
```

**Test Coverage:**
1. ✅ Company name extraction
2. ✅ Anonymous text validation
3. ✅ Client one-liner validation
4. ✅ Stricter prompt generation
5. ✅ End-to-end generation with validation
6. ✅ Edge cases handling

**Sample Output:**
```
TEST 3: Validate Client One-liner
Valid: "The client is developing projects." ✓
Invalid: "A company is..." ✗ (doesn't start with "The client")
Invalid: "Two sentences. Here." ✗ (not single sentence)
Invalid: "The client ABC Corp..." ✗ (anonymity breach)
```

---

## Adding New Templates

To add a new template:

1. **Create template function:**
```python
def generate_my_template(prompt_text: str) -> str:
    """
    Generate text for my specific pattern.
    """
    # Extract data from prompt
    # Build output using formula
    return result
```

2. **Add detection in `generate_beta_text`:**
```python
is_my_pattern = ('keyword' in prompt_lower and 'other' in prompt_lower)

if is_my_pattern:
    result = generate_my_template(prompt_text)
    return post_process_length(result, target_X=Y)
```

3. **Add tests:**
```python
def test_my_template():
    prompt = "..."
    result = generate_my_template(prompt)
    assert condition, "Test failed"
```

---

## Best Practices

### ✅ DO:
- Use clear trigger patterns
- Extract data from prompts systematically
- Apply post-processing for length control
- **Validate output before returning** ✨
- **Implement regeneration for failures** ✨
- Write tests for new templates
- Document the formula/structure

### ❌ DON'T:
- Mix detection and generation logic
- Hard-code specific values
- Skip post-processing
- **Skip validation checks** ✨
- Create brittle regex patterns
- Forget fallback cases

---

## File Structure

```
app.py
├── # Template Functions
├── generate_sector_label()             # Template 1
├── generate_client_oneliner()          # Template 2
├── generate_project_highlight_3para()  # Template 3
├── 
├── # Support Functions
├── extract_structured_prompt_data()    # Data extraction
├── post_process_length()               # Length control
├── generate_generic_content()          # Fallback
├── 
├── # Validation Functions ✨
├── extract_company_names()             # Name extraction
├── validate_anonymous_text()           # Anonymity check
├── validate_client_oneliner()          # Full validation
├── make_stricter_prompt()              # Prompt enhancement
├── 
└── generate_beta_text()                # Main router

test_templates.py                       # Template tests
test_validation.py                      # Validation tests ✨
```
└── generate_beta_text()           # Main router

test_templates.py                  # Test suite
```

---

## Future Enhancements

Potential new templates to add:
- [ ] Executive summary (5-7 sentences)
- [ ] Company background (2 paragraphs)
- [ ] Risk assessment (bullet points)
- [ ] Timeline description (chronological)
- [ ] Team introduction (multi-person)
- [ ] Financial metrics (structured data)

---

## Questions?

See the inline comments in [app.py](app.py) or run the test suite for examples.
