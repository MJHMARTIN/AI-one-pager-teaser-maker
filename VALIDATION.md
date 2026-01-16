# 🛡️ Validation & Quality Control System

## Overview

The AI text generation system now includes **automatic validation and regeneration** to ensure high-quality, compliant output.

## Features

### ✅ Automated Checks

1. **"The client" Prefix Check**
   - Ensures anonymous one-liners start correctly
   - Auto-corrects if missing

2. **Anonymity Protection**
   - Extracts company names from source data
   - Detects identity leaks in generated text
   - Prevents confidential information exposure

3. **Format Validation**
   - Single sentence requirement for one-liners
   - Paragraph count validation
   - Length requirements enforcement

### 🔄 Regeneration Logic

**When validation fails:**
1. Create stricter prompt with explicit instructions
2. Regenerate text once
3. Apply fixes if still failing (e.g., force "The client" prefix)

**Maximum attempts:** 2 (original + 1 regeneration)

## Implementation

### Validation Functions

```python
# Extract company names from data
extract_company_names(row_data) → list[str]

# Check text maintains anonymity
validate_anonymous_text(text, row_data) → (bool, str)

# Full validation for client one-liners
validate_client_oneliner(text, row_data) → (bool, str)

# Create enhanced prompt after failure
make_stricter_prompt(original, failure_reason) → str
```

### Integration

Validation is integrated into `generate_beta_text()`:

```python
def generate_beta_text(prompt_text, row_data, tone):
    # ... pattern detection ...
    
    if is_client_oneliner:
        # Generate
        result = generate_client_oneliner(prompt_text)
        result = post_process_length(result, target_sentences=1)
        
        # Validate
        is_valid, reason = validate_client_oneliner(result, row_data)
        
        # Regenerate if needed
        if not is_valid:
            stricter = make_stricter_prompt(prompt_text, reason)
            result = generate_client_oneliner(stricter)
            result = post_process_length(result, target_sentences=1)
            
            # Force fix if still failing
            if not result.startswith("The client"):
                result = "The client " + result
        
        return result
```

## Validation Rules

### Client One-liner

| Check | Rule | Action on Failure |
|-------|------|-------------------|
| Prefix | Must start with "The client" | Regenerate → Force prefix |
| Length | Exactly 1 sentence | Regenerate |
| Anonymity | No company names | Regenerate with stricter prompt |

### Project Highlight

| Check | Rule | Action on Failure |
|-------|------|-------------------|
| Anonymity | No company names (if required) | Regenerate with stricter prompt |
| Length | 3 paragraphs | Post-processing trim |

## Example Scenarios

### Scenario 1: Missing Prefix

**First Attempt:**
```
Input: "Write one sentence about the client"
Output: "A company is developing solar projects."
Validation: ✗ Does not start with "The client"
```

**Regeneration:**
```
Stricter prompt: "...IMPORTANT: Start with 'The client is'..."
Output: "The client is developing solar projects."
Validation: ✓ Pass
```

### Scenario 2: Anonymity Breach

**First Attempt:**
```
Input: "Write anonymous sentence about ABC Corp"
Output: "ABC Corp is developing renewable energy."
Validation: ✗ Company name 'abc corp' appears in text
```

**Regeneration:**
```
Stricter prompt: "...Do NOT include company names..."
Output: "The client is developing renewable energy."
Validation: ✓ Pass
```

### Scenario 3: Multiple Issues

**First Attempt:**
```
Output: "ABC Corp does solar. They operate in Asia."
Issues: 
- ✗ Missing "The client" prefix
- ✗ Contains company name
- ✗ Multiple sentences
```

**Regeneration:**
```
Stricter prompt with all requirements
Output: "The client is developing and operating solar facilities in Asia."
Fixes:
- ✓ Starts with "The client"
- ✓ Anonymous
- ✓ Single sentence
```

## Testing

### Run Validation Tests
```bash
python test_validation.py
```

### Test Coverage

1. **Extract Company Names**
   - Column detection
   - Name extraction
   - Word filtering

2. **Anonymous Text Validation**
   - Identity leak detection
   - Common word exclusion
   - Case-insensitive matching

3. **Client One-liner Validation**
   - Prefix check
   - Sentence count
   - Anonymity check

4. **Stricter Prompt Generation**
   - Instruction additions
   - Format emphasis
   - Anonymity reminders

5. **End-to-End Generation**
   - Full flow with validation
   - Regeneration logic
   - Quality assurance

## Configuration

### Adjustable Parameters

```python
# In extract_company_names()
MIN_WORD_LENGTH = 3  # Minimum chars to consider as name part

# In validate_anonymous_text()
COMMON_WORDS = {'the', 'client', 'company', 'project', ...}

# In generate_beta_text()
MAX_REGENERATION_ATTEMPTS = 1  # Currently: original + 1 retry
```

## Benefits

### Quality Assurance
- ✅ Consistent format compliance
- ✅ Protected confidentiality
- ✅ Professional output

### Automation
- ✅ Self-correcting system
- ✅ No manual review needed
- ✅ Reliable results

### Maintainability
- ✅ Testable validation logic
- ✅ Clear failure reasons
- ✅ Easy to extend

## Future Enhancements

Potential additions:
- [ ] Configurable regeneration attempts
- [ ] Validation for sector labels
- [ ] Grammar and spell checking
- [ ] Tone consistency validation
- [ ] Custom validation rules per template
- [ ] Validation result logging/analytics

## Questions?

See:
- [AI_TEMPLATES.md](AI_TEMPLATES.md) - Template documentation
- [test_validation.py](test_validation.py) - Test examples
- [app.py](app.py) - Implementation code
