"""
AI Generators Module
Contains all AI text generation functions with validation and quality control.
Now powered by Perplexity API for enhanced AI capabilities.
"""

import pandas as pd
import re
import logging
import requests

# Setup logging
logger = logging.getLogger(__name__)

# Perplexity API Configuration
PERPLEXITY_API_KEY = os.environ.get("PERPLEXITY_API_KEY", "")
PERPLEXITY_API_URL = "https://api.perplexity.ai/chat/completions"


# ============================================================================
# PERPLEXITY API INTEGRATION
# ============================================================================

def call_perplexity_api(prompt: str, max_tokens: int = 200) -> str:
    """
    Call Perplexity API to generate text based on a prompt.
    
    Args:
        prompt: The instruction/prompt for text generation
        max_tokens: Maximum tokens in response (default 200)
    
    Returns:
        Generated text from Perplexity API
    """
    try:
        headers = {
            "Authorization": f"Bearer {PERPLEXITY_API_KEY}",
            "Content-Type": "application/json"
        }
        
        payload = {
            "model": "llama-3.1-sonar-small-128k-online",  # Fast and efficient model
            "messages": [
                {
                    "role": "system",
                    "content": "You are a professional business content writer. Generate precise, concise content that exactly matches the requested format and length. Follow instructions carefully."
                },
                {
                    "role": "user",
                    "content": prompt
                }
            ],
            "max_tokens": max_tokens,
            "temperature": 0.3,  # Lower temperature for more consistent output
            "top_p": 0.9,
            "return_citations": False,
            "stream": False
        }
        
        logger.debug(f"Calling Perplexity API with prompt length: {len(prompt)} chars")
        response = requests.post(PERPLEXITY_API_URL, json=payload, headers=headers, timeout=30)
        
        if response.status_code == 200:
            result = response.json()
            generated_text = result['choices'][0]['message']['content'].strip()
            logger.info(f"✓ Perplexity API call successful, generated {len(generated_text)} chars")
            return generated_text
        else:
            logger.error(f"Perplexity API error {response.status_code}: {response.text}")
            return None
            
    except Exception as e:
        logger.error(f"Exception calling Perplexity API: {str(e)}")
        return None


def fallback_to_local_generation(prompt_text: str, generation_type: str) -> str:
    """
    Fallback to local pattern-based generation if API fails.
    """
    logger.warning(f"Using fallback local generation for: {generation_type}")
    
    if generation_type == "sector_label":
        return generate_sector_label_local(prompt_text, max_words=3)
    elif generation_type == "client_oneliner":
        return generate_client_oneliner_local(prompt_text)
    elif generation_type == "project_highlight":
        return generate_project_highlight_3para_local(prompt_text)
    else:
        return generate_generic_content_local(prompt_text)


# ============================================================================
# TEMPLATE GENERATION FUNCTIONS (API-POWERED)
# ============================================================================

def generate_sector_label(prompt_text: str, max_words: int = 3) -> str:
    """Generate a 2-3 word sector label using Perplexity API."""
    logger.debug(f"Generating sector label (max_words={max_words}) with Perplexity API")
    
    # Create focused API prompt
    api_prompt = f"""Generate a concise 2-3 word sector label based on this description:

{prompt_text}

Requirements:
- EXACTLY 2-3 words maximum
- Title case (e.g., "Solar Energy", "Data Centers")
- Industry sector focus
- No explanations, just the label

Sector label:"""
    
    result = call_perplexity_api(api_prompt, max_tokens=20)
    
    if result:
        # Clean up the result
        result = result.strip().strip('"').strip("'")
        # Ensure it's title case
        result = ' '.join(word.capitalize() for word in result.split())
        logger.info(f"Generated sector label via API: '{result}'")
        return result
    else:
        # Fallback to local generation
        return fallback_to_local_generation(prompt_text, "sector_label")


def generate_client_oneliner(prompt_text: str) -> str:
    """Generate anonymous client one-liner description using Perplexity API."""
    logger.debug("Generating client one-liner with Perplexity API")
    
    # Create focused API prompt
    api_prompt = f"""Generate an anonymous, professional one-sentence description based on these details:

{prompt_text}

CRITICAL Requirements:
- Start with "The client is" or "The client"
- ONE sentence only (no line breaks)
- End with a period
- Maintain complete anonymity (no company names, no specific project names)
- Use generic terms: "the client", "the project", "the facility"
- Be professional and concise

One-liner:"""
    
    result = call_perplexity_api(api_prompt, max_tokens=100)
    
    if result:
        # Clean up the result
        result = result.strip()
        # Ensure it starts with "The client"
        if not result.startswith("The client"):
            result = "The client " + result.lstrip("Tt").lstrip("he ").lstrip("client ")
        # Ensure single sentence
        sentences = [s.strip() for s in result.split('.') if s.strip()]
        if sentences:
            result = sentences[0] + '.'
        logger.info(f"Generated client one-liner via API: {result[:80]}...")
        return result
    else:
        # Fallback to local generation
        return fallback_to_local_generation(prompt_text, "client_oneliner")


def generate_project_highlight_3para(prompt_text: str) -> str:
    """Generate 3-paragraph project highlight using Perplexity API."""
    logger.debug("Generating 3-paragraph project highlight with Perplexity API")
    
    # Create focused API prompt
    api_prompt = f"""Generate a professional 3-paragraph project description based on these specifications:

{prompt_text}

CRITICAL Requirements:
- EXACTLY 3 paragraphs
- Separate paragraphs with double line breaks
- Maintain anonymity (use "The client", "the project", not specific names)
- Professional business writing style
- Each paragraph 2-4 sentences
- Follow the structure outlined in the specifications

Project highlight:"""
    
    result = call_perplexity_api(api_prompt, max_tokens=400)
    
    if result:
        # Clean up the result
        result = result.strip()
        # Ensure proper paragraph separation
        paragraphs = [p.strip() for p in result.split('\n\n') if p.strip()]
        if not paragraphs:
            # Try single newline split
            paragraphs = [p.strip() for p in result.split('\n') if p.strip()]
        # Take first 3 paragraphs
        result = '\n\n'.join(paragraphs[:3])
        logger.info(f"Generated project highlight via API with {len(paragraphs)} paragraphs")
        return result
    else:
        # Fallback to local generation
        return fallback_to_local_generation(prompt_text, "project_highlight")


def generate_generic_content(prompt_text: str, tone: str = "short") -> str:
    """Generate generic content using Perplexity API."""
    logger.debug(f"Generating generic content with Perplexity API (tone={tone})")
    
    # Determine max_tokens based on tone
    token_map = {"short": 100, "medium": 200, "long": 300}
    max_tokens = token_map.get(tone, 200)
    
    # Create API prompt
    api_prompt = f"""Generate professional business content based on this prompt:

{prompt_text}

Generate {tone} content following any specifications in the prompt above.

Content:"""
    
    result = call_perplexity_api(api_prompt, max_tokens=max_tokens)
    
    if result:
        logger.info(f"Generated generic content via API: {result[:80]}...")
        return result.strip()
    else:
        # Fallback to local generation
        return fallback_to_local_generation(prompt_text, "generic")


# ============================================================================
# LOCAL GENERATION FUNCTIONS (FALLBACK)
# ============================================================================

def generate_sector_label_local(prompt_text: str, max_words: int = 3) -> str:
    """Generate a 2-3 word sector label using local pattern matching."""
    logger.debug(f"Generating sector label locally (max_words={max_words})")
    
    prompt_lower = prompt_text.lower()
    
    # Keyword-based sector detection
    sector_patterns = {
        'solar': 'Solar Energy',
        'wind': 'Wind Energy',
        'battery': 'Energy Storage',
        'storage': 'Energy Storage',
        'hydro': 'Hydro Energy',
        'nuclear': 'Nuclear Energy',
        'geothermal': 'Geothermal Energy',
        'biomass': 'Biomass Energy',
        'hydrogen': 'Hydrogen Energy',
        'data center': 'Data Centers',
        'infrastructure': 'Infrastructure',
        'real estate': 'Real Estate',
        'technology': 'Technology',
        'healthcare': 'Healthcare',
        'manufacturing': 'Manufacturing',
    }
    
    # Check for keywords in prompt
    for keyword, label in sector_patterns.items():
        if keyword in prompt_lower:
            logger.info(f"Generated sector label: '{label}' (matched keyword: '{keyword}')")
            return label
    
    # Check for "new energy" or "renewable"
    if 'new energy' in prompt_lower or 'renewable' in prompt_lower:
        logger.info("Generated sector label: 'New Energy'")
        return 'New Energy'
    
    # Extract quoted values as fallback
    quoted_values = re.findall(r'["\']([^"\']+)["\']', prompt_text)
    if quoted_values:
        words = quoted_values[0].split()[:max_words]
        result = ' '.join(words).title()
        logger.info(f"Generated sector label from quoted text: '{result}'")
        return result
    
    logger.info("Generated default sector label: 'New Energy'")
    return 'New Energy'


def generate_client_oneliner_local(prompt_text: str) -> str:
    """Generate anonymous client one-liner description using local patterns."""
    logger.debug("Generating client one-liner locally")
    
    prompt_lower = prompt_text.lower()
    
    # Extract quoted values (asset description and country)
    quoted_values = re.findall(r'["\']([^"\']{20,})["\']', prompt_text)
    if not quoted_values:
        quoted_values = re.findall(r'["\']([^"\']{5,})["\']', prompt_text)
    
    if len(quoted_values) >= 2:
        asset_description = quoted_values[0]
        country = quoted_values[1]
        
        # Determine action verb
        if 'operat' in prompt_lower and 'develop' in prompt_lower:
            verb_phrase = 'developing and operating'
        elif 'operat' in prompt_lower:
            verb_phrase = 'operating'
        else:
            verb_phrase = 'developing'
        
        result = f"The client is {verb_phrase} {asset_description} in {country}."
        logger.info(f"Generated client one-liner: {result[:80]}...")
        return result
    
    elif len(quoted_values) == 1:
        asset_description = quoted_values[0]
        verb_phrase = 'operating' if 'operat' in prompt_lower else 'developing'
        result = f"The client is {verb_phrase} {asset_description}."
        logger.info(f"Generated client one-liner: {result[:80]}...")
        return result
    
    # Fallback
    logger.warning("Using fallback client one-liner")
    return "The client is developing renewable energy projects."


def generate_project_highlight_3para_local(prompt_text: str) -> str:
    """Generate 3-paragraph project highlight using local patterns."""
    logger.debug("Generating 3-paragraph project highlight locally")
    
    data = extract_structured_prompt_data(prompt_text)
    paragraphs = []
    
    # Paragraph 1: Company Operations + Project Focus
    if 'paragraph_1' in data and len(data['paragraph_1']) >= 1:
        company_ops = data['paragraph_1'][0]
        project_focus = data['paragraph_1'][1] if len(data['paragraph_1']) > 1 else ""
        
        if not project_focus and 'called' in prompt_text:
            called_match = re.search(r'called\s+["\']([^"\']+)["\']', prompt_text, re.IGNORECASE)
            if called_match:
                project_focus = called_match.group(1)
        
        sent1 = f"The client {company_ops}."
        sent2 = f"The proposed project is the {project_focus}." if project_focus else ""
        paragraphs.append(f"{sent1} {sent2}".strip())
    
    # Paragraph 2: Service Scope + Partnerships
    if 'paragraph_2' in data and len(data['paragraph_2']) >= 1:
        service_scope = data['paragraph_2'][0]
        partnership = data['paragraph_2'][1] if len(data['paragraph_2']) > 1 else ""
        partnership_benefit = data['paragraph_2'][2] if len(data['paragraph_2']) > 2 else ""
        
        sent1 = f"The client {service_scope}."
        
        if partnership:
            partnership_clean = partnership.strip()
            if partnership_clean.startswith('partner with') or partnership_clean.startswith('partner in'):
                sent2 = f"They {partnership_clean}."
            elif 'partnered' in partnership_clean:
                sent2 = f"They are {partnership_clean}."
            else:
                sent2 = f"They {partnership_clean}."
        else:
            sent2 = ""
        
        sent3 = f"{partnership_benefit}." if partnership_benefit else ""
        para2 = f"{sent1} {sent2} {sent3}".strip()
        paragraphs.append(para2)
    
    # Paragraph 3: Investment Details
    if 'paragraph_3' in data or 'all_values' in data:
        investment_pattern = re.search(
            r'The project will commence with an initial investment of\s+([^,]+),\s*with a projected expansion to\s+([^.]+)',
            prompt_text,
            re.IGNORECASE
        )
        
        if investment_pattern:
            para3 = f"The project will commence with an initial investment of {investment_pattern.group(1)}, with a projected expansion to {investment_pattern.group(2)}."
            paragraphs.append(para3)
    
    result = "\n\n".join(p for p in paragraphs if p)
    result = result if result else "Project details are being finalized."
    
    logger.info(f"Generated project highlight with {len(paragraphs)} paragraphs")
    return result


def extract_structured_prompt_data(prompt_text: str) -> dict:
    """Extract structured data from paragraph-based prompts."""
    data = {}
    
    paragraphs = re.split(r'Paragraph \d+:', prompt_text)
    
    for idx, para in enumerate(paragraphs[1:], 1):
        quoted = re.findall(r'["\']([^"\']{10,})["\']', para)
        if quoted:
            data[f'paragraph_{idx}'] = quoted
    
    all_quoted = re.findall(r'["\']([^"\']{10,})["\']', prompt_text)
    data['all_values'] = all_quoted
    
    return data


def generate_generic_content_local(prompt_text: str, tone: str = "short") -> str:
    """Fallback generic content generator using local patterns."""
    logger.debug(f"Generating generic content locally (tone={tone})")
    
    prompt_lower = prompt_text.lower()
    
    # Detect length requirements
    word_count = None
    sentence_count = None
    paragraph_count = None
    
    word_match = re.search(r'(\d+)(?:-(\d+))?\s+words?', prompt_lower)
    if word_match:
        word_count = (int(word_match.group(1)), int(word_match.group(2) or word_match.group(1)))
    
    sent_match = re.search(r'(\d+)(?:-(\d+))?\s+sentences?', prompt_lower)
    if sent_match:
        sentence_count = (int(sent_match.group(1)), int(sent_match.group(2) or sent_match.group(1)))
    
    para_match = re.search(r'(\d+)(?:-(\d+))?\s+paragraphs?', prompt_lower)
    if para_match:
        paragraph_count = (int(para_match.group(1)), int(para_match.group(2) or para_match.group(1)))
    
    # Extract values
    values = []
    values.extend(re.findall(r'"([^"]+)"', prompt_text))
    values.extend(re.findall(r'\$[\d,]+(?:\s*(?:million|billion|thousand))?', prompt_text, re.IGNORECASE))
    
    capitalized = re.findall(r'\b[A-Z][a-zA-Z]+(?:\s+[A-Z][a-zA-Z]+)*\b', prompt_text)
    excluded = {'Using', 'Use', 'Write', 'Follow', 'Sentence', 'Do', 'Excel', 'English', 'Initial', 'Future', 'The', 'Based'}
    values.extend([c for c in capitalized if c not in excluded and len(c) > 2])
    
    technical_terms = re.findall(r'\b[a-z]+(?:\s+[a-z]+){1,5}\b', prompt_text)
    values.extend([t for t in technical_terms if len(t) > 15 and 'write' not in t][:3])
    
    unique_values = list(dict.fromkeys([v.strip() for v in values if v.strip() and len(v.strip()) > 1]))
    
    # Generate
    if word_count and word_count[1] <= 10:
        result = unique_values[0] if unique_values else "Professional solutions"
    elif paragraph_count or (sentence_count and sentence_count[1] >= 3):
        sentences = []
        target = sentence_count[1] if sentence_count else (paragraph_count[1] * 3 if paragraph_count else 3)
        
        if len(unique_values) >= 2:
            sentences.append(f"{unique_values[0]} delivers innovative solutions in {unique_values[1]}.")
        elif len(unique_values) >= 1:
            sentences.append(f"{unique_values[0]} represents excellence in their field.")
        else:
            sentences.append("Professional solutions delivered with expertise.")
        
        if target >= 2:
            if len(unique_values) >= 3:
                sentences.append(f"Leveraging {unique_values[2]}, the organization drives exceptional outcomes.")
            else:
                sentences.append("Advanced capabilities ensure outstanding results.")
        
        if target >= 3:
            sentences.append("Proven track record of success across diverse initiatives.")
        
        if target >= 4:
            sentences.append("Strategic partnerships and innovation create lasting value.")
        
        result = " ".join(sentences[:target])
    else:
        if len(unique_values) >= 2:
            result = f"{unique_values[0]} delivers excellence in {unique_values[1]}."
        elif len(unique_values) >= 1:
            result = f"{unique_values[0]} drives innovation and results."
        else:
            result = "Excellence delivered through proven expertise."
    
    logger.info(f"Generated generic content: {result[:80]}...")
    return result


def post_process_length(
    text: str,
    target_words: tuple = None,
    target_sentences: int = None,
    target_paragraphs: int = None
) -> str:
    """Post-process generated text to match target length requirements."""
    
    # Word count trimming
    if target_words:
        words = text.split()
        if len(words) > target_words[1]:
            text = ' '.join(words[:target_words[1]])
            if not text.endswith('.'):
                text += '.'
            logger.debug(f"Trimmed to {target_words[1]} words")
    
    # Sentence count trimming
    if target_sentences:
        sentences = re.split(r'[.!?]+\s+', text)
        sentences = [s.strip() for s in sentences if s.strip()]
        if len(sentences) > target_sentences:
            text = '. '.join(sentences[:target_sentences])
            if not text.endswith('.'):
                text += '.'
            logger.debug(f"Trimmed to {target_sentences} sentences")
    
    # Paragraph count trimming
    if target_paragraphs:
        paragraphs = text.split('\n\n')
        paragraphs = [p.strip() for p in paragraphs if p.strip()]
        if len(paragraphs) > target_paragraphs:
            text = '\n\n'.join(paragraphs[:target_paragraphs])
            logger.debug(f"Trimmed to {target_paragraphs} paragraphs")
    
    return text.strip()


# ============================================================================
# VALIDATION FUNCTIONS
# ============================================================================

def extract_company_names(row_data: pd.Series) -> list:
    """Extract potential company/client names from row data."""
    potential_names = []
    
    name_columns = [
        'company_name', 'company', 'client', 'client_name', 
        'issuer', 'sponsor', 'sponsor_name', 'name',
        'company legal name', 'legal name', 'project name'
    ]
    
    for col in row_data.index:
        col_lower = str(col).lower()
        if any(name_key in col_lower for name_key in name_columns):
            value = str(row_data[col])
            if value and value.lower() not in ['nan', 'none', '']:
                words = value.split()
                for word in words:
                    if len(word) >= 3:
                        potential_names.append(word.lower())
                potential_names.append(value.lower())
    
    names = list(set(potential_names))
    if names:
        logger.debug(f"Extracted {len(names)} potential company names")
    return names


def validate_anonymous_text(text: str, row_data: pd.Series) -> tuple:
    """Validate that text maintains anonymity."""
    text_lower = text.lower()
    company_names = extract_company_names(row_data)
    
    for name in company_names:
        if len(name) > 3 and name in text_lower:
            common_words = {'the', 'client', 'company', 'project', 'group', 'inc', 'ltd', 'corp', 'llc'}
            if name not in common_words:
                logger.warning(f"Anonymity breach detected: '{name}' found in generated text")
                return (False, f"Company name '{name}' appears in text (anonymity breach)")
    
    logger.debug("Anonymity validation passed")
    return (True, "")


def validate_client_oneliner(text: str, row_data: pd.Series) -> tuple:
    """Validate client one-liner meets all requirements."""
    text_stripped = text.strip()
    
    # Check 1: Starts with "The client"
    if not text_stripped.startswith("The client"):
        logger.warning("Validation failed: does not start with 'The client'")
        return (False, "Does not start with 'The client'")
    
    # Check 2: Single sentence
    sentences = [s.strip() for s in text_stripped.split('.') if s.strip()]
    if len(sentences) != 1:
        logger.warning(f"Validation failed: contains {len(sentences)} sentences, expected 1")
        return (False, f"Contains {len(sentences)} sentences, expected 1")
    
    # Check 3: Anonymity
    is_anonymous, reason = validate_anonymous_text(text, row_data)
    if not is_anonymous:
        return (False, reason)
    
    logger.debug("Client one-liner validation passed")
    return (True, "")


def make_stricter_prompt(original_prompt: str, validation_failure: str) -> str:
    """Create stricter prompt after validation failure."""
    logger.info(f"Creating stricter prompt (reason: {validation_failure})")
    
    prompt_lower = original_prompt.lower()
    
    if 'the client' in prompt_lower and 'anonymous' in prompt_lower:
        strict_additions = [
            "IMPORTANT: Start the sentence with exactly 'The client is'",
            "Do NOT include any company names, project names, or client identifiers",
            "Use only generic terms like 'The client', 'the project', 'the facility'",
            "Keep it to ONE sentence only"
        ]
        return original_prompt + "\n\n" + "\n".join(strict_additions)
    
    return original_prompt + "\n\nIMPORTANT: Follow the exact format specified. Maintain complete anonymity."


# ============================================================================
# MAIN GENERATION FUNCTION WITH VALIDATION
# ============================================================================

def generate_beta_text(prompt_text: str, row_data: pd.Series, tone: str = "short") -> str:
    """
    Main text generation function with validation and regeneration.
    Now powered by Perplexity API with local fallback.
    
    Args:
        prompt_text: The fully substituted prompt
        row_data: Source data (for validation)
        tone: "short", "medium", or "long"
    
    Returns:
        Generated and validated text
    """
    logger.info("=== Starting text generation (Perplexity API) ===")
    logger.debug(f"Prompt length: {len(prompt_text)} chars, tone: {tone}")
    
    prompt_lower = prompt_text.lower()
    
    # TEMPLATE 1: Sector Label
    is_sector_label = (
        ('sector label' in prompt_lower or 'short label' in prompt_lower) and
        ('1-3 word' in prompt_lower or '2-3 word' in prompt_lower or 'short' in prompt_lower)
    )
    
    if is_sector_label:
        logger.info("Pattern matched: Sector Label")
        result = generate_sector_label(prompt_text, max_words=3)
        result = post_process_length(result, target_words=(2, 3))
        logger.info(f"✓ Generated sector label: '{result}'")
        return result
    
    # TEMPLATE 2: Client One-liner (with validation)
    is_client_oneliner = (
        'the client' in prompt_lower and
        ('one sentence' in prompt_lower or 'anonymous' in prompt_lower or 'briefly states' in prompt_lower)
    )
    
    if is_client_oneliner:
        logger.info("Pattern matched: Client One-liner")
        
        # First attempt
        result = generate_client_oneliner(prompt_text)
        result = post_process_length(result, target_sentences=1)
        
        # Validate
        is_valid, failure_reason = validate_client_oneliner(result, row_data)
        
        if not is_valid:
            logger.warning(f"Validation failed: {failure_reason}. Regenerating...")
            stricter_prompt = make_stricter_prompt(prompt_text, failure_reason)
            result = generate_client_oneliner(stricter_prompt)
            result = post_process_length(result, target_sentences=1)
            
            # Force fix if still failing
            if not result.strip().startswith("The client"):
                logger.info("Force-fixing: adding 'The client' prefix")
                result = "The client " + result.strip()
                if not result.endswith('.'):
                    result += '.'
        
        logger.info(f"✓ Generated client one-liner: {result[:80]}...")
        return result
    
    # TEMPLATE 3: Project Highlight (with anonymity check)
    is_project_highlight = (
        'paragraph 1:' in prompt_lower and 
        'paragraph 2:' in prompt_lower and 
        'paragraph 3:' in prompt_lower
    )
    
    if is_project_highlight:
        logger.info("Pattern matched: Project Highlight (3 paragraphs)")
        requires_anonymity = 'anonymous' in prompt_lower or 'the client' in prompt_lower
        
        # First attempt
        result = generate_project_highlight_3para(prompt_text)
        result = post_process_length(result, target_paragraphs=3)
        
        # Validate anonymity if required
        if requires_anonymity:
            is_anonymous, failure_reason = validate_anonymous_text(result, row_data)
            
            if not is_anonymous:
                logger.warning(f"Anonymity check failed: {failure_reason}. Regenerating...")
                stricter_prompt = make_stricter_prompt(prompt_text, failure_reason)
                result = generate_project_highlight_3para(stricter_prompt)
                result = post_process_length(result, target_paragraphs=3)
        
        logger.info("✓ Generated project highlight")
        return result
    
    # FALLBACK: Generic content
    logger.info("No specific pattern matched, using generic generation")
    result = generate_generic_content(prompt_text, tone)
    logger.info(f"✓ Generated generic content: {result[:80]}...")
    return result
