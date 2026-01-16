"""
Field Mapping Module
Contains mapping dictionaries for different module types.
"""

import logging
from difflib import get_close_matches
from excel_loader import normalize_label

# Setup logging
logger = logging.getLogger(__name__)


def get_module_field_mappings() -> dict:
    """
    Define module-specific field mappings for different Excel formats.
    
    Each module (1, 2, 3) uses different labels for the same logical field.
    This maps each module's specific labels to canonical PPT placeholder tags.
    
    Returns:
        dict: {module_name: {PPT_TAG: excel_label}}
    """
    return {
        "Module1": {
            "ISSUER": "sponsor name",
            "INDUSTRY": "primary focus area",
            "JURISDICTION": "country of incorporation",
            "ISSUANCE_TYPE": "financing type",
            "INITIAL_NOTIONAL": "total financing amount",
            "COUPON_RATE": "coupon rate",
            "COUPON_FREQUENCY": "coupon frequency",
            "TENOR": "requested tenor",
            "CLIENT_SUMMARY": "sponsor summary",
            "PROJECT_HIGHLIGHT": "sponsor background investment strategy"
        },
        "Module2": {
            "ISSUER": "company legal name",
            "INDUSTRY": "primary industry",
            "JURISDICTION": "country of incorporation",
            "ISSUANCE_TYPE": "financing type",
            "INITIAL_NOTIONAL": "total financing amount",
            "COUPON_RATE": "coupon rate",
            "COUPON_FREQUENCY": "coupon frequency",
            "TENOR": "requested tenor",
            "CLIENT_SUMMARY": "company overview business model",
            "PROJECT_HIGHLIGHT": "company growth strategy financial projections"
        },
        "Module3": {
            "ISSUER": "project name",
            "INDUSTRY": "project type",
            "JURISDICTION": "project location country",
            "ISSUANCE_TYPE": "financing type",
            "INITIAL_NOTIONAL": "total project cost",
            "COUPON_RATE": "coupon rate",
            "COUPON_FREQUENCY": "coupon frequency",
            "TENOR": "project tenor",
            "CLIENT_SUMMARY": "project description",
            "PROJECT_HIGHLIGHT": "project overview technical specs impact"
        }
    }


def create_placeholder_mapping() -> dict:
    """
    Define the mapping between PPT placeholder tags and canonical field names.
    
    This allows flexible wording in Excel while maintaining consistent PPT templates.
    COMPREHENSIVE VERSION - includes hundreds of common variations.
    
    Returns:
        dict: {PPT_TAG: [list of possible Excel label variations]}
    """
    base_mapping = {
        # Company/Issuer/Entity Names
        'COMPANY_NAME': [
            'company name', 'company', 'name of company', 'issuer', 'issuer name', 
            'company legal name', 'legal name', 'entity name', 'organization', 
            'organization name', 'firm name', 'business name', 'corporate name',
            'client name', 'client', 'borrower', 'borrower name'
        ],
        'SPONSOR_NAME': [
            'sponsor name', 'sponsor', 'lead sponsor', 'project sponsor', 
            'equity sponsor', 'financial sponsor', 'investor', 'investor name',
            'sponsor entity', 'sponsor legal name'
        ],
        
        # Location Fields
        'LOCATION_COUNTRY': [
            'location country', 'country', 'jurisdiction', 'location', 
            'country of incorporation', 'incorporation country', 'domicile',
            'country of domicile', 'registered country', 'nation', 
            'country of origin', 'home country', 'project location country',
            'project country', 'geographic location', 'geography'
        ],
        'LOCATION_STATE': [
            'location state', 'state', 'province', 'region', 'state province',
            'administrative region', 'locality', 'project location state',
            'project state', 'city state', 'location city', 'city'
        ],
        
        # Industry/Sector/Type
        'ASSET_TYPE': [
            'asset type', 'type of asset', 'asset class', 'sector', 'asset sector',
            'asset category', 'class of asset', 'category'
        ],
        'PROJECT_TYPE': [
            'project type', 'type of project', 'development type', 'project category',
            'type', 'kind of project', 'project classification', 'project class'
        ],
        'INDUSTRY': [
            'industry', 'sector', 'vertical', 'market', 'industry sector',
            'business sector', 'market sector', 'primary industry', 
            'primary focus area', 'focus area', 'industry classification',
            'business line', 'line of business', 'sector classification'
        ],
        
        # Technical/Capacity
        'UNIT': [
            'unit', 'capacity', 'size', 'scale', 'units', 'project size',
            'total capacity', 'installed capacity', 'nameplate capacity',
            'production capacity', 'output', 'volume'
        ],
        'TECHNOLOGY': [
            'technology', 'tech', 'technology type', 'technical approach',
            'technology platform', 'technical solution', 'technical specs',
            'technology specs', 'technical details', 'technical configuration'
        ],
        
        # Financial/Investment
        'INITIAL_INVESTMENT': [
            'initial investment', 'initial capex', 'phase 1 investment', 
            'startup cost', 'initial capital', 'upfront investment',
            'initial capital expenditure', 'initial cost', 'phase i investment',
            'first phase investment', 'initial funding', 'initial amount'
        ],
        'FUTURE_INVESTMENT': [
            'future investment', 'expansion investment', 'phase 2 investment', 
            'growth capex', 'phase ii investment', 'second phase investment',
            'expansion capital', 'future capex', 'additional investment',
            'future capital expenditure', 'expansion cost'
        ],
        'INITIAL_NOTIONAL': [
            'initial notional', 'notional', 'notional amount', 'principal',
            'principal amount', 'total financing amount', 'financing amount',
            'total amount', 'issuance amount', 'issue size', 'total size',
            'total financing', 'total project cost', 'project cost',
            'total investment', 'investment amount', 'transaction size'
        ],
        'COUPON_RATE': [
            'coupon rate', 'coupon', 'rate', 'interest rate', 'yield',
            'annual coupon', 'coupon percentage', 'interest', 'fixed rate'
        ],
        'COUPON_FREQUENCY': [
            'coupon frequency', 'frequency', 'payment frequency', 
            'interest frequency', 'coupon payment frequency', 'payment schedule',
            'interest payment frequency', 'payment period'
        ],
        'TENOR': [
            'tenor', 'term', 'maturity', 'duration', 'requested tenor',
            'project tenor', 'loan tenor', 'financing tenor', 'maturity date',
            'term length', 'time to maturity', 'period'
        ],
        
        # Parties/Partners
        'CONTRACTOR': [
            'contractor', 'epc contractor', 'construction partner', 'builder',
            'construction contractor', 'general contractor', 'epc partner',
            'engineering contractor', 'construction company', 'developer'
        ],
        'OFFTAKER': [
            'offtaker', 'offtake partner', 'purchaser', 'buyer', 'customer',
            'offtake agreement', 'power purchaser', 'product purchaser',
            'offtake counterparty', 'off taker'
        ],
        
        # Status/Timing
        'PROJECT_STATUS': [
            'project status', 'status', 'development stage', 'phase',
            'current status', 'stage', 'project stage', 'development status',
            'project phase', 'current phase', 'stage of development'
        ],
        'COD': [
            'cod', 'commercial operation date', 'operational date', 
            'completion date', 'expected cod', 'commercial operations',
            'start date', 'operations date', 'go live date', 'launch date',
            'commissioning date', 'target cod', 'estimated cod'
        ],
        
        # Descriptions/Summaries
        'DESCRIPTION': [
            'description', 'project description', 'overview', 'summary',
            'project summary', 'project overview', 'details', 'project details',
            'background', 'about', 'information'
        ],
        'CLIENT_SUMMARY': [
            'client summary', 'summary', 'executive summary', 'overview',
            'sponsor summary', 'company overview', 'company summary',
            'company overview business model', 'business model', 'project description',
            'issuer summary', 'issuer overview', 'description', 'about'
        ],
        'PROJECT_HIGHLIGHT': [
            'project highlight', 'highlights', 'key highlights', 'project highlights',
            'sponsor background investment strategy', 'investment strategy',
            'company growth strategy financial projections', 'growth strategy',
            'project overview technical specs impact', 'technical specs',
            'key points', 'important details', 'notable features'
        ],
        'TITLE': [
            'title', 'project title', 'project name', 'deal name', 'name',
            'transaction name', 'facility name', 'deal title', 'project',
            'transaction title', 'name of project', 'name of deal'
        ],
    }
    
    # Add module-specific mappings
    module_mappings = get_module_field_mappings()
    for module_name, field_map in module_mappings.items():
        for tag, excel_label in field_map.items():
            normalized = normalize_label(excel_label)
            if tag in base_mapping:
                # Add to existing list if not already there
                if normalized not in base_mapping[tag]:
                    base_mapping[tag].append(normalized)
            else:
                # Create new entry
                base_mapping[tag] = [normalized]
    
    logger.debug(f"Created placeholder mapping with {len(base_mapping)} tags")
    return base_mapping


def resolve_tag_to_value(tag: str, data_dict: dict, placeholder_mapping: dict, module_type: str = None) -> str:
    """
    Resolve a :TAG: placeholder to its value using the mapping and data dictionary.
    
    Args:
        tag: The tag name (e.g., 'COMPANY_NAME')
        data_dict: Dictionary of normalized_label -> value
        placeholder_mapping: Dictionary of TAG -> [possible label variations]
        module_type: Detected module type ("Module1", "Module2", "Module3")
    
    Returns:
        The resolved value or empty string if not found
    """
    from difflib import get_close_matches
    
    # If we have a detected module, prioritize module-specific labels
    if module_type:
        module_mappings = get_module_field_mappings()
        if module_type in module_mappings and tag in module_mappings[module_type]:
            module_label = module_mappings[module_type][tag]
            normalized = normalize_label(module_label)
            if normalized in data_dict:
                logger.info(f"✓ Resolved tag '{tag}' via {module_type} mapping: '{module_label}' -> value found")
                return data_dict[normalized]
    
    # Check if this tag is in our mapping
    if tag in placeholder_mapping:
        possible_labels = placeholder_mapping[tag]
        
        # Try each possible label variation
        for label_variation in possible_labels:
            normalized = normalize_label(label_variation)
            if normalized in data_dict:
                logger.info(f"✓ Resolved tag '{tag}' to '{label_variation}' -> value found")
                return data_dict[normalized]
        
        # If no exact match, try fuzzy matching with lower threshold (50% instead of 60%)
        available_labels = list(data_dict.keys())
        matches = get_close_matches(normalize_label(possible_labels[0]), available_labels, n=3, cutoff=0.5)
        if matches:
            # Try each match in order of similarity
            for match in matches:
                if match in data_dict:
                    logger.info(f"✓ Resolved tag '{tag}' via fuzzy match: '{match}' (50% threshold) -> value found")
                    return data_dict[match]
    
    # Try using the tag itself as a label (for backwards compatibility)
    normalized_tag = normalize_label(tag)
    if normalized_tag in data_dict:
        logger.info(f"✓ Resolved tag '{tag}' directly (tag name matches label)")
        return data_dict[normalized_tag]
    
    # Last resort: try fuzzy match on the tag itself
    available_labels = list(data_dict.keys())
    matches = get_close_matches(normalized_tag, available_labels, n=1, cutoff=0.5)
    if matches and matches[0] in data_dict:
        logger.info(f"✓ Resolved tag '{tag}' via fuzzy match on tag name: '{matches[0]}'")
        return data_dict[matches[0]]
    
    logger.warning(f"✗ Could not resolve tag '{tag}' - no match found in Excel data")
    return ""
            return data_dict[matches[0]]
    
    # Try using the tag itself as a label (for backwards compatibility)
    normalized_tag = normalize_label(tag)
    if normalized_tag in data_dict:
        logger.debug(f"Resolved tag '{tag}' directly")
        return data_dict[normalized_tag]
    
    # Try fuzzy match on the tag
    available_labels = list(data_dict.keys())
    matches = get_close_matches(normalized_tag, available_labels, n=1, cutoff=0.6)
    if matches and matches[0] in data_dict:
        logger.debug(f"Resolved tag '{tag}' via fuzzy match on tag itself: '{matches[0]}'")
        return data_dict[matches[0]]
    
    logger.warning(f"Could not resolve tag '{tag}' to any value")
    return ""


def fuzzy_match_label(label: str, available_labels: list, cutoff: float = 0.6) -> str:
    """
    Fuzzy match a label against available labels.
    
    Args:
        label: The label to match
        available_labels: List of available labels to match against
        cutoff: Minimum similarity score (0.0 to 1.0)
    
    Returns:
        The best matching label, or None if no match above cutoff
    """
    normalized = normalize_label(label)
    matches = get_close_matches(normalized, available_labels, n=1, cutoff=cutoff)
    
    if matches:
        logger.debug(f"Fuzzy matched '{label}' to '{matches[0]}' (cutoff={cutoff})")
        return matches[0]
    
    logger.debug(f"No fuzzy match found for '{label}' (cutoff={cutoff})")
    return None
