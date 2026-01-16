"""
Module Detector
Detects which module type (Module1, Module2, Module3) based on Excel field names.
"""

import logging
from field_mapping import get_module_field_mappings
from excel_loader import normalize_label

# Setup logging
logger = logging.getLogger(__name__)


def detect_module_type(data_dict: dict) -> str:
    """
    Detect which module type based on Excel field names.
    
    Args:
        data_dict: Dictionary of normalized_label -> value
    
    Returns:
        Module type: "Module1", "Module2", "Module3", or None
    """
    logger.info("Detecting module type from Excel data")
    logger.info(f"Data dict has {len(data_dict)} keys. First 10: {list(data_dict.keys())[:10]}")
    
    module_mappings = get_module_field_mappings()
    
    # Score each module based on how many of its specific fields are present
    scores = {}
    details = {}
    
    for module_name, field_map in module_mappings.items():
        score = 0
        matched_fields = []
        
        # Check unique fields for each module
        for tag, excel_label in field_map.items():
            normalized = normalize_label(excel_label)
            if normalized in data_dict:
                score += 1
                matched_fields.append(excel_label)
        
        scores[module_name] = score
        details[module_name] = matched_fields
        
        logger.info(f"{module_name}: score={score}, matched_fields={matched_fields}")
    
    # Return module with highest score (if any fields matched)
    if scores:
        best_module = max(scores, key=scores.get)
        best_score = scores[best_module]
        
        if best_score > 0:
            logger.info(f"✓ Detected module: {best_module} (score: {best_score}, matched: {details[best_module]})")
            return best_module
        else:
            logger.info("No specific module detected (score: 0 for all modules)")
            return None
    
    logger.warning("Module detection failed: no module mappings available")
    return None
