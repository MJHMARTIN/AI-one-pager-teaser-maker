"""
Excel Loader Module - Search-Based Version
Handles Excel files where labels can appear anywhere in any row,
and values are always in the column immediately to the right.
"""

import pandas as pd
import re
import logging

# Setup logging
logger = logging.getLogger(__name__)


def normalize_label(label: str) -> str:
    """
    Normalize Excel label for consistent matching.
    
    Args:
        label: Raw label from Excel
    
    Returns:
        Normalized label (lowercase, no punctuation, single spaces)
    """
    if not isinstance(label, str):
        return ""
    
    # Lowercase, strip spaces, remove punctuation
    normalized = label.lower().strip()
    # Remove common punctuation but keep underscores
    normalized = re.sub(r'[^\w\s]', '', normalized)
    # Replace multiple spaces with single space
    normalized = re.sub(r'\s+', ' ', normalized)
    
    return normalized


def search_label_in_sheet(df: pd.DataFrame, target_label: str, sheet_name: str = None) -> str:
    """
    Search for a label anywhere in the sheet and return the value from the next column.
    
    Args:
        df: DataFrame to search
        target_label: The label to search for (will be normalized)
        sheet_name: Name of sheet (for logging)
    
    Returns:
        The value from the column to the right of where the label was found, or empty string
    """
    target_normalized = normalize_label(target_label)
    
    logger.debug(f"Searching for '{target_label}' (normalized: '{target_normalized}') in sheet '{sheet_name}'")
    
    # Search through every row and every column
    for row_idx in range(len(df)):
        for col_idx in range(len(df.columns) - 1):  # -1 because we need a column to the right
            cell_value = df.iloc[row_idx, col_idx]
            
            # Skip empty cells
            if pd.isna(cell_value):
                continue
            
            cell_normalized = normalize_label(str(cell_value))
            
            # Check for exact match
            if cell_normalized == target_normalized:
                # Get value from next column (same row)
                value_col_idx = col_idx + 1
                value = df.iloc[row_idx, value_col_idx]
                
                if pd.isna(value):
                    logger.debug(f"  Found label at row {row_idx}, col {col_idx}, but value is empty")
                    return ""
                
                value_str = str(value).strip()
                logger.info(f"  ✓ Found '{target_label}' at row {row_idx}, col {col_idx} → value: '{value_str[:50]}'")
                return value_str
            
            # Also check if the target is contained within the cell (partial match)
            elif target_normalized in cell_normalized and len(target_normalized) > 5:
                value_col_idx = col_idx + 1
                value = df.iloc[row_idx, value_col_idx]
                
                if not pd.isna(value):
                    value_str = str(value).strip()
                    logger.info(f"  ✓ Found partial match '{target_label}' at row {row_idx}, col {col_idx} → value: '{value_str[:50]}'")
                    return value_str
    
    logger.debug(f"  ✗ '{target_label}' not found in sheet '{sheet_name}'")
    return ""


def extract_all_data_from_sheet(df: pd.DataFrame, sheet_name: str = None) -> dict:
    """
    Extract ALL label-value pairs from a sheet by searching every row.
    Assumes value is always in the column immediately to the right of the label.
    
    Args:
        df: DataFrame to parse
        sheet_name: Name of sheet (for logging)
    
    Returns:
        dict: {normalized_label: value}
    """
    data_dict = {}
    
    logger.debug(f"Extracting all data from sheet '{sheet_name}' (shape: {df.shape})")
    
    # Search through every row and every column
    for row_idx in range(len(df)):
        for col_idx in range(len(df.columns) - 1):  # -1 because we need a column to the right
            cell_value = df.iloc[row_idx, col_idx]
            
            # Skip empty cells
            if pd.isna(cell_value):
                continue
            
            cell_str = str(cell_value).strip()
            
            # Skip if it's just a number or very short (likely not a label)
            if len(cell_str) < 3:
                continue
            
            # Skip if it's all numbers (likely a value, not a label)
            if cell_str.replace('.', '').replace(',', '').replace('$', '').replace('%', '').replace(' ', '').isdigit():
                continue
            
            # This looks like it could be a label - get the value from next column
            value_col_idx = col_idx + 1
            value = df.iloc[row_idx, value_col_idx]
            
            if pd.isna(value):
                continue
            
            value_str = str(value).strip()
            
            # Skip if value is empty
            if not value_str:
                continue
            
            # Normalize the label and store
            normalized_label = normalize_label(cell_str)
            
            # Only store if normalized label is meaningful
            if len(normalized_label) > 2:
                data_dict[normalized_label] = value_str
                logger.debug(f"  Row {row_idx}, Col {col_idx}: '{normalized_label}' = '{value_str[:50]}'")
    
    logger.info(f"Extracted {len(data_dict)} label-value pairs from sheet '{sheet_name}'")
    return data_dict


def parse_multi_sheet_excel_search_based(excel_file) -> tuple:
    """
    Parse multi-sheet Excel using search-based approach.
    Searches for labels anywhere in any sheet and extracts values from the column to the right.
    
    Args:
        excel_file: File-like object or path to Excel (.xlsx or .xlsm)
    
    Returns:
        tuple: (unified_data_dict, sheet_info_list)
    """
    logger.info("Reading multi-sheet Excel file with search-based approach")
    
    # Read all sheets
    sheets_dict = pd.read_excel(excel_file, sheet_name=None, engine='openpyxl', header=None)
    
    logger.info(f"Found {len(sheets_dict)} sheets")
    
    unified_data = {}
    sheet_info = []
    
    for sheet_name, df in sheets_dict.items():
        if df.empty:
            logger.warning(f"Sheet '{sheet_name}' is empty, skipping")
            sheet_info.append({
                'name': sheet_name,
                'format': 'empty',
                'fields': 0
            })
            continue
        
        logger.info(f"Processing sheet '{sheet_name}' (shape: {df.shape})")
        
        try:
            # Extract all label-value pairs from this sheet
            sheet_data = extract_all_data_from_sheet(df, sheet_name)
            
            if sheet_data:
                # Merge into unified dict (later sheets override earlier ones if there are duplicates)
                unified_data.update(sheet_data)
                
                sheet_info.append({
                    'name': sheet_name,
                    'format': 'search-based',
                    'fields': len(sheet_data)
                })
                
                logger.info(f"✓ Sheet '{sheet_name}': {len(sheet_data)} fields extracted")
            else:
                sheet_info.append({
                    'name': sheet_name,
                    'format': 'no data found',
                    'fields': 0
                })
                logger.warning(f"Sheet '{sheet_name}': no data extracted")
                
        except Exception as e:
            logger.error(f"Error processing sheet '{sheet_name}': {str(e)}")
            sheet_info.append({
                'name': sheet_name,
                'format': f'error: {str(e)[:30]}',
                'fields': 0
            })
    
    logger.info(f"Search-based parsing complete: {len(unified_data)} total fields from {len(sheet_info)} sheets")
    return unified_data, sheet_info


def parse_excel_with_module_mappings(excel_file, module_mappings: dict) -> dict:
    """
    Parse Excel by searching for specific module field labels across all sheets.
    
    Args:
        excel_file: File-like object or path to Excel
        module_mappings: Dict of {module_name: {TAG: excel_label}}
    
    Returns:
        dict: {normalized_label: value} for all found fields
    """
    logger.info("Reading Excel with module-specific label search")
    
    # Read all sheets without headers
    sheets_dict = pd.read_excel(excel_file, sheet_name=None, engine='openpyxl', header=None)
    
    logger.info(f"Found {len(sheets_dict)} sheets")
    
    # First, extract everything we can find
    unified_data, _ = parse_multi_sheet_excel_search_based(excel_file)
    
    # Then, also do targeted searches for module-specific labels
    for module_name, field_map in module_mappings.items():
        logger.info(f"Searching for {module_name} specific labels...")
        
        for tag, excel_label in field_map.items():
            normalized = normalize_label(excel_label)
            
            # Skip if we already found this
            if normalized in unified_data:
                logger.debug(f"  Already have '{excel_label}'")
                continue
            
            # Search for it across all sheets
            for sheet_name, df in sheets_dict.items():
                value = search_label_in_sheet(df, excel_label, sheet_name)
                if value:
                    unified_data[normalized] = value
                    logger.info(f"  ✓ Found {tag} ('{excel_label}') in sheet '{sheet_name}'")
                    break  # Found it, move to next label
    
    return unified_data


# Keep old functions for backward compatibility
def detect_excel_format(df: pd.DataFrame) -> str:
    """For backward compatibility"""
    return 'search-based'


def parse_hybrid_excel(df: pd.DataFrame, sheet_name: str = None, use_namespacing: bool = False) -> dict:
    """For backward compatibility - now uses search-based approach"""
    return extract_all_data_from_sheet(df, sheet_name)


def parse_row_based_excel(df: pd.DataFrame, sheet_name: str = None, use_namespacing: bool = False) -> dict:
    """For backward compatibility - now uses search-based approach"""
    return extract_all_data_from_sheet(df, sheet_name)


def parse_multi_sheet_excel(excel_file, use_namespacing: bool = False) -> tuple:
    """For backward compatibility - now uses search-based approach"""
    return parse_multi_sheet_excel_search_based(excel_file)
