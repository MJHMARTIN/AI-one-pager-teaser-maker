"""
Excel Loader Module
Handles reading and normalizing Excel files in various formats.
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


def detect_excel_format(df: pd.DataFrame) -> str:
    """
    Detect if Excel is column-based or row-based format.
    
    Args:
        df: DataFrame to analyze
    
    Returns:
        'column-based' or 'row-based'
    """
    logger.debug(f"Detecting format for DataFrame with shape {df.shape}")
    
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
            logger.info("Detected row-based format from column names")
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
                logger.info(f"Detected row-based format from content ({label_like_count} label-like rows)")
                return 'row-based'
    
    # Default to column-based
    logger.info("Detected column-based format (default)")
    return 'column-based'


def parse_hybrid_excel(df: pd.DataFrame, sheet_name: str = None, use_namespacing: bool = False) -> dict:
    """
    Parse Excel in BOTH column-based and row-based formats simultaneously.
    
    - Column-based: First row (headers) become field names, second row has values
    - Row-based: First column = labels, second column = values
    
    This hybrid approach handles any Excel structure.
    
    Args:
        df: DataFrame to parse
        sheet_name: Name of the sheet (for namespacing)
        use_namespacing: If True, prefix keys with sheet name
    
    Returns:
        dict: {normalized_label: value}
    """
    data_dict = {}
    
    logger.debug(f"Parsing hybrid Excel (sheet: {sheet_name}, namespacing: {use_namespacing})")
    
    # PART 1: Read as COLUMN-based (headers = field names)
    # Skip this if columns look like generic label/value headers
    skip_column_based = False
    if len(df.columns) >= 2:
        first_col = str(df.columns[0]).lower()
        second_col = str(df.columns[1]).lower()
        # If columns are named "Label" and "Value" (or similar), skip column-based parsing
        if first_col in ['label', 'field', 'name', 'key'] and second_col in ['value', 'data', 'answer']:
            skip_column_based = True
            logger.debug("Detected Label/Value format - skipping column-based parsing")
    
    if not skip_column_based and len(df) > 0 and len(df.columns) > 0:
        # Use first data row as values
        first_row = df.iloc[0]
        for col_name in df.columns:
            if pd.isna(col_name) or str(col_name).strip() == '':
                continue
            
            normalized = normalize_label(str(col_name))
            if use_namespacing and sheet_name:
                sheet_prefix = normalize_label(sheet_name)
                normalized = f"{sheet_prefix}.{normalized}"
            
            value = first_row[col_name]
            data_dict[normalized] = "" if pd.isna(value) else str(value)
    
    logger.debug(f"Column-based parsing found {len(data_dict)} fields")
    
    # PART 2: Read as ROW-based (column 0 = labels, column 1 = values)
    if len(df.columns) >= 2:
        label_col = df.columns[0]
        value_col = df.columns[1]
        
        row_based_count = 0
        for idx, row in df.iterrows():
            label = row[label_col]
            value = row[value_col]
            
            # Skip empty labels
            if pd.isna(label) or str(label).strip() == '':
                continue
            
            normalized = normalize_label(str(label))
            if use_namespacing and sheet_name:
                sheet_prefix = normalize_label(sheet_name)
                normalized = f"{sheet_prefix}.{normalized}"
            
            # Row-based takes precedence if there's overlap
            data_dict[normalized] = "" if pd.isna(value) else str(value)
            row_based_count += 1
        
        logger.debug(f"Row-based parsing found {row_based_count} additional fields")
    
    logger.info(f"Hybrid parsing complete: {len(data_dict)} total fields")
    return data_dict


def parse_row_based_excel(df: pd.DataFrame, sheet_name: str = None, use_namespacing: bool = False) -> dict:
    """
    Parse row-based Excel format into a normalized key-value dictionary.
    
    Expects Excel with structure:
    - Column 0: Labels/Questions
    - Column 1: Values/Answers
    
    Args:
        df: DataFrame to parse
        sheet_name: Name of the sheet (for namespacing)
        use_namespacing: If True, prefix keys with sheet name
    
    Returns:
        dict: {normalized_label: value}
    """
    data_dict = {}
    
    logger.debug(f"Parsing row-based Excel (sheet: {sheet_name})")
    
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
        
        # Add sheet namespacing if requested
        if use_namespacing and sheet_name:
            sheet_prefix = normalize_label(sheet_name)
            normalized = f"{sheet_prefix}.{normalized}"
        
        # Store the value (convert NaN to empty string)
        data_dict[normalized] = "" if pd.isna(value) else str(value)
    
    logger.info(f"Row-based parsing complete: {len(data_dict)} fields")
    return data_dict


def parse_multi_sheet_excel(excel_file, use_namespacing: bool = False) -> tuple:
    """
    Read all sheets from an Excel file and merge into unified knowledge base.
    
    Args:
        excel_file: File-like object or path to Excel (.xlsx or .xlsm)
        use_namespacing: If True, prefix keys with sheet names
    
    Returns:
        tuple: (unified_data_dict, sheet_info_list)
    """
    logger.info("Reading multi-sheet Excel file")
    
    # Read all sheets (supports both .xlsx and .xlsm, macros ignored)
    sheets_dict = pd.read_excel(excel_file, sheet_name=None, engine='openpyxl')
    
    logger.info(f"Found {len(sheets_dict)} sheets")
    
    unified_data = {}
    sheet_info = []
    
    for sheet_name, df in sheets_dict.items():
        if df.empty:
            logger.warning(f"Sheet '{sheet_name}' is empty, skipping")
            continue
        
        logger.debug(f"Processing sheet '{sheet_name}' with shape {df.shape}")
        
        # In multi-sheet mode, try to parse everything as HYBRID
        try:
            sheet_data = parse_hybrid_excel(df, sheet_name, use_namespacing)
            
            if sheet_data:  # Only count if we got data
                # Merge into unified dict
                unified_data.update(sheet_data)
                
                sheet_info.append({
                    'name': sheet_name,
                    'format': 'hybrid (row+column)',
                    'fields': len(sheet_data)
                })
                
                logger.info(f"Sheet '{sheet_name}': {len(sheet_data)} fields extracted")
            else:
                sheet_info.append({
                    'name': sheet_name,
                    'format': 'empty or invalid',
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
    
    logger.info(f"Multi-sheet parsing complete: {len(unified_data)} total fields from {len(sheet_info)} sheets")
    return unified_data, sheet_info
