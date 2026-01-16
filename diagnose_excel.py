"""
Diagnostic tool to verify Excel parsing and placeholder resolution.
Run this to check if your Excel file is being read correctly.
"""

import pandas as pd
import sys
from excel_loader import parse_hybrid_excel, normalize_label
from module_detector import detect_module_type
from field_mapping import get_module_field_mappings, create_placeholder_mapping
from ppt_filler import resolve_tag_to_value

def diagnose_excel_file(excel_path):
    """Diagnose an Excel file to verify it's being read correctly."""
    
    print("="*80)
    print(f"DIAGNOSING EXCEL FILE: {excel_path}")
    print("="*80)
    
    # Read the Excel file
    try:
        df = pd.read_excel(excel_path, engine='openpyxl')
        print(f"\n✅ Successfully read Excel file")
        print(f"   Shape: {df.shape[0]} rows × {df.shape[1]} columns")
    except Exception as e:
        print(f"\n❌ Error reading Excel file: {e}")
        return
    
    # Show the raw structure
    print(f"\n📊 RAW EXCEL STRUCTURE:")
    print(f"   Column names: {df.columns.tolist()}")
    print(f"\n   First 10 rows:")
    print(df.head(10).to_string(index=True))
    
    # Parse with hybrid parser
    print(f"\n" + "="*80)
    print(f"PARSING WITH HYBRID PARSER (reads both column & row formats)")
    print("="*80)
    
    data_dict = parse_hybrid_excel(df)
    print(f"\n✅ Parsed {len(data_dict)} fields")
    
    # Show all parsed fields
    print(f"\n📋 ALL PARSED FIELDS:")
    for i, (key, value) in enumerate(sorted(data_dict.items()), 1):
        value_preview = value[:60] + "..." if len(value) > 60 else value
        print(f"   {i:2d}. '{key}' = '{value_preview}'")
    
    # Detect module type
    print(f"\n" + "="*80)
    print(f"MODULE DETECTION")
    print("="*80)
    
    module_type = detect_module_type(data_dict)
    if module_type:
        print(f"\n✅ Detected: {module_type}")
    else:
        print(f"\n⚠️  No specific module detected (using generic mapping)")
    
    # Show module field matching
    module_mappings = get_module_field_mappings()
    for module_name, field_map in module_mappings.items():
        matches = 0
        for tag, excel_label in field_map.items():
            normalized = normalize_label(excel_label)
            if normalized in data_dict:
                matches += 1
        
        status = "✅" if module_name == module_type else "  "
        print(f"   {status} {module_name}: {matches}/{len(field_map)} fields matched")
    
    # Test placeholder resolution
    print(f"\n" + "="*80)
    print(f"TESTING PLACEHOLDER RESOLUTION")
    print("="*80)
    
    placeholder_mapping = create_placeholder_mapping()
    
    # Common tags to test
    test_tags = [
        "ISSUER", "COMPANY_NAME", "SPONSOR_NAME",
        "INDUSTRY", "JURISDICTION", "ISSUANCE_TYPE",
        "INITIAL_NOTIONAL", "COUPON_RATE", "COUPON_FREQUENCY",
        "TENOR", "CLIENT_SUMMARY", "PROJECT_HIGHLIGHT"
    ]
    
    print(f"\n   Testing common :TAG: placeholders:")
    for tag in test_tags:
        value = resolve_tag_to_value(tag, data_dict, placeholder_mapping, module_type)
        if value:
            value_preview = value[:50] + "..." if len(value) > 50 else value
            print(f"      ✅ :{tag}: → '{value_preview}'")
        else:
            print(f"      ❌ :{tag}: → (NOT FOUND)")
    
    # Show what fields are available that aren't being matched
    print(f"\n" + "="*80)
    print(f"SUMMARY")
    print("="*80)
    
    if len(data_dict) == 0:
        print(f"\n❌ NO DATA PARSED! This is the issue.")
        print(f"   Your Excel might have an unexpected format.")
        print(f"   Expected formats:")
        print(f"   - Row-based: Column 0 = Labels, Column 1 = Values")
        print(f"   - Column-based: Row 0 = Headers, Row 1 = Values")
    else:
        print(f"\n✅ Excel is being read correctly!")
        print(f"   - {len(data_dict)} fields found")
        print(f"   - Module type: {module_type or 'Generic'}")
        print(f"\n   If placeholders are still empty, check:")
        print(f"   1. PowerPoint template uses correct format:")
        print(f"      - For row-based Excel: Use :TAG: format (e.g., :ISSUER:)")
        print(f"      - For column-based Excel: Use [ColumnName] format")
        print(f"   2. Tag names match expected values")
        print(f"   3. Check application logs for warnings")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python diagnose_excel.py <path_to_excel_file>")
        print("\nExample:")
        print("  python diagnose_excel.py test_module1_sponsor.xlsx")
        sys.exit(1)
    
    excel_path = sys.argv[1]
    diagnose_excel_file(excel_path)
