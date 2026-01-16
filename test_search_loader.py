#!/usr/bin/env python3
"""
Test the new search-based Excel loader
"""

import sys
sys.path.insert(0, '/workspaces/codespaces-blank')

import pandas as pd
from excel_loader_new import (
    parse_multi_sheet_excel_search_based,
    search_label_in_sheet,
    extract_all_data_from_sheet,
    normalize_label
)
from field_mapping import get_module_field_mappings
import logging

logging.basicConfig(level=logging.INFO, format='%(levelname)s - %(message)s')

print("="*80)
print("TEST: Search-Based Excel Loader")
print("="*80)

# Test with the working test file first
print("\n1. Testing with test_module1_sponsor.xlsx")
print("-"*80)

unified_data, sheet_info = parse_multi_sheet_excel_search_based('test_module1_sponsor.xlsx')

print(f"\n✅ Found {len(unified_data)} fields total")
print(f"\nSheet info:")
for info in sheet_info:
    print(f"  - {info['name']}: {info['format']}, {info['fields']} fields")

print(f"\nExtracted fields:")
for key, value in sorted(unified_data.items()):
    print(f"  '{key}' = '{value[:60] if len(value) > 60 else value}'")

# Now test targeted search for Module1 labels
print("\n" + "="*80)
print("2. Testing Module1 Label Search")
print("="*80)

module_mappings = get_module_field_mappings()
module1_map = module_mappings['Module1']

print(f"\nSearching for Module1 specific labels:")
for tag, excel_label in module1_map.items():
    normalized = normalize_label(excel_label)
    found = normalized in unified_data
    status = "✅" if found else "❌"
    value = unified_data.get(normalized, "(NOT FOUND)")[:50]
    print(f"  {status} {tag:20s} '{excel_label}' → '{value}'")

# Test with user's actual file if it exists
print("\n" + "="*80)
print("3. Testing with your actual Excel file (if present)")
print("="*80)

import glob
user_files = glob.glob('DCG*.xlsm') + glob.glob('DCG*.xlsx')

if user_files:
    test_file = user_files[0]
    print(f"\nFound: {test_file}")
    print("Attempting to parse...")
    
    try:
        unified_data, sheet_info = parse_multi_sheet_excel_search_based(test_file)
        
        print(f"\n✅ Parsed {len(unified_data)} fields from {len(sheet_info)} sheets")
        
        print(f"\nSheet breakdown:")
        for info in sheet_info:
            print(f"  - {info['name']}: {info['fields']} fields")
        
        print(f"\nFirst 20 fields found:")
        for i, (key, value) in enumerate(list(unified_data.items())[:20], 1):
            value_preview = value[:50] if len(value) > 50 else value
            print(f"  {i:2d}. '{key}' = '{value_preview}'")
        
        if len(unified_data) > 20:
            print(f"\n  ... and {len(unified_data) - 20} more fields")
        
        # Check for Module1 matches
        print(f"\nChecking Module1 label matches:")
        matches = 0
        for tag, excel_label in module1_map.items():
            normalized = normalize_label(excel_label)
            if normalized in unified_data:
                matches += 1
                print(f"  ✅ Found: '{excel_label}'")
        
        print(f"\n{matches}/{len(module1_map)} Module1 fields found")
        
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()
else:
    print("\nNo DCG*.xlsm or DCG*.xlsx files found in current directory")
    print("Upload your file and run this test again")

print("\n" + "="*80)
print("TEST COMPLETE")
print("="*80)
