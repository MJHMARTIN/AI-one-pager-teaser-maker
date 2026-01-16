#!/usr/bin/env python3
"""
Test the mapping and replacement logic with hard-coded values
to isolate where the problem is.
"""

import sys
sys.path.insert(0, '/workspaces/codespaces-blank')

from excel_loader import parse_hybrid_excel, normalize_label
from module_detector import detect_module_type
from field_mapping import create_placeholder_mapping, get_module_field_mappings
from ppt_filler import resolve_tag_to_value
import pandas as pd

print("="*80)
print("TEST 1: Verify Excel Parsing")
print("="*80)

# Read a working test file
df = pd.read_excel('test_module1_sponsor.xlsx', engine='openpyxl')
data_dict = parse_hybrid_excel(df)

print(f"\n✅ Parsed {len(data_dict)} fields from Excel")
print(f"\nActual data_dict keys and values:")
for key, value in sorted(data_dict.items()):
    print(f"  '{key}' = '{value[:60] if len(value) > 60 else value}'")

print("\n" + "="*80)
print("TEST 2: Module Detection")
print("="*80)

module_type = detect_module_type(data_dict)
print(f"\n✅ Detected module type: {module_type}")

print("\n" + "="*80)
print("TEST 3: Module Field Mappings")
print("="*80)

module_mappings = get_module_field_mappings()
if module_type and module_type in module_mappings:
    print(f"\n{module_type} expects these Excel labels:")
    for tag, excel_label in module_mappings[module_type].items():
        normalized = normalize_label(excel_label)
        found = "✅" if normalized in data_dict else "❌"
        print(f"  {found} {tag:20s} → '{excel_label}' (normalized: '{normalized}')")

print("\n" + "="*80)
print("TEST 4: Tag Resolution")
print("="*80)

placeholder_mapping = create_placeholder_mapping()

test_tags = ["ISSUER", "INDUSTRY", "JURISDICTION", "ISSUANCE_TYPE", "COUPON_RATE"]

print(f"\nTesting tag resolution:")
for tag in test_tags:
    value = resolve_tag_to_value(tag, data_dict, placeholder_mapping, module_type)
    status = "✅" if value else "❌"
    print(f"  {status} :{tag}: → '{value[:60] if value else '(EMPTY)'}'")

print("\n" + "="*80)
print("TEST 5: Check Exact Key Matching")
print("="*80)

# Check if there's a case/spacing issue
module_label = "sponsor name"
normalized = normalize_label(module_label)
print(f"\nLooking for: '{module_label}'")
print(f"Normalized:  '{normalized}'")
print(f"In dict? {normalized in data_dict}")

if normalized in data_dict:
    print(f"✅ FOUND! Value: '{data_dict[normalized]}'")
else:
    print(f"❌ NOT FOUND")
    print(f"\nClosest matches in data_dict:")
    from difflib import get_close_matches
    matches = get_close_matches(normalized, list(data_dict.keys()), n=5, cutoff=0.5)
    for match in matches:
        print(f"  - '{match}' (similarity match)")

print("\n" + "="*80)
print("TEST 6: Simulate PPT Placeholder Replacement")
print("="*80)

# Simulate what happens in the PPT
print("\nSimulating PPT placeholder replacement:")
ppt_placeholders = [":ISSUER:", ":INDUSTRY:", ":JURISDICTION:"]

for placeholder in ppt_placeholders:
    # Extract tag name
    tag = placeholder.strip(':')
    # Resolve it
    value = resolve_tag_to_value(tag, data_dict, placeholder_mapping, module_type)
    # Show result
    print(f"  {placeholder:20s} → '{value if value else '(EMPTY - THIS IS THE PROBLEM!)'}'")

print("\n" + "="*80)
print("SUMMARY")
print("="*80)

all_resolved = all(
    resolve_tag_to_value(tag, data_dict, placeholder_mapping, module_type)
    for tag in test_tags
)

if all_resolved:
    print("\n✅ ALL TESTS PASSED - Mapping works correctly!")
    print("   If PPT is still empty, the problem is:")
    print("   1. PPT template doesn't have these exact placeholders")
    print("   2. PPT traversal code isn't finding the text")
    print("   3. Different data_dict is being used in actual run")
else:
    print("\n❌ MAPPING FAILED - Some tags don't resolve")
    print("   This is where the problem is!")
    print("   Check:")
    print("   1. Excel labels match what module expects")
    print("   2. normalize_label() is working correctly")
    print("   3. data_dict actually has the data")

print()
