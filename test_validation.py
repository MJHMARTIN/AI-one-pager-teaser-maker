#!/usr/bin/env python3
"""
Test validation and regeneration logic for AI text generation.
"""

import sys
sys.path.insert(0, '/workspaces/codespaces-blank')

import pandas as pd
from app import (
    extract_company_names,
    validate_anonymous_text,
    validate_client_oneliner,
    make_stricter_prompt,
    generate_beta_text
)

def test_extract_company_names():
    print("=" * 60)
    print("TEST 1: Extract Company Names")
    print("=" * 60)
    
    # Create sample data
    data = pd.Series({
        'Company Name': 'ABC Solar Corp',
        'Sponsor': 'XYZ Energy Group',
        'Location': 'Vietnam',
        'Capacity': '500 MW'
    })
    
    names = extract_company_names(data)
    print(f"Input data:\n{data}\n")
    print(f"Extracted names: {names}")
    print(f"Should contain: 'abc', 'solar', 'corp', 'xyz', 'energy', 'group' ✓")
    print()


def test_validate_anonymous_text():
    print("=" * 60)
    print("TEST 2: Validate Anonymous Text")
    print("=" * 60)
    
    data = pd.Series({
        'Company Name': 'SunPower Solutions',
        'Location': 'Vietnam'
    })
    
    # Test 1: Valid anonymous text
    text1 = "The client is developing renewable energy projects in Vietnam."
    is_valid1, reason1 = validate_anonymous_text(text1, data)
    print(f"Text: {text1}")
    print(f"Valid: {is_valid1} {'✓' if is_valid1 else '✗'}")
    if not is_valid1:
        print(f"Reason: {reason1}")
    print()
    
    # Test 2: Invalid - contains company name
    text2 = "SunPower Solutions is developing renewable energy projects in Vietnam."
    is_valid2, reason2 = validate_anonymous_text(text2, data)
    print(f"Text: {text2}")
    print(f"Valid: {is_valid2} {'✓' if not is_valid2 else '✗'} (should be False)")
    print(f"Reason: {reason2}")
    print()


def test_validate_client_oneliner():
    print("=" * 60)
    print("TEST 3: Validate Client One-liner")
    print("=" * 60)
    
    data = pd.Series({
        'Company Name': 'ABC Corp',
        'Asset': 'Solar Plant'
    })
    
    # Test 1: Valid
    text1 = "The client is developing a 500 MW solar facility in Vietnam."
    is_valid1, reason1 = validate_client_oneliner(text1, data)
    print(f"Text: {text1}")
    print(f"Valid: {is_valid1} {'✓' if is_valid1 else '✗'}")
    if not is_valid1:
        print(f"Reason: {reason1}")
    print()
    
    # Test 2: Invalid - doesn't start with "The client"
    text2 = "A company is developing solar projects."
    is_valid2, reason2 = validate_client_oneliner(text2, data)
    print(f"Text: {text2}")
    print(f"Valid: {is_valid2} {'✓' if not is_valid2 else '✗'} (should be False)")
    print(f"Reason: {reason2}")
    print()
    
    # Test 3: Invalid - multiple sentences
    text3 = "The client is developing projects. They have experience."
    is_valid3, reason3 = validate_client_oneliner(text3, data)
    print(f"Text: {text3}")
    print(f"Valid: {is_valid3} {'✓' if not is_valid3 else '✗'} (should be False)")
    print(f"Reason: {reason3}")
    print()
    
    # Test 4: Invalid - contains company name
    text4 = "The client ABC Corp is developing solar projects."
    is_valid4, reason4 = validate_client_oneliner(text4, data)
    print(f"Text: {text4}")
    print(f"Valid: {is_valid4} {'✓' if not is_valid4 else '✗'} (should be False)")
    print(f"Reason: {reason4}")
    print()


def test_stricter_prompt():
    print("=" * 60)
    print("TEST 4: Generate Stricter Prompt")
    print("=" * 60)
    
    original = '''Write one anonymous sentence starting with "The client" about:
    Asset: "renewable energy projects"
    Country: "Vietnam"'''
    
    failure = "Does not start with 'The client'"
    
    stricter = make_stricter_prompt(original, failure)
    print(f"Original prompt:\n{original}\n")
    print(f"Failure reason: {failure}\n")
    print(f"Stricter prompt:\n{stricter}\n")
    print(f"Contains 'IMPORTANT': {'✓' if 'IMPORTANT' in stricter else '✗'}")
    print()


def test_generate_with_validation():
    print("=" * 60)
    print("TEST 5: Generate with Validation & Regeneration")
    print("=" * 60)
    
    # Create sample data that might leak
    data = pd.Series({
        'Company Name': 'TestCompany Inc',
        'Asset Description': 'a 500 MW solar photovoltaic power plant',
        'Country': 'Vietnam'
    })
    
    # Test client one-liner generation
    prompt = '''Write one anonymous sentence starting with "The client" that briefly states:
    - Asset_Description: "a 500 MW solar photovoltaic power plant"
    - Country: "Vietnam"
    - Action: developing'''
    
    result = generate_beta_text(prompt, data, tone="short")
    
    print(f"Prompt: {prompt}\n")
    print(f"Generated: {result}\n")
    
    # Validate the result
    is_valid, reason = validate_client_oneliner(result, data)
    print(f"Validation: {is_valid} {'✓' if is_valid else '✗'}")
    if not is_valid:
        print(f"Reason: {reason}")
    else:
        print("✓ Starts with 'The client'")
        print("✓ Single sentence")
        print("✓ Maintains anonymity")
    print()


def test_edge_cases():
    print("=" * 60)
    print("TEST 6: Edge Cases")
    print("=" * 60)
    
    # Case 1: Empty data
    data1 = pd.Series({})
    text1 = "The client is developing projects."
    is_valid1, _ = validate_client_oneliner(text1, data1)
    print(f"Empty data validation: {is_valid1} ✓")
    
    # Case 2: Missing "The client" but otherwise good
    data2 = pd.Series({'Company': 'XYZ'})
    text2 = "Client is developing renewable energy."
    is_valid2, reason2 = validate_client_oneliner(text2, data2)
    print(f"Missing 'The client': {is_valid2} (should be False) ✓")
    print(f"Reason: {reason2}")
    
    # Case 3: Common words should not trigger false positives
    data3 = pd.Series({'Company Name': 'The Group'})
    text3 = "The client is working with the group on projects."
    is_valid3, _ = validate_anonymous_text(text3, data3)
    print(f"Common word 'group': {is_valid3} (should handle gracefully)")
    print()


if __name__ == "__main__":
    print("\n")
    print("╔" + "=" * 58 + "╗")
    print("║" + " " * 8 + "VALIDATION & REGENERATION TEST SUITE" + " " * 13 + "║")
    print("╚" + "=" * 58 + "╝")
    print()
    
    test_extract_company_names()
    test_validate_anonymous_text()
    test_validate_client_oneliner()
    test_stricter_prompt()
    test_generate_with_validation()
    test_edge_cases()
    
    print("=" * 60)
    print("✓ ALL VALIDATION TESTS COMPLETED")
    print("=" * 60)
    print()
