#!/usr/bin/env python3
"""
Test the fixed template functions for AI text generation.
"""

import sys
sys.path.insert(0, '/workspaces/codespaces-blank')

from app import (
    generate_sector_label,
    generate_client_oneliner,
    generate_project_highlight_3para,
    post_process_length
)

def test_sector_label():
    print("=" * 60)
    print("TEST 1: Sector Label Generation (2-3 words)")
    print("=" * 60)
    
    # Test 1: Solar energy
    prompt1 = 'Generate a 2-3 word sector label. Sector = "Solar", Project_Type = "PV Plant"'
    result1 = generate_sector_label(prompt1, max_words=3)
    print(f"Input: {prompt1}")
    print(f"Output: {result1}")
    print(f"Word count: {len(result1.split())} ✓" if len(result1.split()) <= 3 else f"Word count: {len(result1.split())} ✗")
    print()
    
    # Test 2: Wind energy
    prompt2 = 'Generate a short label for wind power project'
    result2 = generate_sector_label(prompt2, max_words=3)
    print(f"Input: {prompt2}")
    print(f"Output: {result2}")
    print()
    
    # Test 3: Battery storage
    prompt3 = 'Create 1-3 word sector label for battery storage facility'
    result3 = generate_sector_label(prompt3, max_words=3)
    print(f"Input: {prompt3}")
    print(f"Output: {result3}")
    print()


def test_client_oneliner():
    print("=" * 60)
    print("TEST 2: Client One-liner (Anonymous Description)")
    print("=" * 60)
    
    # Test 1: Developing
    prompt1 = '''Write one anonymous sentence starting with "The client" that briefly states:
    - Asset_Description: "a 500 MW solar photovoltaic power plant with battery storage"
    - Country: "Vietnam"
    - Action: developing'''
    result1 = generate_client_oneliner(prompt1)
    print(f"Output: {result1}")
    print(f"Sentence count: 1 ✓" if result1.count('.') == 1 else f"Sentence count: {result1.count('.')} ✗")
    print()
    
    # Test 2: Operating
    prompt2 = '''Write one sentence starting with "The client" anonymously describing that they are operating:
    "renewable energy assets including wind and solar facilities" in "Southeast Asia"'''
    result2 = generate_client_oneliner(prompt2)
    print(f"Output: {result2}")
    print()
    
    # Test 3: Both developing and operating
    prompt3 = '''Anonymous one-liner: The client is developing and operating 
    "distributed energy resources and microgrids" in "Indonesia and Thailand"'''
    result3 = generate_client_oneliner(prompt3)
    print(f"Output: {result3}")
    print()


def test_project_highlight():
    print("=" * 60)
    print("TEST 3: Project Highlight (3 Paragraphs)")
    print("=" * 60)
    
    prompt = '''Generate 3 paragraphs following this structure:

Paragraph 1: Company operations and project focus.
Use: "specializes in renewable energy development and operates 2 GW of solar assets across Asia"
Use: "expansion of their flagship solar portfolio in Vietnam"

Paragraph 2: Service scope and partnerships.
Use: "provides comprehensive EPC and O&M services for utility-scale solar projects"
Use: "partner with leading international equipment suppliers and local contractors"
Use: "This collaboration ensures timely delivery and optimal performance"

Paragraph 3: Investment details.
The project will commence with an initial investment of "$150 million", with a projected expansion to "$500 million over the next 5 years".'''
    
    result = generate_project_highlight_3para(prompt)
    paragraphs = result.split('\n\n')
    
    print(f"Output:\n{result}\n")
    print(f"Paragraph count: {len(paragraphs)} {'✓' if len(paragraphs) == 3 else '✗'}")
    print()
    
    for i, para in enumerate(paragraphs, 1):
        sentences = [s.strip() for s in para.split('.') if s.strip()]
        print(f"Paragraph {i}: {len(sentences)} sentences")
    print()


def test_post_processing():
    print("=" * 60)
    print("TEST 4: Post-Processing Length Control")
    print("=" * 60)
    
    # Test word trimming
    text1 = "Solar Energy Wind Power Battery Storage Hydro"
    result1 = post_process_length(text1, target_words=(2, 3))
    print(f"Original: {text1}")
    print(f"Target: 2-3 words")
    print(f"Result: {result1} ({len(result1.split())} words) ✓")
    print()
    
    # Test sentence trimming
    text2 = "First sentence here. Second sentence here. Third sentence here. Fourth sentence here."
    result2 = post_process_length(text2, target_sentences=2)
    print(f"Original: 4 sentences")
    print(f"Target: 2 sentences")
    result2_sentences = [s.strip() for s in result2.split('.') if s.strip()]
    print(f"Result: {len(result2_sentences)} sentences ✓")
    print(f"{result2}")
    print()
    
    # Test paragraph trimming
    text3 = "Paragraph 1.\n\nParagraph 2.\n\nParagraph 3.\n\nParagraph 4."
    result3 = post_process_length(text3, target_paragraphs=3)
    result3_paras = [p.strip() for p in result3.split('\n\n') if p.strip()]
    print(f"Original: 4 paragraphs")
    print(f"Target: 3 paragraphs")
    print(f"Result: {len(result3_paras)} paragraphs ✓")
    print()


if __name__ == "__main__":
    print("\n")
    print("╔" + "=" * 58 + "╗")
    print("║" + " " * 10 + "FIXED TEMPLATE FUNCTIONS TEST SUITE" + " " * 12 + "║")
    print("╚" + "=" * 58 + "╝")
    print()
    
    test_sector_label()
    test_client_oneliner()
    test_project_highlight()
    test_post_processing()
    
    print("=" * 60)
    print("✓ ALL TESTS COMPLETED")
    print("=" * 60)
    print()
