"""
Debug script to examine Ant. 2 parsing for November 10, 2025
"""

from bbgrl_slide_generator_v2_template import BBGRLSlideGeneratorV2
from datetime import datetime
from bs4 import BeautifulSoup
import re

# Initialize generator
generator = BBGRLSlideGeneratorV2()

# Navigate to November 10, 2025
target_date = datetime(2025, 11, 10)
print(f"Fetching Morning Prayer for {target_date.strftime('%B %d, %Y')}...")

html_content = generator._navigate_ibreviary_to_date(target_date)

if html_content:
    soup = BeautifulSoup(html_content, 'html.parser')
    full_text = soup.get_text()
    
    # Find PSALMODY position
    psalmody_pos = full_text.upper().find('PSALMODY')
    if psalmody_pos >= 0:
        text_after_psalmody = full_text[psalmody_pos:]
    else:
        text_after_psalmody = full_text
    
    print("\n" + "="*70)
    print("Testing _extract_antiphon function for Ant. 2:")
    print("="*70)
    
    # Test the function
    result = generator._extract_antiphon(text_after_psalmody, 2)
    print(f"\nExtracted antiphon_2:")
    print(f"Text: {result['text']}")
    print(f"Format: {result['format']}")
    
    # Now let's manually search for "Ant. 2" patterns
    print("\n" + "="*70)
    print("Manual pattern search for Ant. 2:")
    print("="*70)
    
    patterns = [
        r'Ant\.\s*2[:\s]+([^.]+\.)',
        r'Antiphon\s*2[:\s]+([^.]+\.)'
    ]
    
    for pattern in patterns:
        print(f"\nPattern: {pattern}")
        matches = list(re.finditer(pattern, text_after_psalmody, re.IGNORECASE))
        print(f"Found {len(matches)} matches")
        
        for i, match in enumerate(matches[:5]):  # Show first 5 matches
            print(f"\nMatch {i+1}:")
            print(f"  Position: {match.start()}")
            print(f"  Full match: {match.group(0)}")
            print(f"  Captured text: {match.group(1)}")
            
            # Show context around the match
            start_ctx = max(0, match.start() - 100)
            end_ctx = min(len(text_after_psalmody), match.end() + 200)
            context = text_after_psalmody[start_ctx:end_ctx]
            print(f"  Context: ...{context}...")
    
    # Search for all "Ant. 2" occurrences
    print("\n" + "="*70)
    print("All 'Ant. 2' occurrences in text:")
    print("="*70)
    
    pos = 0
    while True:
        pos = text_after_psalmody.find('Ant. 2', pos)
        if pos < 0:
            break
        
        print(f"\nFound 'Ant. 2' at position {pos}")
        start = max(0, pos - 50)
        end = min(len(text_after_psalmody), pos + 150)
        print(f"Context: ...{text_after_psalmody[start:end]}...")
        pos += 1
else:
    print("Failed to fetch HTML content")
