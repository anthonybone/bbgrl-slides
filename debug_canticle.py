"""
Debug script to check Canticle red text for November 10, 2025
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
    print("Looking for Canticle red text:")
    print("="*70)
    
    # Find "Canticle" in the text
    canticle_pos = text_after_psalmody.find('Canticle:')
    if canticle_pos >= 0:
        print(f"\nFound 'Canticle:' at position {canticle_pos}")
        
        # Show context
        start = max(0, canticle_pos - 100)
        end = min(len(text_after_psalmody), canticle_pos + 300)
        context = text_after_psalmody[start:end]
        print(f"\nContext:\n{context}")
        
        # Try to extract just the canticle title line
        canticle_section = text_after_psalmody[canticle_pos:canticle_pos+500]
        lines = canticle_section.split('\n')
        print(f"\nFirst few lines after 'Canticle:':")
        for i, line in enumerate(lines[:5]):
            print(f"  Line {i}: '{line.strip()}'")
    
    # Look in HTML for canticle rubrica spans
    print("\n" + "="*70)
    print("Looking for Canticle in HTML rubrica spans:")
    print("="*70)
    
    for span in soup.find_all('span', class_='rubrica'):
        text = span.get_text().strip()
        if 'Canticle' in text or 'Isaiah' in text:
            print(f"\nFound rubrica: {text}")
            
            # Get next few siblings
            current = span.next_sibling
            context = ""
            for _ in range(5):
                if current:
                    if hasattr(current, 'get_text'):
                        context += current.get_text()
                    else:
                        context += str(current)
                    current = current.next_sibling
            
            print(f"Context after: {context[:200]}...")
else:
    print("Failed to fetch HTML content")
