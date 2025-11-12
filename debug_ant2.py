"""
Debug script to check if Ant. 2 is available in iBreviary HTML for November 10, 2025
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
    
    # Find all rubrica spans (red text)
    print("\n" + "="*70)
    print("All 'Ant.' references found in the HTML:")
    print("="*70)
    
    for span in soup.find_all('span', class_='rubrica'):
        text = span.get_text().strip()
        if 'Ant.' in text:
            print(f"\n{text}")
            # Get some context - next few elements
            next_text = ""
            current = span.next_sibling
            for _ in range(3):  # Get next 3 siblings
                if current:
                    if hasattr(current, 'get_text'):
                        next_text += current.get_text()
                    else:
                        next_text += str(current)
                    current = current.next_sibling
            
            # Clean up and show context
            context = next_text.strip()[:200]  # First 200 chars
            if context:
                print(f"Context: {context}...")
    
    # Specifically look for Ant. 2
    print("\n" + "="*70)
    print("Searching specifically for 'Ant. 2':")
    print("="*70)
    
    ant2_pattern = r'Ant\.\s*2'
    ant2_spans = soup.find_all('span', class_='rubrica', string=re.compile(ant2_pattern))
    
    if ant2_spans:
        for span in ant2_spans:
            print(f"\nFound: {span.get_text().strip()}")
            
            # Get the full antiphon text (next non-rubrica text)
            current = span.next_sibling
            antiphon_text = ""
            
            while current and len(antiphon_text) < 500:
                if hasattr(current, 'get_text'):
                    if current.name == 'span' and 'rubrica' in current.get('class', []):
                        break  # Stop at next rubrica
                    antiphon_text += current.get_text()
                else:
                    antiphon_text += str(current)
                current = current.next_sibling
            
            # Clean up the text
            antiphon_text = re.sub(r'\s+', ' ', antiphon_text).strip()
            # Find first sentence
            match = re.search(r'^([^.]+\.)', antiphon_text)
            if match:
                print(f"Antiphon text: {match.group(1)}")
            else:
                print(f"Antiphon text: {antiphon_text[:200]}...")
    else:
        print("No 'Ant. 2' found in the HTML")
        
        # Check full text for Ant. 2
        full_text = soup.get_text()
        if 'Ant. 2' in full_text or 'Ant.2' in full_text:
            print("\nBut 'Ant. 2' found in full text!")
            # Find position and show context
            pos = full_text.find('Ant. 2')
            if pos < 0:
                pos = full_text.find('Ant.2')
            if pos >= 0:
                context = full_text[max(0, pos-50):pos+200]
                print(f"Context: ...{context}...")
else:
    print("Failed to fetch HTML content")
