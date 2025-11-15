"""Test HTML-based psalm extraction"""
from bs4 import BeautifulSoup
import re

# Load the debug HTML file
with open('debug_morning_prayer.html', 'r', encoding='utf-8') as f:
    html = f.read()

soup = BeautifulSoup(html, 'html.parser')

# Find Ant. 1
ant_1 = soup.find('span', class_='rubrica', string=re.compile(r'Ant\.\s*1'))
if ant_1:
    print("Found Ant. 1")
    print(f"Ant. 1 text: {ant_1.get_text()}")
    
    # Get the next few siblings
    current = ant_1
    count = 0
    print("\nNext 10 siblings:")
    while current and count < 10:
        current = current.find_next_sibling()
        if current:
            count += 1
            print(f"{count}. Tag: {current.name}, Class: {current.get('class')}, Text: {current.get_text()[:100]}")
            
    # Now try to collect text until next Ant or Glory
    print("\n\nCollecting psalm text:")
    current = ant_1
    psalm_parts = []
    while current:
        current = current.find_next_sibling()
        if not current:
            break
            
        # Stop at next antiphon or Glory
        if current.name == 'span' and current.get('class') == ['rubrica']:
            text = current.get_text()
            if re.search(r'(Ant\.\s*2|Glory\s+to\s+the\s+Father|Psalm\s+Prayer)', text, re.IGNORECASE):
                print(f"STOP at: {text[:50]}")
                break
        
        # Collect non-rubrica text
        if current.name != 'span' or current.get('class') != ['rubrica']:
            text = current.get_text().strip()
            if text:
                psalm_parts.append(text)
                print(f"  Added: {text[:80]}")
    
    print(f"\n\nTotal parts collected: {len(psalm_parts)}")
    full_text = '\n'.join(psalm_parts)
    print(f"Total character count: {len(full_text)}")
    print(f"\nFirst 500 chars:\n{full_text[:500]}")
else:
    print("Ant. 1 not found!")
