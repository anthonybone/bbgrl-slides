from bbgrl_slide_generator_v2_template import BBGRLSlideGeneratorV2
from datetime import datetime

gen = BBGRLSlideGeneratorV2()

# Test November 11
print("=" * 60)
print("NOVEMBER 11, 2025")
print("=" * 60)
target_date = datetime(2025, 11, 11)
html = gen._navigate_ibreviary_to_date(target_date)

if html:
    from bs4 import BeautifulSoup
    soup = BeautifulSoup(html, 'html.parser')
    
    print("\nAll rubrica spans containing 'Canticle':")
    for i, span in enumerate(soup.find_all('span', class_='rubrica')):
        span_text = span.get_text().strip()
        if 'Canticle' in span_text or 'canticle' in span_text.lower():
            print(f"  Span {i}: '{span_text}'")

print("\n" + "=" * 60)
print("NOVEMBER 12, 2025")
print("=" * 60)
target_date = datetime(2025, 11, 12)
html = gen._navigate_ibreviary_to_date(target_date)

if html:
    from bs4 import BeautifulSoup
    soup = BeautifulSoup(html, 'html.parser')
    
    print("\nAll rubrica spans containing 'Canticle':")
    for i, span in enumerate(soup.find_all('span', class_='rubrica')):
        span_text = span.get_text().strip()
        if 'Canticle' in span_text or 'canticle' in span_text.lower():
            print(f"  Span {i}: '{span_text}'")
