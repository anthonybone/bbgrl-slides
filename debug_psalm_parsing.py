"""
Debug script to test psalm parsing
"""
from bbgrl_slide_generator_v2_template import BBGRLSlideGeneratorV2
from datetime import datetime

# Create generator instance
generator = BBGRLSlideGeneratorV2()

# Fetch data for a specific date
target_date = datetime(2025, 11, 12)
print(f"Fetching data for {target_date.strftime('%B %d, %Y')}...")

# Navigate to the page
html_content = generator._navigate_ibreviary_to_date(target_date)

if html_content:
    # Write the HTML to a file for inspection
    with open('debug_morning_prayer.html', 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    # Extract text
    from bs4 import BeautifulSoup
    soup = BeautifulSoup(html_content, 'html.parser')
    full_text = soup.get_text()
    
    # Find PSALMODY position
    psalmody_pos = full_text.upper().find('PSALMODY')
    if psalmody_pos >= 0:
        text_after_psalmody = full_text[psalmody_pos:]
        
        # Write the text after PSALMODY to a file
        with open('debug_text_after_psalmody.txt', 'w', encoding='utf-8') as f:
            f.write(text_after_psalmody[:3000])  # First 3000 characters
        
        print("Files written:")
        print("  - debug_morning_prayer.html (full HTML)")
        print("  - debug_text_after_psalmody.txt (text after PSALMODY)")
        
        # Test psalm extraction
        verses = generator._extract_psalm_verses(text_after_psalmody, 1)
        print(f"\nExtracted {len(verses)} verses:")
        for i, verse in enumerate(verses):
            print(f"\n{i+1}. {verse['speaker']}:")
            print(f"   {verse['text'][:100]}...")
else:
    print("Failed to fetch HTML")

# Clean up
if generator.driver:
    generator.driver.quit()
