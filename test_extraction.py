import requests
from bs4 import BeautifulSoup
import re

def test_content_extraction():
    """Test what content we can extract from iBreviary"""
    
    base_url = "https://www.ibreviary.com/m2/"
    session = requests.Session()
    session.headers.update({
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
    })

    print("Testing Morning Prayer extraction...")
    print("=" * 50)
    
    # Test morning prayer
    try:
        response = session.get(f"{base_url}breviario.php?s=lodi")
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # Get just a portion of the text to see structure
        full_text = soup.get_text()
        print("Sample of morning prayer content:")
        print("-" * 30)
        
        # Look for antiphon patterns
        antiphon_matches = re.finditer(r'Ant\.\s+([^.]+\.)', full_text)
        print("Found antiphons:")
        for i, match in enumerate(list(antiphon_matches)[:3]):
            print(f"  {i+1}. {match.group(1).strip()}")
        
        # Look for invitatory
        invitatory_match = re.search(r'(Come, worship the Lord[^.]*alleluia\.)', full_text, re.IGNORECASE)
        if invitatory_match:
            print(f"\nInvitatory found: {invitatory_match.group(1)}")
        
        # Look for psalm verses
        psalm_pattern = r'([A-Z][^.]*Lord[^.]*\.)'
        psalm_matches = re.finditer(psalm_pattern, full_text)
        print("\nSample psalm verses:")
        for i, match in enumerate(list(psalm_matches)[:3]):
            verse = match.group(1).strip()
            if len(verse) > 20 and len(verse) < 150:
                print(f"  {i+1}. {verse}")
                
    except Exception as e:
        print(f"Error fetching morning prayer: {e}")
    
    print("\n" + "=" * 50)
    print("Testing Daily Readings extraction...")
    print("=" * 50)
    
    # Test readings
    try:
        response = session.get(f"{base_url}letture.php?s=letture")
        soup = BeautifulSoup(response.text, 'html.parser')
        
        full_text = soup.get_text()
        
        # Look for first reading
        first_reading_match = re.search(r'First Reading(.{100,800}?)(?=Responsorial|Second Reading|Gospel)', full_text, re.IGNORECASE | re.DOTALL)
        if first_reading_match:
            print("First Reading found:")
            content = first_reading_match.group(1).strip()
            content = re.sub(r'\s+', ' ', content)
            print(f"  {content[:200]}...")
        
        # Look for gospel
        gospel_match = re.search(r'Gospel(.{100,800})', full_text, re.IGNORECASE | re.DOTALL)
        if gospel_match:
            print("\nGospel found:")
            content = gospel_match.group(1).strip()
            content = re.sub(r'\s+', ' ', content)
            print(f"  {content[:200]}...")
            
    except Exception as e:
        print(f"Error fetching readings: {e}")

if __name__ == "__main__":
    test_content_extraction()