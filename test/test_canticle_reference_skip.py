import re
from bs4 import BeautifulSoup
from bbgrl.generator.parsers import extract_canticle_verses


def test_canticle_skips_parenthetical_reference():
    # Minimal HTML simulating a canticle with an initial parenthetical citation
    html = '''
    <div>
      <p>
        <span class="rubrica">Canticle: Isaiah 38:10-20</span>
        (Revelation 1:17-18)<br><br>
        Once I said, “In the noontime of life I must depart! To the gates of the nether world I shall be consigned for the rest of my years.”<br><br>
        I said, I shall see the Lord no more in the land of the living.
      </p>
    </div>
    '''
    soup = BeautifulSoup(html, 'html.parser')
    result = extract_canticle_verses(soup)
    verses = result.get('verses', [])
    assert verses, "Expected verses to be extracted from canticle"
    first = verses[0]['text']
    assert first.startswith("Once I said"), (
        f"First canticle verse should start with Hezekiah's line, got: {first}"
    )
    # Ensure the parenthetical reference is not included anywhere as a verse
    assert not any(re.match(r"^\(Revelation\s+1:17-18\)\.?$", v['text']) for v in verses), (
        "Parenthetical citation should not be included as a verse"
    )
