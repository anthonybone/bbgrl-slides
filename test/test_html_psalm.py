"""Parser-level HTML extraction tests using local debug fixtures."""

import os
import re
from bs4 import BeautifulSoup
from bbgrl.generator import (
    extract_psalm_verses_from_html,
    extract_canticle_info,
    extract_intercessions_html,
)


def _load_fixture(path: str) -> str:
    with open(path, 'r', encoding='utf-8') as f:
        return f.read()


def test_extract_psalm_1_from_html():
    html = _load_fixture(os.path.join('debug_files', 'debug_morning_prayer.html'))
    soup = BeautifulSoup(html, 'html.parser')
    verses = extract_psalm_verses_from_html(soup, 1)
    assert isinstance(verses, list)
    assert len(verses) >= 2
    assert verses[0]['speaker'] in ('Priest', 'People')
    assert all('text' in v and isinstance(v['text'], str) and len(v['text']) > 0 for v in verses)


def test_extract_canticle_info_from_html():
    html = _load_fixture(os.path.join('debug_files', 'debug_morning_prayer.html'))
    soup = BeautifulSoup(html, 'html.parser')
    text = soup.get_text()
    info = extract_canticle_info(soup, text)
    assert isinstance(info, dict)
    assert 'title' in info and isinstance(info['title'], str)


def test_extract_intercessions_from_html():
    html = _load_fixture(os.path.join('debug_files', 'debug_morning_prayer.html'))
    soup = BeautifulSoup(html, 'html.parser')
    text = soup.get_text()
    groups = extract_intercessions_html(soup, text)
    assert isinstance(groups, list)
    # May be empty depending on fixture, but structure should hold
    for g in groups:
        assert 'intentions' in g and isinstance(g['intentions'], list)
