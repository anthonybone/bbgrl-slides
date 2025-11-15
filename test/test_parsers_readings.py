"""Tests for readings/acclamation/gospel parser functions using debug fixtures."""

import os
from bbgrl.generator import (
    extract_psalm_citation,
    extract_psalm_response_verses,
    extract_gospel_acclamation,
    extract_gospel_citation,
    extract_gospel_verses,
)


def _load_fixture(path: str) -> str:
    with open(path, 'r', encoding='utf-8') as f:
        return f.read()


def test_extract_responsorial_psalm_from_readings_html():
    html = _load_fixture(os.path.join('debug_files', 'readings_nov10_raw.html'))
    citation = extract_psalm_citation(html)
    verses = extract_psalm_response_verses(html)
    assert isinstance(citation, str) and len(citation) > 0
    assert isinstance(verses, list) and len(verses) >= 2


def test_extract_gospel_sections_from_readings_html():
    html = _load_fixture(os.path.join('debug_files', 'readings_nov10_raw.html'))
    acclamation = extract_gospel_acclamation(html)
    citation = extract_gospel_citation(html)
    gospel = extract_gospel_verses(html)
    assert isinstance(acclamation, dict)
    assert 'verse' in acclamation
    assert isinstance(citation, str)
    assert isinstance(gospel, dict)
    assert 'text' in gospel and isinstance(gospel['text'], str)
