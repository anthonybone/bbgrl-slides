from __future__ import annotations
import re
from typing import Any, Dict, List, Optional


def _bs4(text_or_html: str):
    try:
        from bs4 import BeautifulSoup  # type: ignore
    except Exception:
        raise
    return BeautifulSoup(text_or_html, 'html.parser')


# ---- Psalm helpers ----

def get_fallback_verses(psalm_number: int) -> List[Dict[str, str]]:
    print(f"  Using fallback verses for Psalm {psalm_number}")
    return [
        {"speaker": "Priest", "text": f"[Psalm {psalm_number} verse 1 - Priest]"},
        {"speaker": "People", "text": f"[Psalm {psalm_number} verse 2 - People]"},
        {"speaker": "Priest", "text": f"[Psalm {psalm_number} verse 3 - Priest]"},
        {"speaker": "People", "text": f"[Psalm {psalm_number} verse 4 - People]"},
    ]


def extract_antiphon_and_psalm_info(text: Any, number: int, text_after_psalmody: Optional[str] = None) -> Dict[str, str]:
    if hasattr(text, 'find_all'):
        soup = text
        text_content = text_after_psalmody if text_after_psalmody else soup.get_text()

        antiphon_text = ""
        psalm_title = ""
        psalm_subtitle = ""

        if number == 1 or number == 3:
            antiphon_patterns = [
                rf'Ant\.\s*{number}[:\s]+(.+?)(?=Psalm\s+\d|\nAnt\.)',
                rf'Antiphon\s*{number}[:\s]+(.+?)(?=Psalm\s+\d|\nAnt\.)',
            ]

            for pattern in antiphon_patterns:
                match = re.search(pattern, text_content, re.IGNORECASE | re.DOTALL)
                if match:
                    antiphon_text = re.sub(r'\s+', ' ', match.group(1).strip()).strip()
                    print(f"  Found Antiphon {number} text: {antiphon_text[:50]}...")
                    break

            try:
                ant_rubrica = soup.find('span', class_='rubrica', string=re.compile(rf'Ant\.\s*{number}'))
                if ant_rubrica:
                    next_rubrica = ant_rubrica.find_next('span', class_='rubrica')
                    if next_rubrica:
                        rubrica_text = next_rubrica.get_text(separator='\n').strip()
                        lines = rubrica_text.split('\n')
                        if lines:
                            psalm_title = lines[0].strip()
                            print(f"  Found red psalm title: {psalm_title}")
                        if len(lines) > 1:
                            psalm_subtitle = lines[1].strip()
                            print(f"  Found red psalm subtitle: {psalm_subtitle}")
            except Exception as e:
                print(f"  WARNING: Could not extract red psalm text from HTML: {e}")

            if not psalm_title:
                psalm_pattern = r'Psalm\s+(\d+)([A-Z])?(?::(\d+)(?:-(\d+))?)?\s*([^\n]*?)(?=\nPsalm|\n\n|Psalm\s+\d|$)'
                psalm_matches = re.finditer(psalm_pattern, text_content if isinstance(text_content, str) else str(text_content), re.IGNORECASE)
                first_psalm_match = None
                for match in psalm_matches:
                    first_psalm_match = match
                    break
                if first_psalm_match:
                    psalm_num = first_psalm_match.group(1)
                    psalm_letter = first_psalm_match.group(2) if first_psalm_match.group(2) else ""
                    verse_start = first_psalm_match.group(3)
                    verse_end = first_psalm_match.group(4)
                    subtitle_raw = first_psalm_match.group(5)
                    if verse_start and verse_end:
                        psalm_title = f"Psalm {psalm_num}{psalm_letter}:{verse_start}-{verse_end}"
                    elif verse_start:
                        psalm_title = f"Psalm {psalm_num}{psalm_letter}:{verse_start}"
                    else:
                        psalm_title = f"Psalm {psalm_num}{psalm_letter}"
                    if subtitle_raw:
                        subtitle = subtitle_raw.strip()
                        subtitle = re.sub(r'Psalm\s+\d.*$', '', subtitle, flags=re.IGNORECASE).strip()
                        subtitle = re.sub(r'\bP?salm\b.*$', '', subtitle, flags=re.IGNORECASE).strip()
                        if len(subtitle) > 100:
                            subtitle = subtitle[:100].rsplit(' ', 1)[0] + '...'
                        psalm_subtitle = subtitle
                    print(f"  Found psalm: {psalm_title} - {psalm_subtitle}")
        else:
            antiphon_patterns = [
                rf'Ant\.\s*{number}[:\s]+([^.]+\.)',
                rf'Antiphon\s*{number}[:\s]+([^.]+\.)',
            ]
            for pattern in antiphon_patterns:
                match = re.search(pattern, text_content, re.IGNORECASE)
                if match:
                    antiphon_text = match.group(1).strip()
                    break

        return {
            "text": antiphon_text if antiphon_text else "",
            "format": "all_response",
            "psalm_title": psalm_title,
            "psalm_subtitle": psalm_subtitle,
        }

    # Fallback plain-text input
    if number == 1:
        antiphon_patterns = [
            r'Ant\.\s+([^.]+\.)',
            rf'Ant\.\s*{number}[:\s]+([^.]+\.)',
            rf'Antiphon\s*{number}[:\s]+([^.]+\.)',
            r'Antiphon[:\s]+([^.]+\.)',
        ]
    else:
        antiphon_patterns = [
            rf'Ant\.\s*{number}[:\s]+([^.]+\.)',
            rf'Antiphon\s*{number}[:\s]+([^.]+\.)',
        ]

    antiphon_text = ""
    for pattern in antiphon_patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            antiphon_text = match.group(1).strip()
            break

    return {
        "text": antiphon_text if antiphon_text else "",
        "format": "all_response",
        "psalm_title": "",
        "psalm_subtitle": "",
    }


def extract_antiphon(text: str, number: int) -> Dict[str, str]:
    patterns = [
        rf'Ant\.\s*{number}[:\s]+([^.!?]+[.!?])',
        rf'Antiphon\s*{number}[:\s]+([^.!?]+[.!?])',
    ]
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            antiphon_text = match.group(1).strip()
            return {"text": antiphon_text, "format": "all_response"}
    return {"text": "", "format": "all_response"}


def extract_psalm_verses_from_html(soup, psalm_number: int) -> List[Dict[str, str]]:
    verses: List[Dict[str, str]] = []
    try:
        ant_pattern = rf'Ant\.\s*{psalm_number}\s*$'
        ant_span = None
        for span in soup.find_all('span', class_='rubrica'):
            if re.match(ant_pattern, span.get_text().strip()):
                ant_span = span
                break
        if not ant_span:
            print(f"  WARNING: Could not find Ant. {psalm_number} in HTML")
            return get_fallback_verses(psalm_number)
        parent = ant_span.parent
        parent_html = str(parent)
        ant_pos = parent_html.find(str(ant_span))
        if ant_pos < 0:
            return get_fallback_verses(psalm_number)
        html_after_ant = parent_html[ant_pos + len(str(ant_span)) :]
        stop_patterns = [
            r'<span class="rubrica">Psalm\s+Prayer</span>',
            rf'<span class="rubrica">Ant\.\s*{psalm_number + 1}</span>',
            r'<span class="rubrica">Ant\.</span>',
        ]
        end_pos = len(html_after_ant)
        for pattern in stop_patterns:
            match = re.search(pattern, html_after_ant, re.IGNORECASE)
            if match:
                end_pos = min(end_pos, match.start())
        psalm_html = html_after_ant[:end_pos]
        soup2 = _bs4(psalm_html)
        for rubrica in soup2.find_all('span', class_='rubrica', string=re.compile(r'Psalm\s+\d')):
            rubrica.decompose()
        for em in soup2.find_all('em'):
            em.decompose()
        remaining_html = str(soup2)
        verse_sections = re.split(r'<br\s*/?>\s*<br\s*/?>', remaining_html)
        verse_count = 0
        skipped_first_section = False
        for section in verse_sections:
            section_soup = _bs4(section)
            for rubrica in section_soup.find_all('span', class_='rubrica'):
                rubrica.decompose()
            verse_text = section_soup.get_text().strip()
            if not verse_text or len(verse_text) < 20:
                continue
            if not skipped_first_section and verse_count == 0:
                if re.search(r'(Each morning|Martin, priest|My heart is ready|You who stand in his sanctuary)', verse_text, re.IGNORECASE):
                    print(f"  Skipping antiphon text in verse extraction: {verse_text[:50]}...")
                    skipped_first_section = True
                    continue
                elif len(verse_text) < 150 and verse_text.endswith(('.', '!', '?')):
                    print(f"  Skipping potential antiphon text in verse extraction: {verse_text[:50]}...")
                    skipped_first_section = True
                    continue
            verse_text = re.sub(r'\s+', ' ', verse_text).strip()
            if not verse_text[-1] in '.!?"':
                verse_text += '.'
            speaker = "Priest" if verse_count % 2 == 0 else "People"
            verses.append({"speaker": speaker, "text": verse_text})
            verse_count += 1
        if verses:
            print(f"  Extracted {len(verses)} verses for Psalm {psalm_number}")
            return verses
    except Exception as e:
        print(f"  WARNING: Error parsing psalm verses from HTML: {e}")
        import traceback
        traceback.print_exc()
    return get_fallback_verses(psalm_number)


def extract_psalm_verses(text: str, psalm_number: int) -> List[Dict[str, str]]:
    verses: List[Dict[str, str]] = []
    try:
        intro_patterns = [
            r'\([^)]{3,50}\)\s*\.',
            r'Psalm\s+\d+[A-Z]?(?::\d+(?:-\d+)?)?\s*[^\n]{10,100}\n',
        ]
        intro_matches = []
        for pattern in intro_patterns:
            intro_matches.extend(list(re.finditer(pattern, text, re.IGNORECASE)))
        intro_matches.sort(key=lambda m: m.start())
        if psalm_number <= len(intro_matches):
            match = intro_matches[psalm_number - 1]
            start_pos = match.end()
        else:
            psalm_pattern = rf'Psalm\s+\d+[A-Z]?(?::\d+(?:-\d+)?)?'
            psalm_matches = list(re.finditer(psalm_pattern, text, re.IGNORECASE))
            if psalm_number <= len(psalm_matches):
                start_pos = psalm_matches[psalm_number - 1].end() + 300
            else:
                return get_fallback_verses(psalm_number)
        glory_match = re.search(r'Glory\s+to\s+the\s+Father', text[start_pos:], re.IGNORECASE)
        if glory_match:
            end_pos = start_pos + glory_match.start()
        else:
            psalm_prayer_match = re.search(r'Psalm\s+Prayer', text[start_pos:], re.IGNORECASE)
            if psalm_prayer_match:
                end_pos = start_pos + psalm_prayer_match.start()
            else:
                end_pos = start_pos + 2000
        verse_section = text[start_pos:end_pos].strip()
        paragraphs = re.split(r'(?:\.\s*\n)|(?:\n\s*\n)', verse_section)
        verse_count = 0
        for para in paragraphs:
            para = para.strip()
            if not para or len(para) < 20:
                continue
            if re.match(r'^(Psalm|Ant\.|Glory|℟)', para, re.IGNORECASE):
                continue
            cleaned = re.sub(r'\s*\*\s*', ' ', para)
            cleaned = re.sub(r'\s+', ' ', cleaned).strip()
            if cleaned and not cleaned[-1] in '.!?"':
                cleaned += '.'
            if len(cleaned) < 20:
                continue
            speaker = "Priest" if verse_count % 2 == 0 else "People"
            verses.append({"speaker": speaker, "text": cleaned})
            verse_count += 1
        if verses:
            print(f"  Extracted {len(verses)} verses for Psalm {psalm_number}")
            return verses
    except Exception as e:
        print(f"  WARNING: Error parsing psalm verses: {e}")
        import traceback
        traceback.print_exc()
    return get_fallback_verses(psalm_number)


# ---- Canticle helpers ----

def get_fallback_canticle_verses() -> Dict[str, Any]:
    print("  Using fallback verses for Canticle")
    return {
        "verses": [
            {"speaker": "Priest", "text": "[Canticle verse 1 - Priest]"},
            {"speaker": "People", "text": "[Canticle verse 2 - People]"},
            {"speaker": "Priest", "text": "[Canticle verse 3 - Priest]"},
            {"speaker": "People", "text": "[Canticle verse 4 - People]"},
        ],
        "omit_glory_be": False,
    }


def extract_canticle_verses(soup, text: Optional[str] = None) -> Dict[str, Any]:
    verses: List[Dict[str, str]] = []
    omit_glory_be = False
    try:
        canticle_span = None
        for span in soup.find_all('span', class_='rubrica'):
            span_text = span.get_text().strip()
            if span_text.startswith('Canticle:') and re.search(r'\d+:\d+', span_text):
                canticle_span = span
                break
        if not canticle_span:
            print("  WARNING: Could not find Canticle marker in HTML")
            return get_fallback_canticle_verses()
        parent = canticle_span.parent
        parent_html = str(parent)
        canticle_pos = parent_html.find(str(canticle_span))
        if canticle_pos < 0:
            return get_fallback_canticle_verses()
        html_after_canticle = parent_html[canticle_pos + len(str(canticle_span)) :]
        if re.search(r'Glory\s+to\s+the\s+Father.*?is\s+not\s+said', html_after_canticle, re.IGNORECASE | re.DOTALL):
            omit_glory_be = True
            print("  ✓ Detected: Glory to the Father is not said for this canticle")
        stop_patterns = [
            r'<span class="rubrica">Glory to the Father</span>',
            r'<span class="rubrica">Ant\.\s*3</span>',
            r'<span class="rubrica">Ant\.</span>',
            r'Glory to the Father',
        ]
        end_pos = len(html_after_canticle)
        for pattern in stop_patterns:
            match = re.search(pattern, html_after_canticle, re.IGNORECASE)
            if match:
                end_pos = min(end_pos, match.start())
        canticle_html = html_after_canticle[:end_pos]
        canticle_soup = _bs4(canticle_html)
        for em in canticle_soup.find_all('em'):
            em.decompose()
        remaining_html = str(canticle_soup)
        verse_sections = re.split(r'<br\s*/?>\s*<br\s*/?>', remaining_html)
        verse_count = 0
        skipped_first_section = False
        for section in verse_sections:
            section_soup = _bs4(section)
            for rubrica in section_soup.find_all('span', class_='rubrica'):
                rubrica.decompose()
            verse_text = section_soup.get_text().strip()
            if not verse_text or len(verse_text) < 20:
                continue
            if not skipped_first_section and verse_count == 0:
                if len(verse_text) < 150 and verse_text.endswith('.'):
                    print(f"  Skipping antiphon text in canticle extraction: {verse_text[:50]}...")
                    skipped_first_section = True
                    continue
            verse_text = re.sub(r'\s+', ' ', verse_text).strip()
            if not verse_text[-1] in '.!?"':
                verse_text += '.'
            speaker = "Priest" if verse_count % 2 == 0 else "People"
            verses.append({"speaker": speaker, "text": verse_text})
            verse_count += 1
        if verses:
            print(f"  Extracted {len(verses)} verses for Canticle")
            return {"verses": verses, "omit_glory_be": omit_glory_be}
    except Exception as e:
        print(f"  WARNING: Error parsing canticle verses from HTML: {e}")
        import traceback
        traceback.print_exc()
    return get_fallback_canticle_verses()


def extract_canticle_info(soup, text: str) -> Dict[str, str]:
    try:
        for span in soup.find_all('span', class_='rubrica'):
            span_text = span.get_text().strip()
            if span_text.startswith('Canticle:') and re.search(r'\d+:\d+', span_text):
                match = re.match(r'(Canticle:\s+[A-Za-z\s]+\d+:\d+(?:[-—]\d+(?::\d+)?)?(?:,\s*\d+)?)(.*)', span_text, re.IGNORECASE)
                if match:
                    title = match.group(1).strip()
                    subtitle = re.sub(r'^[—\-\s]+', '', match.group(2).strip())
                    print(f"  Found Canticle title: {title}")
                    if subtitle:
                        print(f"  Found Canticle subtitle: {subtitle}")
                    return {"title": title, "subtitle": subtitle}
                else:
                    verse_end = re.search(r'\d+([A-Z])', span_text)
                    if verse_end:
                        split_pos = verse_end.start(1)
                        title = span_text[:split_pos].strip()
                        subtitle = span_text[split_pos:].strip()
                        print(f"  Found Canticle title: {title}")
                        if subtitle:
                            print(f"  Found Canticle subtitle: {subtitle}")
                        return {"title": title, "subtitle": subtitle}
                    else:
                        print(f"  Found Canticle (no subtitle split): {span_text}")
                        return {"title": span_text, "subtitle": ""}
    except Exception as e:
        print(f"  WARNING: Error extracting canticle info: {e}")
    return {"title": "[Canticle title]", "subtitle": ""}


# ---- Reading, Responsory, Gospel ----

def extract_short_reading(text: str) -> Dict[str, str]:
    try:
        reading_matches = list(re.finditer(r'READING', text, re.IGNORECASE))
        if not reading_matches:
            print("  WARNING: No READING marker found")
            return {"citation": "", "text": ""}
        reading_start = None
        for match in reading_matches:
            test_start = match.end()
            responsory_test = re.search(r'RESPONSORY', text[test_start:test_start + 1000], re.IGNORECASE)
            if responsory_test:
                reading_start = test_start
                break
        if reading_start is None:
            print("  WARNING: No READING with RESPONSORY found")
            return {"citation": "", "text": ""}
        responsory_match = re.search(r'RESPONSORY', text[reading_start:], re.IGNORECASE)
        if not responsory_match:
            print("  WARNING: No RESPONSORY marker found after READING")
            return {"citation": "", "text": ""}
        reading_end = reading_start + responsory_match.start()
        reading_section = text[reading_start:reading_end].strip()
        reading_section = re.sub(r'^\[.*?\]\s*', '', reading_section)
        citation_match = re.match(r'^([1-3]?\s*[A-Za-z]+\s+\d+:\d+[a-z]?(?:-\d+[a-z]?)?)', reading_section)
        if citation_match:
            citation = citation_match.group(1).strip()
            reading_text = reading_section[citation_match.end():].strip()
        else:
            citation = ""
            reading_text = reading_section
        print(f"  Found READING: {citation}")
        print(f"    Text preview: {reading_text[:100]}...")
        return {"citation": citation, "text": reading_text}
    except Exception as e:
        print(f"  WARNING: Error extracting short reading: {e}")
        return {"citation": "", "text": ""}


def extract_responsory_from_html(soup, text: str) -> List[Dict[str, Any]]:
    try:
        responsory_match = re.search(r'RESPONSORY', text, re.IGNORECASE)
        if not responsory_match:
            print("  WARNING: No RESPONSORY marker found in text")
            return []
        responsory_start = responsory_match.end()
        stop_patterns = [r'\bOr:', r'GOSPEL\s+CANTICLE', r'CANTICLE\s+OF\s+ZECHARIAH']
        responsory_end = len(text)
        for pattern in stop_patterns:
            stop_match = re.search(pattern, text[responsory_start:], re.IGNORECASE)
            if stop_match:
                responsory_end = responsory_start + stop_match.start()
                break
        responsory_section = text[responsory_start:responsory_end].strip()
        normalized_section = responsory_section.replace('\u2014', '—').replace('\u2013', '—').replace('\u2015', '—')
        em_dash_parts = [part.strip() for part in normalized_section.split('—') if part.strip()]
        all_segments: List[str] = []
        if len(em_dash_parts) > 0:
            all_segments.append(em_dash_parts[0])
        if len(em_dash_parts) > 1:
            part1_sentences: List[str] = []
            current = ""
            for i, char in enumerate(em_dash_parts[1]):
                current += char
                if char == '.' and i + 1 < len(em_dash_parts[1]):
                    next_char = em_dash_parts[1][i + 1]
                    if next_char.isupper() or (i + 2 < len(em_dash_parts[1]) and em_dash_parts[1][i + 2].isupper()):
                        part1_sentences.append(current.strip())
                        current = ""
            if current.strip():
                part1_sentences.append(current.strip())
            all_segments.extend(part1_sentences)
        if len(em_dash_parts) > 2:
            glory_pattern = r'(.*?)(Glory\s+to\s+the\s+Father.*)'
            glory_match = re.search(glory_pattern, em_dash_parts[2], re.IGNORECASE | re.DOTALL)
            if glory_match:
                shortened = glory_match.group(1).strip()
                glory = glory_match.group(2).strip()
                if shortened:
                    all_segments.append(shortened)
                if glory:
                    all_segments.append(glory)
            else:
                all_segments.append(em_dash_parts[2])
        for i in range(3, len(em_dash_parts)):
            all_segments.append(em_dash_parts[i])
        if len(all_segments) < 6:
            print(f"  WARNING: Expected 6 segments but found {len(all_segments)}, structure may be incorrect")
            return []
        responsory_verses: List[Dict[str, Any]] = []
        combined_first = all_segments[0].strip()
        if len(all_segments) > 1:
            combined_first += "\n— " + all_segments[1].strip()
        responsory_verses.append({"speaker": "All", "text": combined_first, "include_title": True})
        combined_verse = all_segments[2].strip()
        if len(all_segments) > 3:
            combined_verse += "\n— " + all_segments[3].strip()
        responsory_verses.append({"speaker": "Priest", "text": combined_verse})
        combined_glory = all_segments[4].strip()
        if len(all_segments) > 5:
            combined_glory += "\n— " + all_segments[5].strip()
        responsory_verses.append({"speaker": "Priest", "text": combined_glory})
        print(f"  Found RESPONSORY with {len(responsory_verses)} parts")
        return responsory_verses
    except Exception as e:
        print(f"  WARNING: Error extracting responsory from HTML: {e}")
        import traceback
        traceback.print_exc()
        return []


def extract_responsory(text: str) -> List[Dict[str, str]]:
    try:
        responsory_match = re.search(r'RESPONSORY', text, re.IGNORECASE)
        if not responsory_match:
            print("  WARNING: No RESPONSORY marker found")
            return []
        responsory_start = responsory_match.end()
        stop_patterns = [r'GOSPEL\s+CANTICLE', r'CANTICLE\s+OF\s+ZECHARIAH', r'\bOR\b', r'INTERCESSIONS']
        responsory_end = len(text)
        for pattern in stop_patterns:
            stop_match = re.search(pattern, text[responsory_start:], re.IGNORECASE)
            if stop_match:
                responsory_end = responsory_start + stop_match.start()
                break
        responsory_section = text[responsory_start:responsory_end].strip()
        print(f"  DEBUG: Responsory section length: {len(responsory_section)}")
        print(f"  DEBUG: Responsory section preview: {responsory_section[:300]}")
        normalized_section = responsory_section.replace('\r\n', '\n').replace('\r', '\n')
        all_lines = normalized_section.split('\n')
        lines = [line.strip() for line in all_lines if line.strip() and len(line.strip()) > 3]
        print(f"  DEBUG: Found {len(lines)} non-empty lines")
        for i, line in enumerate(lines[:10]):
            print(f"    Line {i}: {line[:80]}")
        if len(lines) < 6:
            print(f"  WARNING: Responsory has {len(lines)} lines, expected at least 6")
            return []
        responsory_verses = []
        responsory_verses.append({"speaker": "All", "text": lines[0]})
        responsory_verses.append({"speaker": "Priest", "text": lines[2]})
        response_text = lines[3]
        if response_text.startswith('—'):
            response_text = response_text[1:].strip()
        responsory_verses.append({"speaker": "All", "text": response_text})
        responsory_verses.append({"speaker": "Priest", "text": lines[4]})
        final_response = lines[5]
        if final_response.startswith('—'):
            final_response = final_response[1:].strip()
        responsory_verses.append({"speaker": "All", "text": final_response})
        print(f"  Found RESPONSORY with {len(responsory_verses)} parts")
        return responsory_verses
    except Exception as e:
        print(f"  WARNING: Error extracting responsory: {e}")
        import traceback
        traceback.print_exc()
        return []


def extract_gospel_antiphon(text: str) -> str:
    try:
        gc_match = re.search(r'GOSPEL\s+CANTICLE', text, re.IGNORECASE)
        if not gc_match:
            print("  WARNING: No GOSPEL CANTICLE marker found")
            return ""
        start_pos = gc_match.end()
        ant_match = re.search(r'Ant\.', text[start_pos:start_pos + 500], re.IGNORECASE)
        if not ant_match:
            print("  WARNING: No antiphon marker found after GOSPEL CANTICLE")
            return ""
        ant_start = start_pos + ant_match.end()
        stop_patterns = [r'Canticle\s+of\s+Zechariah', r'Benedictus', r'Canticle:', r'INTERCESSIONS', r'Let us pray']
        end_pos = len(text)
        for pattern in stop_patterns:
            stop_match = re.search(pattern, text[ant_start:ant_start + 2000], re.IGNORECASE)
            if stop_match:
                end_pos = ant_start + stop_match.start()
                break
        antiphon_text = text[ant_start:end_pos].strip()
        antiphon_text = re.sub(r'\s+', ' ', antiphon_text).strip()
        antiphon_text = re.sub(r'(Canticle|Benedictus|INTERCESSIONS).*$', '', antiphon_text, flags=re.IGNORECASE).strip()
        if antiphon_text:
            print(f"  Found Gospel Canticle antiphon: {antiphon_text[:80]}...")
            return antiphon_text
    except Exception as e:
        print(f"  WARNING: Error extracting gospel antiphon: {e}")
        import traceback
        traceback.print_exc()
    return ""


def extract_benedictus_verses(text: str) -> List[str]:
    return [
        "Blessed be the Lord, the God of Israel; he has come to his people and set them free.",
        "He has raised up for us a mighty savior, born of the house of his servant David.",
    ]


def extract_intercessions_text(text: str) -> str:
    return "[Intercessions for today]"


# ---- Daily Readings (Mass) ----

def extract_first_reading_citation(text: str) -> str:
    try:
        match = re.search(r'(?:First Reading|FIRST READING)\s*\n\s*([\w\s,:.-]+?)\s*\n', text, re.IGNORECASE)
        if match:
            citation = match.group(1).strip()
            print(f"  Found First Reading citation: {citation}")
            return citation
        return ""
    except Exception as e:
        print(f"  WARNING: Error extracting first reading citation: {e}")
        return ""


def extract_first_reading_verses(text: str) -> List[str]:
    try:
        start_match = re.search(
            r'A reading from (?:the )?(?:'
            r'Book of [^.\n]{5,40}\.?|'
            r'Letter of (?:Saint )?Paul to the [^.\n]{5,40}\.?|'
            r'(?:First|Second|Third) Letter of [^.\n]{5,40}\.?|'
            r'Gospel according to [^.\n]{5,40}\.?|'
            r'(?:Prophet )?[A-Z][a-z]+ [0-9:, -]+\.?'
            r')',
            text, re.IGNORECASE,
        )
        if not start_match:
            start_match = re.search(r'A reading from [^\n]{10,100}?\.', text, re.IGNORECASE)
        if not start_match:
            print(f"  WARNING: Could not find 'A reading from' in text")
            return []
        start_pos = start_match.start()
        end_match = re.search(r'The word of the Lord\.?', text[start_pos:], re.IGNORECASE)
        if not end_match:
            print(f"  WARNING: Could not find 'The word of the Lord' in text")
            return []
        end_pos = start_pos + end_match.end()
        reading_text = text[start_pos:end_pos].strip()
        reading_text = reading_text.replace('\u25a1', '').replace('□', '')
        reading_text = re.sub(r'\s+', ' ', reading_text)
        reading_from_match = re.search(r'A reading from ', reading_text, re.IGNORECASE)
        if reading_from_match:
            start_of_intro = reading_from_match.start()
            remaining = reading_text[start_of_intro:]
            period_match = re.search(r'[.!?]', remaining[15:120])
            if period_match:
                end_of_intro = start_of_intro + 15 + period_match.end()
                reading_intro = reading_text[start_of_intro:end_of_intro].strip()
            else:
                reading_intro = remaining[:80].strip()
                if reading_intro and reading_intro[-1] not in '.!?':
                    reading_intro += '.'
        else:
            reading_intro = start_match.group(0).strip()
            reading_intro = re.sub(r'\s+', ' ', reading_intro)
        if reading_intro and reading_intro[-1] not in '.!?':
            reading_intro += '.'
        content_after_intro = reading_text[len(reading_intro):].strip()
        if content_after_intro and len(reading_intro) > 10:
            match = re.search(r'([a-z])([A-Z][a-z])', reading_intro)
            if match:
                split_pos = match.start(2)
                content_after_intro = reading_intro[split_pos:] + ' ' + content_after_intro
                reading_intro = reading_intro[:split_pos].strip()
                if reading_intro and reading_intro[-1] not in '.!?':
                    reading_intro += '.'
        the_word_ending = 'The word of the Lord.'
        if content_after_intro.endswith(the_word_ending):
            main_content = content_after_intro[:-len(the_word_ending)].strip()
        elif content_after_intro.lower().endswith('the word of the lord'):
            main_content = content_after_intro[:-len('the word of the lord')].strip()
        else:
            word_match = re.search(r'The word of the Lord\.?', content_after_intro, re.IGNORECASE)
            if word_match:
                main_content = content_after_intro[:word_match.start()].strip()
            else:
                main_content = content_after_intro
        main_content = re.sub(r'\.\s+([a-z])', r' \1', main_content)
        main_content = re.sub(r',\s*\.\s+', ', ', main_content)
        lines: List[str] = [reading_intro]
        main_content = re.sub(r'([.!?])([A-Z])', r'\1 \2', main_content)
        main_content = re.sub(r'([,:;])\s+', r'\1 ', main_content)
        main_content = re.sub(r'([,:;])([^\s])', r'\1 \2', main_content)
        sentences = re.split(r'(?<=[.!?])\s+(?=[A-Z][a-z])', main_content)
        for sentence in sentences:
            sentence = sentence.strip()
            if not sentence:
                continue
            if len(sentence) > 80:
                parts = re.split(r'(?<=[,;:])\s+(?=[A-Z])', sentence)
                for part in parts:
                    part = part.strip()
                    if not part:
                        continue
                    if len(part) > 80:
                        subparts = re.split(r'\s+(and|but)\s+', part, flags=re.IGNORECASE)
                        current_line = ''
                        for i, subpart in enumerate(subparts):
                            if i % 2 == 1:
                                current_line += ' ' + subpart + ' '
                            else:
                                if current_line and len(current_line + subpart) > 80:
                                    lines.append(current_line.strip())
                                    current_line = subpart
                                else:
                                    current_line += subpart
                        if current_line.strip():
                            lines.append(current_line.strip())
                    else:
                        lines.append(part)
            else:
                lines.append(sentence)
        lines.append('The word of the Lord.')
        print(f"  Extracted First Reading with {len(lines)} lines")
        return lines
    except Exception as e:
        print(f"  WARNING: Error extracting first reading verses: {e}")
        import traceback
        traceback.print_exc()
        return []


def extract_psalm_citation(html_or_text: str) -> str:
    from bs4 import BeautifulSoup  # type: ignore
    if html_or_text.strip().startswith('<'):
        soup = BeautifulSoup(html_or_text, 'html.parser')
        text = soup.get_text()
    else:
        text = html_or_text
    match = re.search(r'(?:Responsorial Psalm|RESPONSORIAL PSALM)\s*([Pp]s?\s*[\d:,\s-]+)', text, re.IGNORECASE)
    if match:
        citation = match.group(1).strip()
        if not citation.startswith('Ps '):
            citation = 'Ps ' + citation.lstrip('Psp')
        return citation
    return "Ps [citation not found]"


def extract_psalm_response_verses(html_content: str) -> List[str]:
    from bs4 import BeautifulSoup  # type: ignore
    if html_content.strip().startswith('<'):
        soup = BeautifulSoup(html_content, 'html.parser')
        psalm_heading = soup.find(string=re.compile(r'Responsorial Psalm', re.IGNORECASE))
        if not psalm_heading:
            return ["\u211f. [Response not found]", "[Verses not found]"]
        current = psalm_heading.parent
        psalm_paragraphs = []
        for sibling in current.find_all_next(['p', 'hr']):
            if sibling.name == 'hr':
                break
            psalm_paragraphs.append(sibling)
            if sibling.get_text() and re.search(r'Second Reading|Gospel|Acclamation', sibling.get_text(), re.IGNORECASE):
                break
        response = "\u211f. [Response not found]"
        response_short = "[Response not found]"
        for i, p in enumerate(psalm_paragraphs):
            text = p.get_text()
            if text.strip() == "R. :":
                if i + 1 < len(psalm_paragraphs):
                    next_p = psalm_paragraphs[i + 1]
                    html_str = str(next_p).replace('<br>', '\n').replace('<br/>', '\n').replace('<br />', '\n')
                    temp_soup = BeautifulSoup(html_str, 'html.parser')
                    resp_text = temp_soup.get_text()
                    response_match = re.search(r'\u211f\.\s*\(([^)]+)\)\s*([^\n]+?)(?=\s*or:|\s*\n|$)', resp_text)
                    if response_match:
                        ref = response_match.group(1).strip()
                        resp_text_clean = response_match.group(2).strip()
                        response = f"\u211f. ({ref}) {resp_text_clean}"
                        response_short = resp_text_clean
                    else:
                        response_match = re.search(r'\u211f\.\s*([^\n]+?)(?=\s*or:|\s*\n|$)', resp_text)
                        if response_match:
                            resp_text_clean = response_match.group(1).strip()
                            response = f"\u211f. {resp_text_clean}"
                            response_short = resp_text_clean
                break
        result = [response, ""]
        for p in psalm_paragraphs:
            text_content = p.get_text()
            if any(x in text_content for x in ['R. :', 'Ps ', 'Responsorial Psalm', 'Second Reading']):
                continue
            html_str = str(p).replace('<br>', '\n').replace('<br/>', '\n').replace('<br />', '\n')
            temp_soup = BeautifulSoup(html_str, 'html.parser')
            verse_text = temp_soup.get_text()
            verse_text = re.sub(r'\u211f\.\s*[^\n]*(?:\n\s*or:\s*\n\s*\u211f\.\s*[^\n]*)?', '\n\n', verse_text)
            stanzas = re.split(r'\n\s*\n+', verse_text)
            for stanza in stanzas:
                stanza = stanza.strip()
                if not stanza or len(stanza) < 15:
                    continue
                if stanza.lower().startswith('or:') or stanza.lower() == 'alleluia.' or stanza.lower() == 'alleluia':
                    continue
                stanza = re.sub(r'\s+Alleluia\.\s*$', '', stanza, flags=re.IGNORECASE)
                lines = [line.strip() for line in stanza.split('\n') if line.strip()]
                if lines:
                    formatted_verse = '\n   '.join(lines)
                    result.append(formatted_verse)
                    result.append("")
                    result.append(f"\u211f. {response_short}")
                    result.append("")
        while result and result[-1] == "":
            result.pop()
        return result if len(result) > 2 else [response, "[Verses not found]"]
    else:
        text = html_content
        return ["\u211f. [Response not found]", "[Verses - HTML parsing required]"]


def extract_gospel_acclamation(html_content: str) -> Dict[str, str]:
    try:
        from bs4 import BeautifulSoup  # type: ignore
        soup = BeautifulSoup(html_content, 'html.parser')
        title_span = soup.find('span', class_='titolo', string=re.compile(r'Acclamation before the Gospel', re.IGNORECASE))
        if not title_span:
            print("  WARNING: Could not find Acclamation section")
            return {"citation": "", "verse": ""}
        title_p = title_span.find_parent('p')
        citation_span = title_p.find('span', class_='citazione')
        citation = citation_span.get_text().strip() if citation_span else ""
        verse_p = title_p.find_next_sibling('p')
        if not verse_p:
            print("  WARNING: Could not find verse paragraph after acclamation title")
            return {"citation": citation, "verse": ""}
        verse_html = str(verse_p)
        verse_html = re.sub(r'<span class="rubrica">℟\.</span>\s*<strong>Alleluia, alleluia\.</strong>', '', verse_html, flags=re.IGNORECASE)
        verse_soup = BeautifulSoup(verse_html, 'html.parser')
        for br in verse_soup.find_all('br'):
            br.replace_with('\n')
        verse_text = verse_soup.get_text().strip()
        verse_text = re.sub(r'\n\s*\n+', '\n', verse_text)
        print(f"  Extracted Acclamation: {citation}")
        print(f"    Verse preview: {verse_text[:60]}...")
        return {"citation": citation, "verse": verse_text}
    except Exception as e:
        print(f"  WARNING: Error extracting gospel acclamation: {e}")
        import traceback
        traceback.print_exc()
        return {"citation": "", "verse": ""}


def extract_gospel_citation(html_content: str) -> str:
    try:
        from bs4 import BeautifulSoup  # type: ignore
        soup = BeautifulSoup(html_content, 'html.parser')
        title_span = soup.find('span', class_='titolo', string=re.compile(r'^Gospel$', re.IGNORECASE))
        if not title_span:
            print("  WARNING: Could not find Gospel section")
            return ""
        title_p = title_span.find_parent('p')
        citation_span = title_p.find('span', class_='citazione')
        citation = citation_span.get_text().strip() if citation_span else ""
        return citation
    except Exception as e:
        print(f"  WARNING: Error extracting gospel citation: {e}")
        return ""


def extract_gospel_verses(html_content: str) -> Dict[str, str]:
    try:
        from bs4 import BeautifulSoup  # type: ignore
        soup = BeautifulSoup(html_content, 'html.parser')
        title_span = soup.find('span', class_='titolo', string=re.compile(r'^Gospel$', re.IGNORECASE))
        if not title_span:
            print("  WARNING: Could not find Gospel section")
            return {"intro_text": "", "proclamation": "", "text": "", "closing": "", "response": ""}
        title_p = title_span.find_parent('p')
        intro_p = title_p.find_next_sibling('p')
        intro_text = intro_p.get_text().strip() if intro_p else ""
        gospel_p = intro_p.find_next_sibling('p') if intro_p else None
        if not gospel_p:
            print("  WARNING: Could not find Gospel text paragraph")
            return {"intro_text": intro_text, "proclamation": "", "text": "", "closing": "", "response": ""}
        proclamation = ""
        for strong in gospel_p.find_all('strong'):
            t = strong.get_text().strip()
            if 'reading from the holy Gospel' in t:
                proclamation = "✠ " + t
                break
        gospel_html = str(gospel_p)
        gospel_html = re.sub(r'.*?</strong><br><br>', '', gospel_html, count=1, flags=re.DOTALL)
        gospel_end_match = re.search(r'<strong>The Gospel of the Lord\.</strong>', gospel_html, re.IGNORECASE)
        if gospel_end_match:
            gospel_text_html = gospel_html[:gospel_end_match.start()]
        else:
            gospel_text_html = gospel_html
        gospel_soup = BeautifulSoup(gospel_text_html, 'html.parser')
        for rubrica in gospel_soup.find_all('span', class_='rubrica'):
            rubrica.decompose()
        for br in gospel_soup.find_all('br'):
            br.replace_with('\n')
        gospel_text = gospel_soup.get_text()
        gospel_text = re.sub(r'\xa0+', '   ', gospel_text)
        gospel_text = re.sub(r' +', ' ', gospel_text)
        gospel_text = re.sub(r'\n +', '\n   ', gospel_text)
        gospel_text = gospel_text.strip()
        gospel_text = re.sub(r'<[^>]+>', '', gospel_text)
        gospel_text = re.sub(r'At the end of the Gospel[^\n]*\n?', '', gospel_text, flags=re.IGNORECASE)
        gospel_text = re.sub(r'Then he kisses[^\n]*\n?', '', gospel_text, flags=re.IGNORECASE)
        gospel_text = re.sub(r'Through the words of the Gospel[^\n]*\n?', '', gospel_text, flags=re.IGNORECASE)
        if gospel_text.startswith('✠'):
            first_newline = gospel_text.find('\n')
            if first_newline > 0:
                gospel_text = gospel_text[first_newline + 1 :].strip()
        elif gospel_text.startswith('A reading from the holy Gospel'):
            first_newline = gospel_text.find('\n')
            if first_newline > 0:
                gospel_text = gospel_text[first_newline + 1 :].strip()
        print(f"  Extracted Gospel with {len(gospel_text)} characters")
        print(f"    Intro: {intro_text[:60]}...")
        return {
            "intro_text": intro_text,
            "proclamation": proclamation,
            "text": gospel_text,
            "closing": "The Gospel of the Lord.",
            "response": "Praise to you, Lord Jesus Christ.",
        }
    except Exception as e:
        print(f"  WARNING: Error extracting gospel verses: {e}")
        import traceback
        traceback.print_exc()
        return {"intro_text": "", "proclamation": "", "text": "", "closing": "", "response": ""}


# ---- Intercessions ----

def extract_intercessions_html(soup, text: str) -> List[Dict[str, Any]]:
    try:
        html_content = str(soup)
        intercessions_matches = list(re.finditer(r'INTERCESSIONS', html_content, re.IGNORECASE))
        if not intercessions_matches:
            print("  WARNING: No INTERCESSIONS marker found")
            return []
        intercessions_pos = intercessions_matches[-1].start()
        html_from_intercessions = html_content[intercessions_pos:]
        end_match = re.search(r"THE LORD.S PRAYER|Let us pray\.", html_from_intercessions, re.IGNORECASE)
        if end_match:
            intercessions_section = html_from_intercessions[: end_match.start()]
        else:
            intercessions_section = html_from_intercessions[:3000]
        print(f"  Found INTERCESSIONS section ({len(intercessions_section)} chars)")
        intercessions_groups: List[Dict[str, Any]] = []
        category_pattern = r'\[(Martyrs|Pastors|Doctors|Virgins|Holy Men and Women)\]'
        parts = re.split(category_pattern, intercessions_section, flags=re.IGNORECASE)
        if len(parts) == 1:
            intercessions_groups.append({'category': None, 'text': parts[0]})
        else:
            for i in range(1, len(parts), 2):
                if i + 1 < len(parts):
                    category = parts[i]
                    text_content = parts[i + 1]
                    intercessions_groups.append({'category': category, 'text': text_content})
        all_intercessions: List[Dict[str, Any]] = []
        for group in intercessions_groups:
            group_text = group['text']
            category = group['category']
            intro_match = re.search(r'(.*?)(?=Nourish your people|You redeemed us)', group_text, re.DOTALL | re.IGNORECASE)
            if intro_match:
                introduction = intro_match.group(1).strip()
                introduction = re.sub(r'INTERCESSIONS', '', introduction, flags=re.IGNORECASE).strip()
                introduction = re.sub(r'<[^>]+>', '', introduction).strip()
                intentions_text = group_text[intro_match.end():]
            else:
                introduction = ""
                intentions_text = group_text
            response_match = re.search(r'(Nourish your people, Lord\.|You redeemed us by your blood\.)', group_text, re.IGNORECASE)
            response_line = response_match.group(1) if response_match else ""
            cleaned_text = re.sub(r'<em>\s*(Nourish your people, Lord\.|You redeemed us by your blood\.)\s*</em>', '', intentions_text, flags=re.IGNORECASE)
            cleaned_text = re.sub(r'(Nourish your people, Lord\.|You redeemed us by your blood\.)', '', cleaned_text, flags=re.IGNORECASE)
            cleaned_text = re.sub(r'<span class="rubrica">—</span>', '—', cleaned_text, flags=re.IGNORECASE)
            cleaned_text = re.sub(r'<br\s*/?>', ' ', cleaned_text, flags=re.IGNORECASE)
            cleaned_text = re.sub(r'<[^>]+>', ' ', cleaned_text, flags=re.IGNORECASE)
            intention_pattern = r'([^—<]+?)—\s*([^<]+?)\.'
            intentions: List[Dict[str, str]] = []
            for m in re.finditer(intention_pattern, cleaned_text):
                petition = m.group(1).strip()
                response = m.group(2).strip()
                if len(petition) < 20:
                    continue
                if re.search(r'Nourish your people|You redeemed us|INTERCESSIONS', petition, re.IGNORECASE):
                    continue
                petition = re.sub(r'<[^>]+>', '', petition).strip()
                petition = re.sub(r'^\s*[,\s]+', '', petition).strip()
                response = re.sub(r'<[^>]+>', '', response).strip()
                response = re.sub(r'\s+', ' ', response).strip()
                if response and not response.endswith('.'):
                    response += '.'
                if petition and response:
                    intentions.append({'petition': petition, 'response': response})
            if introduction or intentions:
                all_intercessions.append({
                    'category': category,
                    'introduction': introduction,
                    'response_line': response_line,
                    'intentions': intentions,
                })
        print(f"  Extracted {len(all_intercessions)} intercession group(s)")
        for i, group in enumerate(all_intercessions):
            print(f"    Group {i+1}: {group['category'] or 'Default'}, {len(group['intentions'])} intentions")
        return all_intercessions
    except Exception as e:
        print(f"  WARNING: Error extracting intercessions: {e}")
        import traceback
        traceback.print_exc()
        return []


# ---- Concluding Prayer ----

def extract_concluding_prayer(text: str) -> str:
    try:
        prayer_match = re.search(r'CONCLUDING\s+PRAYER', text, re.IGNORECASE)
        if not prayer_match:
            print("  WARNING: No CONCLUDING PRAYER marker found")
            return ""
        prayer_start = prayer_match.end()
        stop_patterns = [r'\bOr:', r'SACRED\s+HEART', r'MASS\s+READINGS', r'FIRST\s+READING']
        prayer_end = len(text)
        for pattern in stop_patterns:
            stop_match = re.search(pattern, text[prayer_start:], re.IGNORECASE)
            if stop_match:
                prayer_end = prayer_start + stop_match.start()
                break
        prayer_section = text[prayer_start:prayer_end].strip()
        amen_match = re.search(r'—\s*Amen\.?', prayer_section, re.IGNORECASE)
        if amen_match:
            prayer_section = prayer_section[:amen_match.end()].strip()
        prayer_section = re.sub(r'\n\s*\n+', '\n', prayer_section)
        prayer_section = re.sub(r'[ \t]+', ' ', prayer_section)
        prayer_section = re.sub(r'\n ', '\n', prayer_section)
        print(f"  Found CONCLUDING PRAYER: {prayer_section[:50]}...")
        return prayer_section
    except Exception as e:
        print(f"  WARNING: Error extracting concluding prayer: {e}")
        import traceback
        traceback.print_exc()
        return ""
