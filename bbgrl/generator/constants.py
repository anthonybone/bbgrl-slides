from pptx.dml.color import RGBColor
from pptx.util import Pt


def get_reference_template():
    """Build and return the reference template metadata and formatting rules.

    Extracted from the original monolith to make the structure easy to find
    and edit without scrolling through implementation details.
    """
    return {
        "metadata": {
            "total_expected_slides": 135,
            "title_pattern": "{date} Morning Readings & Prayers",
            "structure_sections": [
                "opening_slides",
                "psalmody_section",
                "reading_section",
                "gospel_canticle_section",
                "intercessions_section",
                "sacred_heart_hymns",
                "mass_readings",
                "post_communion_prayers",
                "transition_slides",
                "jubilee_prayer",
                "st_joseph_prayer",
            ],
        },
        "section_templates": {
            "opening_slides": {
                "slide_count": 2,
                "slides": [
                    {"type": "title", "content": "title_slide"},
                    {"type": "blank", "content": "transition"},
                ],
            },
            "psalmody_section": {
                "expected_elements": [
                    "antiphon_1",
                    "psalm_1",
                    "glory_be",
                    "repeat_antiphon_1",
                    "antiphon_2",
                    "canticle",
                    "glory_be",
                    "repeat_antiphon_2",
                    "antiphon_3",
                    "psalm_2",
                    "glory_be",
                    "repeat_antiphon_3",
                ],
                "slide_pattern": "alternating_priest_people",
                "title_slide": {"text": "PSALMODY", "include": True},
            },
            "reading_section": {
                "expected_elements": ["short_reading", "responsory"],
                "slide_pattern": "reading_format",
                "title_slide": {"text": "READING", "include": True},
            },
            "gospel_canticle_section": {
                "expected_elements": [
                    "gospel_antiphon",
                    "benedictus",
                    "glory_be",
                    "repeat_antiphon",
                ],
                "slide_pattern": "canticle_format",
                "title_slide": {"text": "GOSPEL CANTICLE", "include": True},
            },
            "intercessions_section": {
                "expected_elements": [
                    "intercessions",
                    "lords_prayer",
                    "concluding_prayer",
                ],
                "slide_pattern": "intercession_format",
                "title_slide": {"text": "INTERCESSIONS", "include": True},
            },
            "sacred_heart_hymns": {"slide_count": 6, "content_type": "static_devotional"},
            "mass_readings": {
                "expected_elements": [
                    "first_reading",
                    "responsorial_psalm",
                    "gospel_acclamation",
                    "gospel",
                ],
                "slide_pattern": "mass_reading_format",
            },
            "post_communion_prayers": {"slide_count": 17, "content_type": "static_devotional"},
            "transition_slides": {"slide_count": 10, "content_type": "blank_transitions"},
            "jubilee_prayer": {"slide_count": 7, "content_type": "static_prayer"},
            "st_joseph_prayer": {"slide_count": 12, "content_type": "static_prayer"},
        },
        "formatting_rules": {
            # Named colors
            "priest_color": RGBColor(200, 0, 0),
            "people_color": RGBColor(0, 100, 200),
            "all_color": RGBColor(100, 0, 100),
            "title_color": RGBColor(0, 51, 102),
            "reading_color": RGBColor(100, 0, 0),
            "devotional_color": RGBColor(139, 0, 0),
            # Font sizes
            "font_sizes": {
                "title": Pt(48),
                "subtitle": Pt(32),
                "priest_text": Pt(32),
                "people_text": Pt(32),
                "all_text": Pt(36),
                "reading_text": Pt(30),
                "prayer_text": Pt(30),
            },
        },
    }
