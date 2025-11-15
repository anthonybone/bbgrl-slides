from datetime import datetime
from .static_content import get_static_devotional_content


def get_fallback_morning_prayer():
    """Fallback morning prayer structure if iBreviary fails."""
    return {
        "psalmody": {
            "antiphon_1": {
                "text": "",
                "format": "all_response",
                "psalm_title": "",
                "psalm_subtitle": "",
            },
            "psalm_1": [{"speaker": "Priest", "text": ""}],
            "antiphon_2": {"text": "", "format": "all_response"},
            "canticle_info": {"title": "", "subtitle": ""},
            "canticle": {
                "verses": [{"speaker": "Priest", "text": ""}],
                "omit_glory_be": False,
            },
            "antiphon_3": {
                "text": "",
                "format": "all_response",
                "psalm_title": "",
                "psalm_subtitle": "",
            },
            "psalm_3": [{"speaker": "Priest", "text": ""}],
        },
        "reading": {
            "short_reading": {"citation": "", "text": ""},
            "responsory": [],
        },
        "gospel_canticle": {
            "antiphon": "",
            "benedictus_verses": [],
        },
        "intercessions": [],
        "concluding_prayer": "",
    }


def get_fallback_readings():
    """Fallback readings if iBreviary fails."""
    return {
        "first_reading": {"citation": "[Citation]", "verses": ["[Reading text]"]},
        "responsorial_psalm": {
            "citation": "[Psalm citation]",
            "verses": ["[Psalm text]"],
        },
        "gospel_acclamation": {"verse": "[Alleluia verse]"},
        "gospel": {"citation": "[Gospel citation]", "verses": ["[Gospel text]"]},
    }


def get_fallback_data(target_date=None):
    """Complete fallback data structure."""
    if target_date is None:
        target_date = datetime.now()

    return {
        "date": target_date.strftime("%B %d, %Y"),
        "morning_prayer": get_fallback_morning_prayer(),
        "mass_readings": get_fallback_readings(),
        "static_content": get_static_devotional_content(),
    }
