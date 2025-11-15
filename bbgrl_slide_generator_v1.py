"""
Legacy entry wrapper for the BBGRL slide generator.

This file re-exports the orchestrator class from `bbgrl.generator.generator`
to maintain backwards compatibility for imports like:

    from bbgrl_slide_generator_v1 import bbgrlslidegeneratorv1
"""

from datetime import datetime
from bbgrl.generator.generator import bbgrlslidegeneratorv1


def main():
    import sys

    print("BBGRL Slide Generator V1 - Template-Based Dynamic Generator")
    print("=" * 60)
    print("Fetching live liturgical data and applying reference structure...")

    generator = bbgrlslidegeneratorv1()

    # Accept optional date arg in format MM-DD-YYYY; default to a sample date
    if len(sys.argv) > 1:
        date_str = sys.argv[1]
        try:
            target_date = datetime.strptime(date_str, "%m-%d-%Y")
        except ValueError:
            print(f"Error: Invalid date format '{date_str}'. Use MM-DD-YYYY")
            return
    else:
        target_date = datetime(2025, 11, 11)

    print(f"Generating slides for: {target_date.strftime('%B %d, %Y')}")
    liturgical_data = generator.fetch_live_liturgical_data(target_date)

    # Build requested filename and directory: olph_slides_[MM]_[DD]_[YYYY].pptx in new_slides/
    out_name = f"olph_slides_{target_date.month:02d}_{target_date.day:02d}_{target_date.year}.pptx"
    generator.create_presentation_from_template(
        liturgical_data,
        output_filename=out_name,
        output_dir="new_slides",
    )
    print("\n✓ Template-based presentation created successfully!")
    print("✓ Uses live liturgical data with exact reference formatting")
    print("✓ File naming: olph_slides_[MM]_[DD]_[YYYY].pptx in new_slides/")


if __name__ == "__main__":
    main()
