"""
Test script to generate slides for multiple consecutive days
Verifies that Selenium-based date navigation retrieves unique content for each day
"""

from bbgrl_slide_generator_v2_template import BBGRLSlideGeneratorV2
from datetime import datetime, timedelta

def main():
    print("=" * 70)
    print("Multi-Day Slide Generation Test")
    print("=" * 70)
    print("This will generate slides for 3 consecutive days using Selenium navigation")
    print("to verify that iBreviary returns unique content for each date.\n")
    
    generator = BBGRLSlideGeneratorV2()
    
    # Generate slides for 3 consecutive days
    start_date = datetime(2025, 11, 10)
    
    for i in range(3):
        target_date = start_date + timedelta(days=i)
        
        print(f"\n{'=' * 70}")
        print(f"Generating slides for: {target_date.strftime('%B %d, %Y')} (Day {i+1}/3)")
        print(f"{'=' * 70}\n")
        
        try:
            # Fetch liturgical data for the specified date
            liturgical_data = generator.fetch_live_liturgical_data(target_date)
            
            # Display the antiphon for verification
            antiphon_1 = liturgical_data['morning_prayer']['psalmody']['antiphon_1']
            print(f"\nAntiphon 1: {antiphon_1['text'][:80]}...")
            
            # Create presentation
            output_path = generator.create_presentation_from_template(liturgical_data)
            
            print(f"\nSuccessfully created: {output_path}")
            
        except Exception as e:
            print(f"\nError generating slides for {target_date.strftime('%B %d, %Y')}: {e}")
            import traceback
            traceback.print_exc()
    
    print(f"\n{'=' * 70}")
    print("Multi-day generation test complete!")
    print(f"{'=' * 70}")
    print("\nCheck the output_v2 folder to verify that each presentation has unique content.")

if __name__ == "__main__":
    main()
