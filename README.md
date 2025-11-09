# Catholic Mass Slides Generator

This program automatically generates PowerPoint slides for daily Catholic Mass with morning prayers and daily readings, designed with large text for elderly congregation members to read from the back of the church.

## Features

- **Large, readable text**: Optimized for elderly congregation members
- **Color-coded content**: 
  - Blue text for audience and priest together (Antiphon 1)
  - Red text for priest-only responses
- **Automatic content fetching**: Gets current daily content from iBreviary.com
- **Professional formatting**: Widescreen slides with proper spacing

## Requirements

- Python 3.7+
- Required packages (install with `pip install -r requirements.txt`):
  - requests
  - beautifulsoup4
  - python-pptx
  - lxml

## Usage

### Basic Usage
```bash
python enhanced_slide_generator.py
```

This will create `enhanced_daily_mass_slides.pptx` with:
1. Title slide with current date
2. Antiphon 1 (Blue text - audience and priest together)
3. Priest Response (Red text - priest only)
4. First Reading
5. Gospel reading

### Current Slide Order (as requested)

1. **Antiphon 1** (Blue text)
   - The audience and priest read together
   - Example: "Come, worship the Lord, for we are his people, the flock he shepherds, alleluia."

2. **Priest Response** (Red text)
   - The priest reads the first paragraph after Antiphon 1
   - Usually a psalm verse or prayer response

3. **Additional steps** will be added as requested

## Data Source

The program fetches content from:
- **Morning Prayers**: https://www.ibreviary.com/m2/breviario.php?s=lodi
- **Daily Readings**: https://www.ibreviary.com/m2/letture.php?s=letture

The content is automatically retrieved for the current day, excluding Sundays as requested.

## File Structure

```
├── enhanced_slide_generator.py    # Main program (recommended)
├── catholic_slide_generator.py    # Basic version
├── test_extraction.py            # Testing/debugging script
├── requirements.txt              # Python dependencies
└── README.md                     # This file
```

## Customization

You can modify:
- Text sizes in the `_add_*_slide` methods
- Colors by changing RGB values
- Slide layout and positioning
- Content extraction patterns

## Notes

- The program works with weekday content (excludes Sundays as requested)
- Text is optimized for elderly readability with large fonts
- Slides are in widescreen format (13.33" x 7.5")
- Content is automatically cleaned and formatted for display

## Troubleshooting

If you encounter issues:
1. Check your internet connection (program needs to fetch live content)
2. Ensure all dependencies are installed
3. Run `test_extraction.py` to see what content is being fetched
4. The website structure may occasionally change - content extraction patterns may need updates

## Future Enhancements

Planned features:
- Additional steps in the morning prayer sequence
- Multiple days/week generation
- Custom formatting options
- Offline content caching
