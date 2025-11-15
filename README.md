# Catholic Mass Slides Generator

This program automatically generates PowerPoint slides for daily Catholic Mass with morning prayers and daily readings, designed with large text for elderly congregation members to read from the back of the church.

## Desktop/Web UI (new)

You can now run a simple desktop-friendly web UI with a date picker, a 0–100% progress bar showing the current step, and a final Download button for the generated PowerPoint.

### Run the UI locally

```bash
pip install -r requirements.txt
python ui_app/app.py
```

Then open http://127.0.0.1:5000 in your browser. Select the date and click Generate. When it finishes, click Download PowerPoint.

### Windows Installer (recommended)

Download a ready-to-run Windows installer from GitHub Releases once the workflow completes.

For maintainers: a GitHub Actions workflow builds the installer automatically. Trigger it via a tag (e.g. `v1.0.0`) or manually:

```bash
# Local manual build (optional alternative)
package_windows.bat
```

What the installer does:
- Installs the app under Program Files
- Creates a Start Menu shortcut (and optional desktop shortcut)
- Launches the app after install

Troubleshooting on Windows:
- If the browser doesn’t open automatically after launch, open the URL printed by the app (it picks a free port starting at 5000).
- The app writes a log file next to the installed EXE: `ui_app.log`.


Notes:
- Chrome/Chromium must be available for Selenium.
- First run may take longer while ChromeDriver is initialized.

### Chrome OS

Use the Linux (Crostini) environment:

```bash
sudo apt-get update
sudo apt-get install -y python3-pip
pip3 install -r requirements.txt
python3 ui_app/app.py
```

Open Chrome to http://127.0.0.1:5000. The Download button saves the PowerPoint to your Linux home; you can move it to your Chrome OS files.

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
  - Flask

## Usage

### Command-line Generation (new)
- Generate a slides deck for a specific date and save it into `new_slides/` with the required filename format `olph_slides_[MM]_[DD]_[YYYY].pptx`.

```bash
# Syntax (MM-DD-YYYY)
python bbgrl_slide_generator_v1.py 11-15-2025

# Output
# new_slides/olph_slides_11_15_2025.pptx
```

- Notes:
  - Date argument is required in the format `MM-DD-YYYY` (strict).
  - The script fetches Morning Prayer and Daily Readings live from iBreviary and assembles the full deck.
  - Ensure dependencies are installed (`pip install -r requirements.txt`) and Chrome is available for Selenium.

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
