import requests
from bs4 import BeautifulSoup
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import re
from datetime import datetime, timedelta
import sys
import os

class EnhancedCatholicSlideGenerator:
    def __init__(self):
        self.base_url = "https://www.ibreviary.com/m2/"
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        })

    def get_liturgy_of_hours_structure(self, is_first_hour=True):
        """
        Get the proper Liturgy of the Hours structure for Morning Prayer
        """
        liturgy_structure = {
            "date": datetime.now().strftime('%B %d, %Y'),
            "is_first_hour": is_first_hour,
            "invitatory": None,
            "morning_prayer": None
        }
        
        # Invitatory (only if it's the first hour of the day)
        if is_first_hour:
            liturgy_structure["invitatory"] = {
                "sequence": [
                    "Opening Verse: Lord, open my lips.",
                    "Response: And my mouth will proclaim your praise.",
                    "Invitatory Antiphon",
                    "Invitatory Psalm (95, 100, 67, or 24)",
                    "Repeat Invitatory Antiphon"
                ]
            }
        
        # Morning Prayer structure
        liturgy_structure["morning_prayer"] = {
            "sequence": [
                "Opening Verse: God, come to my assistance.",
                "Response: Lord, make haste to help me.",
                "Glory to the Father, and to the Son, and to the Holy Spirit.",
                "As it was in the beginning, is now, and will be forever. Amen.",
                "Optional: Alleluia (omit during Lent)",
                "Hymn",
                "Antiphon 1",
                "Psalm 1",
                "Glory to the Father",
                "Repeat Antiphon 1",
                "Antiphon 2", 
                "Psalm 2",
                "Glory to the Father",
                "Repeat Antiphon 2",
                "Antiphon 3",
                "Old Testament Canticle",
                "Glory to the Father",
                "Repeat Antiphon 3",
                "Short Reading (Scripture)",
                "Responsory",
                "Gospel Canticle Antiphon",
                "Benedictus (Luke 1:68–79)",
                "Glory to the Father",
                "Repeat Gospel Canticle Antiphon",
                "Intercessions",
                "The Lord's Prayer",
                "Concluding Prayer (Collect)",
                "Blessing or Dismissal",
                "Optional: Marian Antiphon"
            ]
        }
        
        return liturgy_structure

    def fetch_morning_prayer_detailed(self):
        """
        Fetch and parse morning prayer with detailed extraction from iBreviary
        """
        url = f"{self.base_url}breviario.php?s=lodi"
        
        try:
            response = self.session.get(url)
            response.raise_for_status()
            soup = BeautifulSoup(response.text, 'html.parser')
            
            full_text = soup.get_text()
            
            # Get the liturgy structure
            liturgy = self.get_liturgy_of_hours_structure(is_first_hour=True)
            
            # Try to extract actual content from iBreviary for specific elements
            liturgy['extracted_content'] = {}
            
            # Find the main invitatory antiphon
            invitatory_pattern = r'(Come, worship the Lord[^.]*alleluia\.)'
            invitatory_match = re.search(invitatory_pattern, full_text, re.IGNORECASE)
            if invitatory_match:
                liturgy['extracted_content']['invitatory_antiphon'] = invitatory_match.group(1).strip()
            
            # Find psalm content
            psalm_patterns = [
                r'Come, let us sing to the Lord[^.]*\.',
                r'Let us approach him with praise[^.]*\.',
                r'The Lord is God, the mighty God[^.]*\.',
            ]
            
            for pattern in psalm_patterns:
                match = re.search(pattern, full_text, re.IGNORECASE)
                if match:
                    liturgy['extracted_content']['psalm_verse'] = match.group(0).strip()
                    break
            
            return liturgy
            
        except Exception as e:
            print(f"Error fetching morning prayer: {e}")
            # Return basic structure even if fetching fails
            return self.get_liturgy_of_hours_structure(is_first_hour=True)

    def fetch_daily_readings_detailed(self):
        """
        Fetch daily readings with better parsing
        """
        url = f"{self.base_url}letture.php?s=letture"
        
        try:
            response = self.session.get(url)
            response.raise_for_status()
            soup = BeautifulSoup(response.text, 'html.parser')
            
            full_text = soup.get_text()
            
            readings = {}
            
            # Extract first reading
            first_reading_match = re.search(r'First Reading(.{50,1500}?)(?=Responsorial Psalm|Second Reading|Gospel)', full_text, re.IGNORECASE | re.DOTALL)
            if first_reading_match:
                content = first_reading_match.group(1).strip()
                content = re.sub(r'\s+', ' ', content)
                # Clean up and extract the actual reading text
                reading_match = re.search(r'A reading from[^.]*\.(.+)', content, re.DOTALL)
                if reading_match:
                    readings['first_reading'] = reading_match.group(1).strip()[:1000]
                else:
                    readings['first_reading'] = content[:800]
            
            # Extract Gospel
            gospel_match = re.search(r'Gospel(.{50,1500})$', full_text, re.IGNORECASE | re.DOTALL)
            if gospel_match:
                content = gospel_match.group(1).strip()
                content = re.sub(r'\s+', ' ', content)
                # Clean up
                gospel_reading_match = re.search(r'A reading from[^.]*\.(.+)', content, re.DOTALL)
                if gospel_reading_match:
                    readings['gospel'] = gospel_reading_match.group(1).strip()[:1000]
                else:
                    readings['gospel'] = content[:800]
            
            return readings
            
        except Exception as e:
            print(f"Error fetching readings: {e}")
            return {}

    def create_enhanced_slides(self, prayer_data, readings_data, output_filename=None):
        """
        Create PowerPoint slides with enhanced formatting for elderly congregation
        """
        # Generate filename with date if not provided
        if output_filename is None:
            date_str = datetime.now().strftime("%Y-%m-%d")
            output_filename = f"enhanced_daily_mass_slides_{date_str}.pptx"
        prs = Presentation()
        
        # Set slide dimensions for widescreen
        prs.slide_width = Inches(13.33)
        prs.slide_height = Inches(7.5)

        # 1. Title slide
        self._add_enhanced_title_slide(prs, prayer_data.get('date', ''))
        
        # 2. Liturgy of the Hours - Invitatory (if first hour)
        if prayer_data and prayer_data.get('is_first_hour') and prayer_data.get('invitatory'):
            self._add_liturgy_sequence_slides(prs, "Invitatory", prayer_data['invitatory']['sequence'])
        
        # 3. Liturgy of the Hours - Morning Prayer
        if prayer_data and prayer_data.get('morning_prayer'):
            self._add_liturgy_sequence_slides(prs, "Morning Prayer", prayer_data['morning_prayer']['sequence'])
        
        # 4. First Reading slide
        if readings_data and readings_data.get('first_reading'):
            self._add_reading_slide(prs, "First Reading", readings_data['first_reading'])
        
        # 5. Gospel slide
        if readings_data and readings_data.get('gospel'):
            self._add_reading_slide(prs, "Gospel", readings_data['gospel'])
        
        # Ensure output directory exists
        output_dir = "output"
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        # Create full path for output file
        output_path = os.path.join(output_dir, output_filename)
        
        # Remove existing file if it exists (replacing old file with new for same date)
        if os.path.exists(output_path):
            os.remove(output_path)
            print(f"Replaced existing file: {output_path}")
        
        # Save the presentation
        prs.save(output_path)
        print(f"Enhanced slides saved as: {output_path}")

    def _add_liturgy_sequence_slides(self, prs, section_title, sequence):
        """
        Add slides for Liturgy of the Hours sequences (Invitatory and Morning Prayer)
        """
        for i, item in enumerate(sequence):
            slide_layout = prs.slide_layouts[6]  # Blank layout
            slide = prs.slides.add_slide(slide_layout)
            
            # Section title at top
            title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(12.33), Inches(1.2))
            title_frame = title_box.text_frame
            title_frame.text = f"{section_title} - Step {i + 1}"
            title_para = title_frame.paragraphs[0]
            title_para.font.size = Pt(36)
            title_para.font.bold = True
            title_para.font.color.rgb = RGBColor(0, 51, 102)  # Dark blue
            title_para.alignment = PP_ALIGN.CENTER
            
            # Main content
            content_box = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(11.33), Inches(4.5))
            content_frame = content_box.text_frame
            content_frame.word_wrap = True
            content_frame.margin_bottom = Inches(0.2)
            content_frame.margin_top = Inches(0.2)
            content_frame.margin_left = Inches(0.2)
            content_frame.margin_right = Inches(0.2)
            
            # Determine text color and style based on content type
            text_color = self._get_liturgy_text_color(item)
            
            content_frame.text = item
            content_para = content_frame.paragraphs[0]
            content_para.font.size = Pt(44)
            content_para.font.bold = True
            content_para.font.color.rgb = text_color
            content_para.alignment = PP_ALIGN.CENTER
            content_para.line_spacing = 1.2
    
    def _get_liturgy_text_color(self, item):
        """
        Determine text color based on liturgy item type
        """
        item_lower = item.lower()
        
        # Responses and congregation parts - Blue
        if any(keyword in item_lower for keyword in ['response:', 'glory to the father', 'amen', 'alleluia', 'repeat']):
            return RGBColor(0, 100, 200)  # Blue for congregation responses
        
        # Priest-only parts - Red
        elif any(keyword in item_lower for keyword in ['opening verse:', 'blessing', 'dismissal', 'concluding prayer']):
            return RGBColor(200, 0, 0)  # Red for priest only
        
        # Psalms and readings - Purple
        elif any(keyword in item_lower for keyword in ['psalm', 'canticle', 'reading', 'benedictus', 'magnificat']):
            return RGBColor(128, 0, 128)  # Purple for psalms/readings
        
        # Default - Dark blue
        else:
            return RGBColor(0, 51, 102)  # Dark blue for general liturgical text

    def _add_enhanced_title_slide(self, prs, date_str):
        """Enhanced title slide with large, readable text"""
        slide_layout = prs.slide_layouts[6]  # Blank slide layout
        slide = prs.slides.add_slide(slide_layout)
        
        # Main title with proper text boundaries - positioned higher for better balance, full width
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(12.5), Inches(2.5))
        title_frame = title_box.text_frame
        title_frame.text = "Daily Catholic Mass"
        title_frame.word_wrap = True
        title_frame.margin_left = Inches(0.3)
        title_frame.margin_right = Inches(0.3)
        title_paragraph = title_frame.paragraphs[0]
        title_paragraph.font.size = Pt(60)
        title_paragraph.font.bold = True
        title_paragraph.alignment = PP_ALIGN.CENTER
        title_paragraph.font.color.rgb = RGBColor(0, 0, 100)  # Dark blue
        
        # Subtitle with proper boundaries - more space, full width
        subtitle_box = slide.shapes.add_textbox(Inches(0.5), Inches(4.2), Inches(12.5), Inches(2))
        subtitle_frame = subtitle_box.text_frame
        subtitle_frame.text = f"Liturgy of the Hours - Morning Prayer and Readings"
        subtitle_frame.word_wrap = True
        subtitle_frame.margin_left = Inches(0.3)
        subtitle_frame.margin_right = Inches(0.3)
        subtitle_paragraph = subtitle_frame.paragraphs[0]
        subtitle_paragraph.font.size = Pt(40)
        subtitle_paragraph.alignment = PP_ALIGN.CENTER
        
        # Date with proper boundaries - lower position with more space, full width
        date_box = slide.shapes.add_textbox(Inches(0.5), Inches(6.5), Inches(12.5), Inches(1.5))
        date_frame = date_box.text_frame
        date_frame.text = date_str
        date_frame.word_wrap = True
        date_frame.margin_left = Inches(0.3)
        date_frame.margin_right = Inches(0.3)
        date_paragraph = date_frame.paragraphs[0]
        date_paragraph.font.size = Pt(32)
        date_paragraph.alignment = PP_ALIGN.CENTER
        date_paragraph.font.italic = True

    def _add_reading_slide(self, prs, reading_title, reading_text):
        """Add reading slides with large, readable text"""
        # Split long text across multiple slides if necessary
        max_chars_per_slide = 1000  # Increased from 800 since we have more space
        text_chunks = self._split_text_for_slides(reading_text, max_chars_per_slide)
        
        for i, chunk in enumerate(text_chunks):
            slide_layout = prs.slide_layouts[6]  # Blank slide
            slide = prs.slides.add_slide(slide_layout)
            
            # Title (add part indicator if multiple slides) - full width
            title_text = reading_title if len(text_chunks) == 1 else f"{reading_title} (Part {i+1}/{len(text_chunks)})"
            title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(12.5), Inches(1))
            title_frame = title_box.text_frame
            title_frame.text = title_text
            title_frame.word_wrap = True
            title_paragraph = title_frame.paragraphs[0]
            title_paragraph.font.size = Pt(48)
            title_paragraph.font.bold = True
            title_paragraph.alignment = PP_ALIGN.CENTER
            title_paragraph.font.color.rgb = RGBColor(0, 0, 100)
            
            # Reading text with proper boundaries - extended to use full slide width
            content_box = slide.shapes.add_textbox(Inches(0.4), Inches(1.8), Inches(12.8), Inches(6.2))
            content_frame = content_box.text_frame
            content_frame.word_wrap = True
            content_frame.margin_left = Inches(0.4)
            content_frame.margin_right = Inches(0.4)
            content_frame.margin_top = Inches(0.2)
            content_frame.margin_bottom = Inches(0.2)
            
            content_frame.text = chunk
            
            # Set font size based on text length - larger sizes for extended space
            char_count = len(chunk)
            if char_count <= 400:
                font_size = 32
            elif char_count <= 600:
                font_size = 30
            elif char_count <= 800:
                font_size = 28
            elif char_count <= 1000:
                font_size = 26
            else:
                font_size = 24
            
            for paragraph in content_frame.paragraphs:
                paragraph.font.size = Pt(font_size)
                paragraph.alignment = PP_ALIGN.LEFT
                paragraph.space_after = Pt(12)  # Add space between paragraphs
    
    def _split_text_for_slides(self, text, max_chars_per_slide):
        """Split text into chunks suitable for individual slides"""
        if len(text) <= max_chars_per_slide:
            return [text]
        
        chunks = []
        current_chunk = ""
        sentences = text.split('. ')
        
        for sentence in sentences:
            # If adding this sentence would exceed the limit, start a new chunk
            if len(current_chunk + sentence) > max_chars_per_slide and current_chunk:
                chunks.append(current_chunk.strip())
                current_chunk = sentence + '. '
            else:
                current_chunk += sentence + '. '
        
        # Add the last chunk if it has content
        if current_chunk.strip():
            chunks.append(current_chunk.strip())
        
        # If we still have chunks that are too long, split them more aggressively
        final_chunks = []
        for chunk in chunks:
            if len(chunk) <= max_chars_per_slide:
                final_chunks.append(chunk)
            else:
                # Split by words if sentences are still too long
                words = chunk.split(' ')
                current_word_chunk = ""
                for word in words:
                    if len(current_word_chunk + word) > max_chars_per_slide and current_word_chunk:
                        final_chunks.append(current_word_chunk.strip())
                        current_word_chunk = word + ' '
                    else:
                        current_word_chunk += word + ' '
                if current_word_chunk.strip():
                    final_chunks.append(current_word_chunk.strip())
        
        return final_chunks if final_chunks else [text[:max_chars_per_slide]]

def main(include_invitatory=True):
    print("Enhanced Catholic Slide Generator")
    print("=" * 40)
    
    generator = EnhancedCatholicSlideGenerator()
    
    print("Fetching morning prayer content...")
    prayer_data = generator.fetch_morning_prayer_detailed()
    
    # Override the is_first_hour setting if specified
    if prayer_data and not include_invitatory:
        prayer_data['is_first_hour'] = False
    
    if prayer_data:
        print(f"✓ Found Liturgy structure for: {prayer_data['date']}")
        if prayer_data.get('is_first_hour'):
            print("✓ Including Invitatory (first hour of day)")
        print("✓ Including Morning Prayer sequence")
        
        # Show extracted content if available
        if prayer_data.get('extracted_content'):
            extracted = prayer_data['extracted_content']
            if extracted.get('invitatory_antiphon'):
                print(f"✓ Found Invitatory Antiphon: {extracted['invitatory_antiphon'][:50]}...")
            if extracted.get('psalm_verse'):
                print(f"✓ Found Psalm verse: {extracted['psalm_verse'][:50]}...")
    
    print("\nFetching daily readings...")
    readings_data = generator.fetch_daily_readings_detailed()
    
    if readings_data:
        if 'first_reading' in readings_data:
            print(f"✓ Found First Reading: {readings_data['first_reading'][:50]}...")
        if 'gospel' in readings_data:
            print(f"✓ Found Gospel: {readings_data['gospel'][:50]}...")
    
    print("\nCreating enhanced slides...")
    generator.create_enhanced_slides(prayer_data, readings_data)
    print("✓ Enhanced slides created successfully!")

if __name__ == "__main__":
    # Check for command line arguments
    include_invitatory = True
    if len(sys.argv) > 1 and sys.argv[1].lower() in ['--no-invitatory', '-n']:
        include_invitatory = False
        print("Note: Running without Invitatory (not first hour)")
    
    main(include_invitatory)