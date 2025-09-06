import json
import io
import math
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_VERTICAL_ANCHOR, MSO_AUTO_SIZE
import numpy as np
from datetime import datetime

# --- Dark Mode Branding Constants ---
text_font = 'Arial'
key_font = 'Arial Bold'
heading_font = 'Arial Bold'

# --- Slide Dimension Constants ---
SLIDE_WIDTH = Inches(16)  # Actual slide width in the presentation
SLIDE_HEIGHT = Inches(9)  # Actual slide height in the presentation
SLIDE_MARGIN = Inches(0.5)  # Safe margin from edges
TITLE_HEIGHT = Inches(0.8)  # Height reserved for title
CONTENT_TOP = Inches(1.2)  # Top position for content (after title)
CONTENT_MAX_WIDTH = SLIDE_WIDTH - (2 * SLIDE_MARGIN)  # Max content width
CONTENT_MAX_HEIGHT = SLIDE_HEIGHT - CONTENT_TOP - SLIDE_MARGIN  # Max content height

SLIDE_BACKGROUND_COLOR = RGBColor(15, 22, 50)
DEFAULT_TEXT_COLOR = RGBColor(0xFF, 0xFF, 0xFF)
TABLE_HEADER_BG_COLOR = RGBColor(0x44, 0x54, 0x6A)
TABLE_HEADER_FONT_COLOR = RGBColor(0xFF, 0xFF, 0xFF)
ROW_COLOR_DARK = RGBColor(0x2A, 0x39, 0x50)
BRAND_PIE_COLORS = ['#007ACC', '#09534F', '#000000']
HYPERLINK_COLOR = RGBColor(0x9B, 0xC1, 0xE4)

# --- Branding Helper Functions ---
def set_title_style(title_shape, presentation_width):
    title_shape.left = Inches(0)
    title_shape.width = SLIDE_WIDTH  # Use constant instead of parameter
    title_shape.height = TITLE_HEIGHT
    title_shape.fill.background() # Make title bar transparent
    line = title_shape.line
    line.color.rgb = TABLE_HEADER_BG_COLOR
    line.width = Pt(1)
    font = title_shape.text_frame.paragraphs[0].font
    font.name = heading_font
    font.size = Pt(28)
    font.color.rgb = DEFAULT_TEXT_COLOR
    title_shape.text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT
    tf = title_shape.text_frame
    tf.margin_left = Inches(0.2)
    tf.margin_right = Inches(0.2)
    tf.margin_top = Inches(0.1)
    tf.margin_bottom = Inches(0.1)
    tf.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE

print ('libraries and branding assets loaded successfully.' )

# --- Slide Guardrail Functions ---
def remove_unused_placeholders(slide):
    """Remove unused placeholder text boxes from slide to eliminate 'Click to add text' boxes"""
    shapes_to_remove = []
    
    for shape in slide.shapes:
        try:
            # Check if it's a placeholder
            if hasattr(shape, 'is_placeholder') and shape.is_placeholder:
                # Check if it has placeholder format
                if hasattr(shape, 'placeholder_format'):
                    placeholder_type = shape.placeholder_format.type
                    # Only remove content placeholders (not title=0 or subtitle=1)
                    # Remove types: 7=content, 13=content, 14=content, 15=content, etc.
                    if placeholder_type in (7, 13, 14, 15, 16, 17, 18, 19):  # Content placeholder types only
                        # Check if the placeholder is empty or has default text
                        if hasattr(shape, 'text_frame') and shape.text_frame:
                            text_content = shape.text_frame.text.strip()
                            if not text_content or text_content == "Click to add text":
                                shapes_to_remove.append(shape)
                        else:
                            shapes_to_remove.append(shape)
            # Also check for non-placeholder shapes that contain "Click to add text"
            elif hasattr(shape, 'text_frame') and shape.text_frame:
                text_content = shape.text_frame.text.strip()
                if text_content == "Click to add text":
                    # Make sure it's not a title shape we want to keep
                    if not (shape == slide.shapes.title or 
                           (hasattr(shape, 'name') and ('Title' in str(shape.name)))):
                        shapes_to_remove.append(shape)
        except Exception:
            # Skip shapes that cause errors
            continue
    
    # Remove the identified placeholders
    for shape in shapes_to_remove:
        try:
            slide.shapes._spTree.remove(shape._element)
        except Exception:
            # Skip if removal fails
            continue

def ensure_content_fits(left, top, width, height):
    """Ensures content positioning stays within slide boundaries"""
    # Ensure content doesn't go beyond slide boundaries
    max_left = SLIDE_WIDTH - SLIDE_MARGIN
    max_top = SLIDE_HEIGHT - SLIDE_MARGIN
    
    # Adjust left position
    if left < SLIDE_MARGIN:
        left = SLIDE_MARGIN
    elif left > max_left:
        left = max_left
    
    # Adjust top position
    if top < CONTENT_TOP:
        top = CONTENT_TOP
    elif top > max_top:
        top = max_top
    
    # Adjust width to fit within boundaries
    available_width = SLIDE_WIDTH - left - SLIDE_MARGIN
    if width > available_width:
        width = available_width
    
    # Adjust height to fit within boundaries
    available_height = SLIDE_HEIGHT - top - SLIDE_MARGIN
    if height > available_height:
        height = available_height
    
    return left, top, width, height

def calculate_safe_table_dimensions(num_rows, num_cols):
    """Calculate safe table dimensions that fit within slide boundaries"""
    # Reserve space for title
    available_height = CONTENT_MAX_HEIGHT
    available_width = CONTENT_MAX_WIDTH
    
    # Calculate minimum row height (including header)
    min_row_height = Inches(0.4)
    total_height_needed = min_row_height * (num_rows + 1)  # +1 for header
    
    # If table is too tall, reduce row height
    if total_height_needed > available_height:
        actual_row_height = available_height / (num_rows + 1)
        if actual_row_height < Inches(0.3):  # Minimum readable height
            actual_row_height = Inches(0.3)
    else:
        actual_row_height = min_row_height
    
    # Calculate actual table height
    actual_height = actual_row_height * (num_rows + 1)
    if actual_height > available_height:
        actual_height = available_height
    
    return SLIDE_MARGIN, CONTENT_TOP, available_width, actual_height

def truncate_text_if_needed(text, max_length=500):
    """Truncate text if it's too long to prevent overflow"""
    if len(text) > max_length:
        return text[:max_length-3] + "..."
    return text

def calculate_chart_size():
    """Calculate safe chart size that fits within slide boundaries"""
    # Leave some space for text around charts
    chart_width = CONTENT_MAX_WIDTH * 0.45  # Use 45% of available width
    chart_height = CONTENT_MAX_HEIGHT * 0.8  # Use 80% of available height
    return chart_width, chart_height

def set_header_row_height(self, table, height):
    """Sets a fixed height for the header row (the first row of a table)."""
    header_row = table.rows[0]
    header_row.height = height

def apply_uniform_row_heights_and_autofit(self, table, height):
    for row in list(table.rows)[1:]:
        row.height = height
        for cell in row.cells:
            cell.text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
            cell.text_frame.word_wrap = True

def set_cell_style(self, cell, text, is_header_false, is_dark_row=False):
    cell.text = text
    
    # Enable auto-sizing and word wrapping for better text fitting
    cell.text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    cell.text_frame.word_wrap = True
    
    p = cell.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    p.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
    font = p.font
    font.name = text_font
    font.size = Pt(10)  # Slightly smaller default font
    cell.margin_left = Inches(0.05)  # Smaller margins for more content space
    cell.margin_right = Inches(0.05)
    if is_header_false:
        font.name = key_font
        font.size = Pt(11)  # Slightly smaller header font
        font.color.rgb = TABLE_HEADER_FONT_COLOR
        cell.fill.solid()
        cell.fill.fore_color.rgb = TABLE_HEADER_BG_COLOR
    else:
        font.color.rgb = DEFAULT_TEXT_COLOR
        fill = cell.fill
        fill.solid()
        if is_dark_row:
            fill.fore_color.rgb = ROW_COLOR_DARK
        else:
            fill.background()

print("libraries and branding assets loaded successfully.")

# # Cell 3: Load Data from Your JSON File (Run this in a new notebook cell)
json_filename = 'esg_analysis(4).json'
try:
    with open(json_filename, 'r') as f:
        esg_analysis_data = json.load(f)
    print(f"Successfully loaded data from '{json_filename}'.")
except FileNotFoundError:
    print(f"ERROR: '{json_filename}' not found. Please ensure it is in the same directory.")
    esg_analysis_data = {}
except json.JSONDecodeError as e:
    print(f"ERROR: Could not parse '{json_filename}'. Error: {e}")
    esg_analysis_data = {}

# # Cell 4: Presentation Generator class (Run this in a new notebook cell)
class ESG_Presentation:
    def __init__(self, data):
        if not data:
            raise ValueError("Input data is empty.")
        self.data = data
        self.prs = Presentation()
        self.prs.slide_width = Inches(16)
        self.prs.slide_height = Inches(9)
        self.MAX_ROWS_PER_TABLE = 10
        for layout in self.prs.slide_layouts:
            layout.background.fill.solid()
            layout.background.fill.fore_color.rgb = SLIDE_BACKGROUND_COLOR
    
    def _set_header_row_height(self, table, height):
        """Sets a fixed height for the header row (the first row of a table)."""
        header_row = table.rows[0]
        header_row.height = height

    def _apply_uniform_row_heights_and_autofit(self, table, height):
        for row in list(table.rows)[1:]:
            row.height = height
            for cell in row.cells:
                cell.text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
                cell.text_frame.word_wrap = True

    def _set_cell_style(self, cell, text, is_header=False, is_dark_row=False):
        cell.text = text
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        p.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
        font = p.font
        font.name = text_font
        font.size = Pt(11)
        cell.margin_left = Inches(0.1)
        cell.margin_right = Inches(0.1)
        if is_header:
            font.name = key_font
            font.size = Pt(12)
            font.color.rgb = TABLE_HEADER_FONT_COLOR
            cell.fill.solid()
            cell.fill.fore_color.rgb = TABLE_HEADER_BG_COLOR
        else:
            font.color.rgb = DEFAULT_TEXT_COLOR
            fill = cell.fill
            fill.solid()
            if is_dark_row:
                fill.fore_color.rgb = ROW_COLOR_DARK
            else:
                fill.background()

    
    
    
    
    
    def add_title_slide(self, title, subtitle):
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[0])
        slide.shapes.title.text = title
        slide.placeholders[1].text = subtitle
        slide.shapes.title.text_frame.paragraphs[0].font.name = heading_font
        subtitle_font = slide.placeholders[1].text_frame.paragraphs[0].font
        subtitle_font.name = text_font
        subtitle_font.color.rgb = DEFAULT_TEXT_COLOR
        
        # Remove unused placeholders to eliminate "Click to add text" boxes
        remove_unused_placeholders(slide)

    def _create_country_sentiment_barchart(self, country_data):
        """Generates a stacked bar chart of sentiments by country with controlled legend."""
        labels = list(country_data.keys())
        positive_counts = [d['Positive'] for d in country_data.values()]
        neutral_counts = [d['Neutral'] for d in country_data.values()]
        negative_counts = [d['Negative'] for d in country_data.values()]

        x = np.arange(len(labels))
        width = 0.5
        fig, ax = plt.subplots(figsize=(10, 3))
        
        ax.bar(x, positive_counts, width, label='Positive', color=BRAND_PIE_COLORS[0])
        ax.bar(x, neutral_counts, width, bottom=positive_counts, label='Neutral', color=BRAND_PIE_COLORS[2])
        bottom_for_negative = np.add(positive_counts, neutral_counts).tolist()
        ax.bar(x, negative_counts, width, bottom=np.add(positive_counts, neutral_counts), label='Negative', color='#09534F')

        # --- CHANGE 1: updated y-axis label ---
        ax.set_ylabel('Number of news articles', color='white', fontname=text_font)

        ax.set_title('Sentiment Distribution by Country', color='white', fontname=key_font)
        ax.set_xticks(x)
        ax.set_xticklabels(labels, rotation=45, ha='right', fontname=text_font)
        ax.tick_params(axis='x', colors='white')
        ax.tick_params(axis='y', colors='white')
        fig.patch.set_alpha(0.0)
        ax.patch.set_alpha(0.0)

        # --- CHANGE 2: Controlled legend position ---
        # Places the legend horizontally above the chart area
        legend = ax.legend(loc='upper center', bbox_to_anchor=(0.5, 1.05),
                  ncol=3, frameon=False)
        plt.setp(legend.get_texts(), color='white', fontname=text_font)

        plt.tight_layout()
        chart_buffer = io.BytesIO()
        plt.savefig(chart_buffer, format='png', bbox_inches='tight', transparent=True)
        plt.close(fig)
        chart_buffer.seek(0)
        return chart_buffer
        
    def _prepare_country_sentiment_data(self):
        """Processes regionalData to get the SUM of article counts per sentiment per country."""
        country_sentiments = {}
        for region, countries in self.data['regionalData'].items():
            for country, categories in countries.items():
                sentiments = {'Positive': 0, 'Neutral': 0, 'Negative': 0}
                for category_name, themes in categories.items():
                    for theme in themes:
                        if isinstance(theme, dict):  # Ensure theme is a dictionary
                            sentiment = theme.get('sentiment')
                            # use .get("articleCount") as a fallback in case the field is missing
                            article_count = theme.get('articleCount', 0)
                            if sentiment in sentiments:
                                sentiments[sentiment] += article_count
                country_sentiments[country] = sentiments
        return country_sentiments

    def _create_pie_chart(self):
        labels = self.data['sentimentSummary'].keys()
        sizes = self.data['sentimentSummary'].values()
        fig, ax = plt.subplots()
        wedges, texts, autotexts = ax.pie(
            sizes, labels=labels, autopct='%.1f%%', startangle=90,
            wedgeprops=dict(width=0.4), colors=BRAND_PIE_COLORS,
            pctdistance=0.8, labeldistance=1.1)
        plt.setp(autotexts, size=10, weight="bold", fontname=key_font, color="white")
        plt.setp(texts, size=12, fontname=text_font, color="white")
        ax.axis('equal')
        plt.tight_layout()
        chart_buffer = io.BytesIO()
        plt.savefig(chart_buffer, format='png', bbox_inches='tight', transparent=True)
        plt.close(fig)
        chart_buffer.seek(0)
        return chart_buffer
    
    def add_slide1_summary(self):
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[1])
        slide.shapes.title.text = 'Executive Summary'
        set_title_style(slide.shapes.title, self.prs.slide_width)
        
        # Apply guardrails to text box positioning and sizing
        text_left, text_top, text_width, text_height = ensure_content_fits(
            SLIDE_MARGIN, CONTENT_TOP, CONTENT_MAX_WIDTH * 0.5, Inches(2.5)
        )
        txBox = slide.shapes.add_textbox(text_left, text_top, text_width, text_height)
        tf = txBox.text_frame
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        p = tf.paragraphs[0]
        
        # Truncate text if needed to prevent overflow
        summary_text = self.data.get('executiveSummary', 'No executive summary available')
        p.text = truncate_text_if_needed(summary_text, 800)
        p.font.name = text_font
        p.font.size = Pt(16)  # Slightly smaller font to fit better
        p.font.color.rgb = DEFAULT_TEXT_COLOR
        
        # Calculate safe chart dimensions
        chart_width, chart_height = calculate_chart_size()
        
        # Position pie chart with guardrails
        pie_chart_left = SLIDE_WIDTH - chart_width - SLIDE_MARGIN
        pie_chart_top = CONTENT_TOP
        pie_chart_left, pie_chart_top, chart_width, chart_height = ensure_content_fits(
            pie_chart_left, pie_chart_top, chart_width, chart_height
        )
        
        pie_chart_image = self._create_pie_chart()
        slide.shapes.add_picture(pie_chart_image, pie_chart_left, pie_chart_top, width=chart_width)
        
        country_data = self._prepare_country_sentiment_data()
        if country_data:
            barchart_image = self._create_country_sentiment_barchart(country_data)
            
            # Position bar chart with guardrails
            barchart_width = CONTENT_MAX_WIDTH
            barchart_height = Inches(2.5)
            barchart_left = SLIDE_MARGIN
            barchart_top = SLIDE_HEIGHT - barchart_height - SLIDE_MARGIN
            
            barchart_left, barchart_top, barchart_width, barchart_height = ensure_content_fits(
                barchart_left, barchart_top, barchart_width, barchart_height
            )
            
            slide.shapes.add_picture(barchart_image, barchart_left, barchart_top, 
                                   width=barchart_width, height=barchart_height)
        
        # Remove unused placeholders to eliminate "Click to add text" boxes
        remove_unused_placeholders(slide)

    def add_paginated_impact_slide(self):
        slide_title = "Potential Business Impact"
        headers = ["Country", "ESG Theme", "Impact Area", "Level", "Rationale"]
        # Use .get() for safety in case the key is missing from the JSON
        rows_data = self.data.get('impactAnalysis', []) 
        if not rows_data:
            return

        # Flatten the data for the table
        all_rows = [[item['country'], item['theme'], item['impactArea'], item['impactLevel'], item['rationale']] for item in rows_data]

        total_pages = math.ceil(len(all_rows) / self.MAX_ROWS_PER_TABLE)
        for i in range(0, len(all_rows), self.MAX_ROWS_PER_TABLE):
            chunk = all_rows[i:i + self.MAX_ROWS_PER_TABLE]
            current_page = (i // self.MAX_ROWS_PER_TABLE) + 1
            
            paginated_title = f"{slide_title} (Page {current_page} of {total_pages})" if total_pages > 1 else slide_title
            slide = self.prs.slides.add_slide(self.prs.slide_layouts[1])
            slide.shapes.title.text = truncate_text_if_needed(paginated_title, 100)
            set_title_style(slide.shapes.title, self.prs.slide_width)
            
            # Use safe table dimensions
            table_left, table_top, table_width, table_height = calculate_safe_table_dimensions(
                len(chunk), len(headers)
            )
            
            table_shape = slide.shapes.add_table(
                len(chunk)+1, len(headers), table_left, table_top, table_width, table_height
            ).table
            self._set_header_row_height(table_shape, Inches(0.4))
            
            # Calculate column widths proportionally to fit within table width
            total_width_units = 2.0 + 2.0 + 3.0 + 1.0 + 6.8  # Original total
            table_width_value = table_width / Inches(1)  # Convert to numeric value
            
            table_shape.columns[0].width = Inches(table_width_value * (2.0 / total_width_units))
            table_shape.columns[1].width = Inches(table_width_value * (2.0 / total_width_units))
            table_shape.columns[2].width = Inches(table_width_value * (3.0 / total_width_units))
            table_shape.columns[3].width = Inches(table_width_value * (1.0 / total_width_units))
            table_shape.columns[4].width = Inches(table_width_value * (6.8 / total_width_units))
            
            for idx, h in enumerate(headers):
                self._set_cell_style(table_shape.cell(0, idx), truncate_text_if_needed(h, 50), is_header=True)
            
            for r_idx, row_data in enumerate(chunk):
                for c_idx, cell_text in enumerate(row_data):
                    is_dark_row = (r_idx % 2 == 1)
                    # Truncate cell text to prevent overflow
                    truncated_text = truncate_text_if_needed(str(cell_text), 200)
                    self._set_cell_style(table_shape.cell(r_idx + 1, c_idx), truncated_text, is_header=False, is_dark_row=is_dark_row)
            self._apply_uniform_row_heights_and_autofit(table_shape, Inches(0.5))
            
            # Remove unused placeholders to eliminate "Click to add text" boxes
            remove_unused_placeholders(slide)

    def add_paginated_regional_trends(self):
        slide_title = "Regional ESG & Sustainability Trends"
        headers = ["Region", "Country", "Category", "Theme"]
        all_rows = []
        for region, countries in self.data['regionalData'].items():
            for country, categories in countries.items():
                for category_name, themes in categories.items():
                    for theme_details in themes:
                        all_rows.append([region, country, category_name, theme_details['theme']])
        
        if not all_rows:
            return

        total_pages = math.ceil(len(all_rows) / self.MAX_ROWS_PER_TABLE)
        for i in range(0, len(all_rows), self.MAX_ROWS_PER_TABLE):
            chunk = all_rows[i:i + self.MAX_ROWS_PER_TABLE]
            current_page = (i // self.MAX_ROWS_PER_TABLE) + 1
            
            paginated_title = f"{slide_title} (Page {current_page} of {total_pages})" if total_pages > 1 else slide_title
            slide = self.prs.slides.add_slide(self.prs.slide_layouts[1])
            slide.shapes.title.text = truncate_text_if_needed(paginated_title, 100)
            set_title_style(slide.shapes.title, self.prs.slide_width)
            
            # Use safe table dimensions
            table_left, table_top, table_width, table_height = calculate_safe_table_dimensions(
                len(chunk), len(headers)
            )
            
            table = slide.shapes.add_table(
                len(chunk)+1, len(headers), table_left, table_top, table_width, table_height
            ).table
            set_header_row_height(self, table, Inches(0.4))
            
            # Calculate column widths proportionally
            total_width_units = 2.0 + 2.0 + 3.0 + 8.0  # Original total
            table_width_value = table_width / Inches(1)  # Convert to numeric value
            
            table.columns[0].width = Inches(table_width_value * (2.0 / total_width_units))
            table.columns[1].width = Inches(table_width_value * (2.0 / total_width_units))
            table.columns[2].width = Inches(table_width_value * (3.0 / total_width_units))
            table.columns[3].width = Inches(table_width_value * (8.0 / total_width_units))
            
            for idx, h in enumerate(headers):
                self._set_cell_style(table.cell(0, idx), truncate_text_if_needed(h, 50), is_header=True)
            
            for r_idx, row_data in enumerate(chunk):
                for c_idx, cell_text in enumerate(row_data):
                    is_dark_row = (r_idx % 2 == 1)
                    # Truncate cell text to prevent overflow
                    truncated_text = truncate_text_if_needed(str(cell_text), 150)
                    self._set_cell_style(table.cell(r_idx + 1, c_idx), truncated_text, is_header=False, is_dark_row=is_dark_row)
            self._apply_uniform_row_heights_and_autofit(table, Inches(0.5))
            
            # Remove unused placeholders to eliminate "Click to add text" boxes
            remove_unused_placeholders(slide)

    def add_slide_section(self, append=""):
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[2])
        slide_title = slide.shapes.title
        slide_title.text = "Appendix"
        slide_title.text_frame.paragraphs[0].font.name = heading_font
        slide_title.text_frame.paragraphs[0].font.size = Pt(54)
        slide_title.text_frame.paragraphs[0].font.color.rgb = DEFAULT_TEXT_COLOR
        if len(slide.placeholders) > 1:
            subtitle_text = "Supporting Data and Sources"
            subtitle_font = slide.placeholders[1].text_frame.paragraphs[0].font
            subtitle_font.name = text_font
            subtitle_font.color.rgb = DEFAULT_TEXT_COLOR
        
        # Remove unused placeholders to eliminate "Click to add text" boxes
        remove_unused_placeholders(slide)

    def add_sentiment_justification_slides(self):
        all_rows = []
        # for _countries in self.data["regionalData"]:
        # for category in countries.items():
        # for _themes in categories.items():
        # for theme in themes:
        # all_rows.append([country, theme["theme"], theme["sentiment"], theme["justification"]])
        
        # Sentiments to process ["Positive", "Neutral", "Negative"]
        sentiments_to_process = ["Positive", "Neutral", "Negative"]
        # """Retrieving all justification rows from all rows in all rows"""
        for region in self.data["regionalData"]:
            for country in self.data["regionalData"][region]:
                for category in self.data["regionalData"][region][country]:
                    for theme in self.data["regionalData"][region][country][category]:
                        if isinstance(theme, dict):  # Ensure theme is a dictionary
                            _sentiment = theme.get("sentiment")
                            if _sentiment in sentiments_to_process:
                                all_rows.append([country, theme.get("theme"), _sentiment, theme.get("justification")])
        
        if not all_rows:
            return

        total_pages = math.ceil(len(all_rows) / self.MAX_ROWS_PER_TABLE)
        for i in range(0, len(all_rows), self.MAX_ROWS_PER_TABLE):
            chunk = all_rows[i:i + self.MAX_ROWS_PER_TABLE]
            current_page = (i // self.MAX_ROWS_PER_TABLE) + 1
            slide_title = "Sentiment Justification"
            
            justification_title = f"{slide_title} (Page {current_page} of {total_pages})" if total_pages > 1 else slide_title
            slide = self.prs.slides.add_slide(self.prs.slide_layouts[1])
            slide.shapes.title.text = truncate_text_if_needed(justification_title, 100)
            set_title_style(slide.shapes.title, self.prs.slide_width)
            
            headers = ["Country", "Theme", "Sentiment", "Justification"]
            
            # Use safe table dimensions
            table_left, table_top, table_width, table_height = calculate_safe_table_dimensions(
                len(chunk), len(headers)
            )
            
            table = slide.shapes.add_table(
                len(chunk) + 1, len(headers), table_left, table_top, table_width, table_height
            ).table
            set_header_row_height(self, table, Inches(0.4))
            
            # Calculate column widths proportionally
            total_width_units = 2.0 + 6.5 + 1.5 + 5.0  # Original total
            table_width_value = table_width / Inches(1)  # Convert to numeric value
            
            table.columns[0].width = Inches(table_width_value * (2.0 / total_width_units))
            table.columns[1].width = Inches(table_width_value * (6.5 / total_width_units))
            table.columns[2].width = Inches(table_width_value * (1.5 / total_width_units))
            table.columns[3].width = Inches(table_width_value * (5.0 / total_width_units))
            
            for idx, h in enumerate(headers):
                set_cell_style(self, table.cell(0, idx), truncate_text_if_needed(h, 50), is_header_false=True)
            
            for r_idx, row_data in enumerate(chunk):
                for c_idx, cell_text in enumerate(row_data):
                    is_dark_row = (r_idx % 2 == 1)
                    # Truncate cell text to prevent overflow, especially for justification column
                    max_length = 300 if c_idx == 3 else 150  # Longer for justification column
                    truncated_text = truncate_text_if_needed(str(cell_text), max_length)
                    set_cell_style(self, table.cell(r_idx + 1, c_idx), truncated_text, is_header_false=False, is_dark_row=is_dark_row)
            self._apply_uniform_row_heights_and_autofit(table, Inches(0.5))
            
            # Remove unused placeholders to eliminate "Click to add text" boxes
            remove_unused_placeholders(slide)

    def add_paginated_sources(self):
        slide_title = "Sources and Reliability"
        headers = ["Source", "Reliability Score", "Justification"]
        sources_data = self.data.get("dataSources", [])
        if not sources_data:
            return

        # Convert dictionary data to rows
        rows_data = []
        for source in sources_data:
            rows_data.append([
                source.get("source", ""),
                source.get("reliabilityScore", ""),
                source.get("justification", "")
            ])

        total_pages = math.ceil(len(sources_data) / self.MAX_ROWS_PER_TABLE)
        for i in range(0, len(sources_data), self.MAX_ROWS_PER_TABLE):
            chunk_sources = sources_data[i:i + self.MAX_ROWS_PER_TABLE]
            chunk = rows_data[i:i + self.MAX_ROWS_PER_TABLE]
            current_page = (i // self.MAX_ROWS_PER_TABLE) + 1
            
            paginated_title = f"{slide_title} (Page {current_page} of {total_pages})" if total_pages > 1 else slide_title
            slide = self.prs.slides.add_slide(self.prs.slide_layouts[1])
            slide.shapes.title.text = truncate_text_if_needed(paginated_title, 100)
            set_title_style(slide.shapes.title, self.prs.slide_width)
            
            # Use safe table dimensions
            table_left, table_top, table_width, table_height = calculate_safe_table_dimensions(
                len(chunk), len(headers)
            )
            
            table = slide.shapes.add_table(
                len(chunk)+1, len(headers), table_left, table_top, table_width, table_height
            ).table
            set_header_row_height(self, table, Inches(0.4))

            # Calculate column widths proportionally
            total_width_units = 5.0 + 3.0 + 7.0  # Original total
            table_width_value = table_width / Inches(1)  # Convert to numeric value
            
            table.columns[0].width = Inches(table_width_value * (5.0 / total_width_units))
            table.columns[1].width = Inches(table_width_value * (3.0 / total_width_units))
            table.columns[2].width = Inches(table_width_value * (7.0 / total_width_units))
            
            for idx, h in enumerate(headers):
                set_cell_style(self, table.cell(0, idx), truncate_text_if_needed(h, 50), is_header_false=True)
            
            for r_idx, (row_data, source_data) in enumerate(zip(chunk, chunk_sources)):
                for c_idx, cell_text in enumerate(row_data):
                    is_dark_row = (r_idx % 2 == 1)
                    cell = table.cell(r_idx + 1, c_idx)
                    
                    # Truncate cell text to prevent overflow
                    max_length = 250 if c_idx == 2 else 100  # Longer for justification column
                    truncated_text = truncate_text_if_needed(str(cell_text), max_length)
                    cell.text = truncated_text
                    
                    p = cell.text_frame.paragraphs[0]
                    p.alignment = PP_ALIGN.LEFT
                    p.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
                    fill = cell.fill
                    fill.solid()
                    if is_dark_row:
                        fill.fore_color.rgb = ROW_COLOR_DARK
                    else:
                        fill.background()
                    font = p.font
                    font.color.rgb = DEFAULT_TEXT_COLOR
                    font.name = text_font
                    font.size = Pt(10)  # Slightly smaller font to fit better
                    cell.margin_left = Inches(0.05)  # Smaller margins
                    cell.margin_right = Inches(0.05)

                    # Add hyperlink for source column if URL exists
                    if c_idx == 0 and source_data.get("url"):  # Source column
                        url = source_data.get("url")
                        run = p.add_run()
                        run.text = f"\n{url}"
                        run.hyperlink.address = url
                        font = run.font
                        font.color.rgb = HYPERLINK_COLOR
                        font.underline = True

            self._apply_uniform_row_heights_and_autofit(table, Inches(0.4))
            
            # Remove unused placeholders to eliminate "Click to add text" boxes
            remove_unused_placeholders(slide)
            
            # --- CHANGE 1: Adjusted table header styling ---
            for idx, h in enumerate(headers):
                set_cell_style(self, table.cell(0, idx), h, is_header_false=True)
            
            # --- CHANGE 2: Looped through rows and added content with styling ---
            for r_idx, row_data in enumerate(chunk):
                for c_idx, cell_text in enumerate(row_data):
                    is_dark_row = (r_idx % 2 == 1)
                    cell = table.cell(r_idx + 1, c_idx)
            