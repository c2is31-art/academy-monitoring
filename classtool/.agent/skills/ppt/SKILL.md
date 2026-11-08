---
name: ppt
description: >-
  Create, edit, format, and inspect PowerPoint presentations (.pptx) using python-pptx.
  Use when the user requests generating slides, building presentation decks, creating PPT reports,
  converting summaries into presentations, or automating PowerPoint documents.
---

# PowerPoint (PPTX) Automation Skill

This skill provides comprehensive instructions and reusable python-pptx patterns to generate clean, modern, and professional PowerPoint presentations.

## 1. Quick Start

### Python Environment & Library
The `python-pptx` library is installed and available:
```bash
python -c "import pptx; print(pptx.__version__)"
```

### Standard Dimensions (16:9 Widescreen)
Always set 16:9 widescreen layout for modern displays:
```python
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank_slide_layout = prs.slide_layouts[6]  # completely blank slide
```

---

## 2. Design System & Style Guidelines

### A. Professional Color Palettes
Choose a cohesive palette based on the presentation topic:
- **Navy Professional (Business/Reporting - Recommended)**:
  - Primary Dark: `RGBColor(15, 30, 54)` (`#0F1E36`)
  - Accent Blue: `RGBColor(37, 99, 235)` (`#2563EB`)
  - Light Background: `RGBColor(248, 250, 252)` (`#F8FAFC`)
  - Card/Box Fill: `RGBColor(255, 255, 255)` (`#FFFFFF`)
  - Border/Line: `RGBColor(226, 232, 240)` (`#E2E8F0`)
  - Dark Text: `RGBColor(30, 41, 59)` (`#1E293B`)
  - Subdued Text: `RGBColor(100, 116, 139)` (`#64748B`)
- **Emerald Growth (Finance/Performance)**:
  - Primary: `RGBColor(6, 78, 59)` / Accent: `RGBColor(16, 185, 129)`
- **Crimson Impact (Urgent/Executive Summary)**:
  - Primary: `RGBColor(136, 19, 55)` / Accent: `RGBColor(225, 29, 72)`

### B. Typography
For Korean & English bilingual clarity:
- **Title Font**: `맑은 고딕` (Malgun Gothic) or `Pretendard`, 28pt ~ 38pt, Bold
- **Section/Subtitle Font**: 18pt ~ 22pt, Medium/Bold
- **Body Text**: 14pt ~ 16pt, Regular
- **Captions & Footnotes**: 11pt ~ 12pt, Subdued Gray

---

## 3. Reusable Slide Patterns

### Pattern 1: Title Slide (Cover)
```python
def create_title_slide(prs, title, subtitle, presenter="", date_str=""):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # Background accent banner or card
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), prs.slide_width, prs.slide_height)
    bg.fill.solid()
    bg.fill.fore_color.rgb = RGBColor(15, 30, 54)  # Deep Navy
    bg.line.fill.background()
    
    # Title box
    tx_box = slide.shapes.add_textbox(Inches(1.2), Inches(2.2), Inches(10.9), Inches(2.2))
    tf = tx_box.text_frame
    tf.word_wrap = True
    
    p = tf.paragraphs[0]
    p.text = title
    p.font.name = "맑은 고딕"
    p.font.size = Pt(40)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)
    
    # Subtitle
    p2 = tf.add_paragraph()
    p2.text = subtitle
    p2.font.name = "맑은 고딕"
    p2.font.size = Pt(20)
    p2.font.color.rgb = RGBColor(147, 197, 253)  # Soft Blue
    p2.space_before = Pt(14)
    
    # Meta (Presenter & Date)
    if presenter or date_str:
        meta_box = slide.shapes.add_textbox(Inches(1.2), Inches(5.8), Inches(10.9), Inches(0.8))
        tf_meta = meta_box.text_frame
        p_meta = tf_meta.paragraphs[0]
        p_meta.text = f"{presenter}  |  {date_str}".strip(" | ")
        p_meta.font.name = "맑은 고딕"
        p_meta.font.size = Pt(13)
        p_meta.font.color.rgb = RGBColor(203, 213, 225)
```

### Pattern 2: Slide Header (Standard for all content slides)
```python
def add_slide_header(slide, title, category=""):
    # Category tag
    if category:
        cat_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(11.5), Inches(0.35))
        tf_c = cat_box.text_frame
        tf_c.word_wrap = True
        p_c = tf_c.paragraphs[0]
        p_c.text = category.upper()
        p_c.font.name = "맑은 고딕"
        p_c.font.size = Pt(11)
        p_c.font.bold = True
        p_c.font.color.rgb = RGBColor(37, 99, 235)  # Accent Blue

    # Main Title
    t_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.8), Inches(11.5), Inches(0.7))
    tf = t_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title
    p.font.name = "맑은 고딕"
    p.font.size = Pt(24)
    p.font.bold = True
    p.font.color.rgb = RGBColor(15, 23, 42)
```

### Pattern 3: Multi-Column Feature Cards (2, 3, or 4 Columns)
```python
def add_card_columns(slide, cards_data, top_inch=1.8, height_inch=4.8):
    """
    cards_data: list of dicts [{'title': '...', 'badge': '...', 'bullets': [...]}]
    """
    num_cards = len(cards_data)
    total_width = 11.733  # 13.333 - 2 * 0.8 margins
    gap = 0.3
    card_width = (total_width - (gap * (num_cards - 1))) / num_cards
    
    for i, c_data in enumerate(cards_data):
        left = 0.8 + i * (card_width + gap)
        
        # Outer Card Box
        card = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(left), Inches(top_inch), Inches(card_width), Inches(height_inch)
        )
        card.fill.solid()
        card.fill.fore_color.rgb = RGBColor(255, 255, 255)
        card.line.color.rgb = RGBColor(226, 232, 240)
        card.line.width = Pt(1.5)
        
        # Text Inside Card
        tb = slide.shapes.add_textbox(Inches(left + 0.25), Inches(top_inch + 0.25), Inches(card_width - 0.5), Inches(height_inch - 0.5))
        tf = tb.text_frame
        tf.word_wrap = True
        
        # Header / Title
        p0 = tf.paragraphs[0]
        p0.text = c_data.get('title', '')
        p0.font.name = "맑은 고딕"
        p0.font.size = Pt(18)
        p0.font.bold = True
        p0.font.color.rgb = RGBColor(15, 23, 42)
        
        # Bullets
        for b in c_data.get('bullets', []):
            p = tf.add_paragraph()
            p.text = f"• {b}"
            p.font.name = "맑은 고딕"
            p.font.size = Pt(13.5)
            p.font.color.rgb = RGBColor(71, 85, 105)
            p.space_before = Pt(8)
```

### Pattern 4: Data Table Slide
```python
def add_table_slide(slide, headers, rows_data, left=0.8, top=1.8, width=11.733, height=4.5):
    rows = len(rows_data) + 1
    cols = len(headers)
    table_shape = slide.shapes.add_table(rows, cols, Inches(left), Inches(top), Inches(width), Inches(height))
    table = table_shape.table
    
    # Header format
    for j, h in enumerate(headers):
        cell = table.cell(0, j)
        cell.text = str(h)
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(30, 41, 59)
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.font.name = "맑은 고딕"
        p.font.bold = True
        p.font.size = Pt(13)
        p.font.color.rgb = RGBColor(255, 255, 255)
        
    # Data rows format
    for i, row in enumerate(rows_data, start=1):
        for j, val in enumerate(row):
            cell = table.cell(i, j)
            cell.text = str(val)
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(248, 250, 252) if i % 2 == 1 else RGBColor(255, 255, 255)
            p = cell.text_frame.paragraphs[0]
            p.font.name = "맑은 고딕"
            p.font.size = Pt(12)
            p.font.color.rgb = RGBColor(30, 41, 59)
            if j == 0:
                p.alignment = PP_ALIGN.CENTER
```

---

## 4. Execution Workflow

When generating PPTX for the user:
1. **Outline Structure**: Create a logical 4~10 slide outline (Cover, Agenda/Overview, Core Insights/Data, Comparison/Details, Summary/Action Items).
2. **Write Script**: Write a self-contained Python script utilizing `python-pptx` with the design patterns above.
3. **Execute & Verify**: Run the script to generate the `.pptx` file and verify there are no syntax or layout overlap errors.
4. **Deliver Output**: Provide the generated `.pptx` file path as a direct clickable link for the user.

