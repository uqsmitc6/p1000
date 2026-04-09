"""
UQ PowerPoint Template Layout Recipes

This module catalogues all 46 UQ PowerPoint template layouts with detailed
"recipes" for how to use each one. The recipes include content slot mappings,
matching criteria, and scoring rules to support automatic layout selection
for slide rebuilding.

Each layout can be matched against incoming slide content to determine the
best template layout to apply during automatic slide reconstruction.
"""

from dataclasses import dataclass, field
from typing import Dict, Tuple, Optional


@dataclass
class ContentSlot:
    """Configuration for a content placeholder in a layout."""
    ph_idx: int
    content_type: str  # "title", "subtitle", "body", "image", "table", "object"
    required: bool = False
    max_chars: Optional[int] = None
    description: str = ""


@dataclass
class MatchCriteria:
    """Scoring rules for matching content to a layout."""
    is_first_slide: int = 0
    is_last_slide: int = 0
    has_title: int = 0
    has_subtitle: int = 0
    has_body_text: int = 0
    has_single_image: int = 0
    has_multiple_images: int = 0
    image_count_range: Optional[Tuple[int, int]] = None
    has_table: int = 0
    text_heavy: int = 0  # Lots of text, minimal images
    image_heavy: int = 0  # Mainly images
    has_quote_pattern: int = 0
    body_text_count_range: Optional[Tuple[int, int]] = None
    is_section_break: int = 0  # Short title, minimal body
    has_two_columns: int = 0
    has_three_columns: int = 0
    minimal_content: int = 0  # Very short content (blank/title-only)
    complex_layout_penalty: int = 0  # Penalise complex layouts for simple content


@dataclass
class LayoutRecipe:
    """Complete recipe for applying a template layout to slide content."""
    name: str
    index: int
    category: str  # "cover", "divider", "content", "image", "table", "quote", "ending", "special", "blank"
    description: str
    content_slots: Dict[str, ContentSlot] = field(default_factory=dict)
    match_criteria: MatchCriteria = field(default_factory=MatchCriteria)
    priority: int = 100  # Lower = preferred in tiebreaker
    skip_matching: bool = False  # If True, don't auto-match this layout


# ============================================================================
# COVER LAYOUTS (0-2)
# ============================================================================

LAYOUT_0_COVER_1 = LayoutRecipe(
    name="Cover 1",
    index=0,
    category="cover",
    description="Simple cover slide with title and subtitle stacked. Best for basic presentations.",
    content_slots={
        "title": ContentSlot(
            ph_idx=0,
            content_type="title",
            required=True,
            max_chars=80,
            description="Main presentation title (0.55\", 3.07\", 5.1\"x1.31\")"
        ),
        "subtitle_1": ContentSlot(
            ph_idx=11,
            content_type="subtitle",
            required=False,
            max_chars=60,
            description="Subtitle line 1 (0.55\", 2.34\", 5.1\"x0.54\")"
        ),
        "subtitle_2": ContentSlot(
            ph_idx=10,
            content_type="subtitle",
            required=False,
            max_chars=60,
            description="Subtitle line 2 (0.55\", 4.69\", 5.1\"x0.7\")"
        ),
    },
    match_criteria=MatchCriteria(
        is_first_slide=100,
        has_title=50,
        minimal_content=30,
    ),
    priority=10,
)

LAYOUT_1_COVER_2 = LayoutRecipe(
    name="Cover 2",
    index=1,
    category="cover",
    description="Wide cover slide with title spanning full width and subtitle below.",
    content_slots={
        "title": ContentSlot(
            ph_idx=0,
            content_type="title",
            required=True,
            max_chars=100,
            description="Main presentation title (0.52\", 1.68\", 8.19\"x1.2\")"
        ),
        "logo_text": ContentSlot(
            ph_idx=11,
            content_type="body",
            required=False,
            max_chars=40,
            description="Logo/org reference text (0.52\", 0.5\", 4.1\"x0.41\")"
        ),
        "subtitle": ContentSlot(
            ph_idx=10,
            content_type="subtitle",
            required=False,
            max_chars=70,
            description="Subtitle (0.52\", 3.52\", 8.19\"x0.7\")"
        ),
    },
    match_criteria=MatchCriteria(
        is_first_slide=100,
        has_title=50,
        minimal_content=30,
    ),
    priority=11,
)

LAYOUT_2_COVER_3 = LayoutRecipe(
    name="Cover 3",
    index=2,
    category="cover",
    description="Cover slide with title, subtitle, and image placeholder on right side.",
    content_slots={
        "title": ContentSlot(
            ph_idx=0,
            content_type="title",
            required=True,
            max_chars=80,
            description="Main presentation title (0.55\", 3.15\", 4.62\"x1.2\")"
        ),
        "subtitle_1": ContentSlot(
            ph_idx=11,
            content_type="subtitle",
            required=False,
            max_chars=60,
            description="Subtitle (0.55\", 2.17\", 4.59\"x0.54\")"
        ),
        "subtitle_2": ContentSlot(
            ph_idx=10,
            content_type="subtitle",
            required=False,
            max_chars=60,
            description="Subtitle line 2 (0.55\", 4.77\", 4.62\"x0.7\")"
        ),
        "image": ContentSlot(
            ph_idx=12,
            content_type="image",
            required=False,
            description="Cover image on right (5.84\", 0.0\", 7.5\"x7.5\")"
        ),
    },
    match_criteria=MatchCriteria(
        is_first_slide=100,
        has_title=50,
        has_single_image=40,
        minimal_content=30,
    ),
    priority=12,
)

# ============================================================================
# SECTION DIVIDER LAYOUT (5)
# ============================================================================

LAYOUT_5_SECTION_DIVIDER = LayoutRecipe(
    name="Section Divider",
    index=5,
    category="divider",
    description="Section divider with number and description. Use for breaking presentations into sections.",
    content_slots={
        "title": ContentSlot(
            ph_idx=0,
            content_type="title",
            required=True,
            max_chars=50,
            description="Section title (0.55\", 2.88\", 4.95\"x1.2\")"
        ),
        "section_number": ContentSlot(
            ph_idx=11,
            content_type="body",
            required=False,
            max_chars=10,
            description="Section number (0.55\", 2.33\", 1.95\"x0.38\")"
        ),
        "description": ContentSlot(
            ph_idx=13,
            content_type="body",
            required=False,
            max_chars=100,
            description="Section description (7.81\", 2.89\", 3.59\"x1.36\")"
        ),
        "footer": ContentSlot(
            ph_idx=10,
            content_type="body",
            required=False,
            description="Footer text"
        ),
        "slide_number": ContentSlot(
            ph_idx=14,
            content_type="body",
            required=False,
            description="Slide number"
        ),
    },
    match_criteria=MatchCriteria(
        is_section_break=80,
        has_title=60,
        minimal_content=50,
    ),
    priority=20,
)

# ============================================================================
# MAIN CONTENT LAYOUTS (6-16)
# These are the workhorses - highly detailed recipes
# ============================================================================

LAYOUT_6_TITLE_AND_CONTENT = LayoutRecipe(
    name="Title and Content",
    index=6,
    category="content",
    description="Standard layout with title at top and single content area. Most versatile layout for text, images, tables, or objects.",
    content_slots={
        "title": ContentSlot(
            ph_idx=0,
            content_type="title",
            required=True,
            max_chars=100,
            description="Slide title (0.52\", 0.98\", 9.76\"x0.51\")"
        ),
        "subtitle": ContentSlot(
            ph_idx=31,
            content_type="subtitle",
            required=False,
            max_chars=100,
            description="Optional subtitle/secondary heading (0.52\", 1.72\", 9.76\"x0.55\")"
        ),
        "content": ContentSlot(
            ph_idx=10,
            content_type="object",
            required=True,
            description="Main content area - accepts text, images, tables, shapes (0.52\", 2.47\", 12.28\"x4.51\")"
        ),
        "footer": ContentSlot(
            ph_idx=17,
            content_type="body",
            required=False,
            description="Footer text"
        ),
        "slide_number": ContentSlot(
            ph_idx=18,
            content_type="body",
            required=False,
            description="Slide number"
        ),
    },
    match_criteria=MatchCriteria(
        has_title=40,
        has_body_text=30,
        has_single_image=25,
        has_table=30,
        text_heavy=20,
        minimal_content=-50,
        body_text_count_range=(1, 5),
    ),
    priority=50,
)

LAYOUT_7_TWO_CONTENT = LayoutRecipe(
    name="Two Content",
    index=7,
    category="content",
    description="Side-by-side layout with title, subtitle, and two equal content areas. Perfect for comparisons or parallel information.",
    content_slots={
        "title": ContentSlot(
            ph_idx=0,
            content_type="title",
            required=True,
            max_chars=100,
            description="Slide title (0.52\", 0.98\", 9.76\"x0.51\")"
        ),
        "subtitle": ContentSlot(
            ph_idx=31,
            content_type="subtitle",
            required=False,
            max_chars=100,
            description="Optional subtitle (0.52\", 1.72\", 9.76\"x0.55\")"
        ),
        "content_left": ContentSlot(
            ph_idx=10,
            content_type="object",
            required=True,
            description="Left content area (0.52\", 2.47\", 5.75\"x4.51\")"
        ),
        "content_right": ContentSlot(
            ph_idx=32,
            content_type="object",
            required=True,
            description="Right content area (7.06\", 2.47\", 5.75\"x4.51\")"
        ),
        "footer": ContentSlot(
            ph_idx=17,
            content_type="body",
            required=False,
            description="Footer text"
        ),
        "slide_number": ContentSlot(
            ph_idx=18,
            content_type="body",
            required=False,
            description="Slide number"
        ),
    },
    match_criteria=MatchCriteria(
        has_title=40,
        has_two_columns=70,
        has_multiple_images=25,
        text_heavy=15,
    ),
    priority=51,
)

LAYOUT_8_THREE_CONTENT = LayoutRecipe(
    name="Three Content",
    index=8,
    category="content",
    description="Layout with title and three content areas. Use for comparing three options or displaying three parallel concepts.",
    content_slots={
        "title": ContentSlot(
            ph_idx=0,
            content_type="title",
            required=True,
            description="Slide title"
        ),
        "subtitle": ContentSlot(
            ph_idx=31,
            content_type="subtitle",
            required=False,
            description="Optional subtitle"
        ),
        "content_left": ContentSlot(
            ph_idx=10,
            content_type="object",
            required=True,
            description="Left content area"
        ),
        "content_centre": ContentSlot(
            ph_idx=35,
            content_type="object",
            required=True,
            description="Centre content area"
        ),
        "content_right": ContentSlot(
            ph_idx=34,
            content_type="object",
            required=True,
            description="Right content area"
        ),
        "footer": ContentSlot(
            ph_idx=17,
            content_type="body",
            required=False,
            description="Footer text"
        ),
        "slide_number": ContentSlot(
            ph_idx=18,
            content_type="body",
            required=False,
            description="Slide number"
        ),
    },
    match_criteria=MatchCriteria(
        has_title=40,
        has_three_columns=80,
        has_multiple_images=20,
    ),
    priority=52,
)

LAYOUT_13_TWO_GRAPHS = LayoutRecipe(
    name="Title, Subtitle, 2 Graphs",
    index=13,
    category="content",
    description="Layout optimized for two side-by-side charts with labels. Ideal for comparing metrics or KPIs.",
    content_slots={
        "title": ContentSlot(
            ph_idx=0,
            content_type="title",
            required=True,
            description="Slide title"
        ),
        "subtitle": ContentSlot(
            ph_idx=31,
            content_type="subtitle",
            required=False,
            description="Optional subtitle"
        ),
        "chart_left": ContentSlot(
            ph_idx=10,
            content_type="object",
            required=True,
            description="Left chart/graph"
        ),
        "label_left": ContentSlot(
            ph_idx=32,
            content_type="body",
            required=False,
            description="Left chart label"
        ),
        "chart_right": ContentSlot(
            ph_idx=33,
            content_type="object",
            required=True,
            description="Right chart/graph"
        ),
        "label_right": ContentSlot(
            ph_idx=34,
            content_type="body",
            required=False,
            description="Right chart label"
        ),
    },
    match_criteria=MatchCriteria(
        has_title=40,
        has_multiple_images=60,
        image_count_range=(2, 2),
    ),
    priority=53,
)

LAYOUT_14_TWO_CONTENT_HORIZONTAL = LayoutRecipe(
    name="Two Content Horizontal",
    index=14,
    category="content",
    description="Stacked layout with title and two full-width content areas (top and bottom). Use for before/after or sequential content.",
    content_slots={
        "title": ContentSlot(
            ph_idx=0,
            content_type="title",
            required=True,
            description="Slide title"
        ),
        "subtitle": ContentSlot(
            ph_idx=31,
            content_type="subtitle",
            required=False,
            description="Optional subtitle"
        ),
        "content_top": ContentSlot(
            ph_idx=10,
            content_type="object",
            required=True,
            description="Top content area"
        ),
        "content_bottom": ContentSlot(
            ph_idx=34,
            content_type="object",
            required=True,
            description="Bottom content area"
        ),
    },
    match_criteria=MatchCriteria(
        has_title=40,
        has_multiple_images=25,
    ),
    priority=54,
)

LAYOUT_15_ONE_THIRD_TWO_THIRD = LayoutRecipe(
    name="One Third Two Third",
    index=15,
    category="content",
    description="Layout with left sidebar (1/3) and main content area (2/3). Use for focus on main content with supporting sidebar.",
    content_slots={
        "title": ContentSlot(
            ph_idx=0,
            content_type="title",
            required=True,
            description="Slide title"
        ),
        "subtitle": ContentSlot(
            ph_idx=31,
            content_type="subtitle",
            required=False,
            description="Optional subtitle"
        ),
        "content_left_small": ContentSlot(
            ph_idx=10,
            content_type="object",
            required=True,
            description="Left sidebar (1/3 width)"
        ),
        "content_right_large": ContentSlot(
            ph_idx=35,
            content_type="object",
            required=True,
            description="Main content (2/3 width)"
        ),
    },
    match_criteria=MatchCriteria(
        has_title=40,
        has_two_columns=50,
    ),
    priority=55,
)

LAYOUT_16_TWO_THIRD_ONE_THIRD = LayoutRecipe(
    name="Two Third One Third",
    index=16,
    category="content",
    description="Layout with main content area (2/3) on left and sidebar (1/3) on right. Use when main content dominates.",
    content_slots={
        "title": ContentSlot(
            ph_idx=0,
            content_type="title",
            required=True,
            description="Slide title"
        ),
        "subtitle": ContentSlot(
            ph_idx=31,
            content_type="subtitle",
            required=False,
            description="Optional subtitle"
        ),
        "content_left_large": ContentSlot(
            ph_idx=10,
            content_type="object",
            required=True,
            description="Main content (2/3 width)"
        ),
        "content_right_small": ContentSlot(
            ph_idx=34,
            content_type="object",
            required=True,
            description="Right sidebar (1/3 width)"
        ),
    },
    match_criteria=MatchCriteria(
        has_title=40,
        has_two_columns=50,
    ),
    priority=56,
)

LAYOUT_40_TITLE_AND_TABLE = LayoutRecipe(
    name="Title and Table",
    index=40,
    category="table",
    description="Optimized layout for tables. Title at top with full-width table below.",
    content_slots={
        "title": ContentSlot(
            ph_idx=0,
            content_type="title",
            required=True,
            description="Slide title"
        ),
        "subtitle": ContentSlot(
            ph_idx=31,
            content_type="subtitle",
            required=False,
            description="Optional subtitle"
        ),
        "table": ContentSlot(
            ph_idx=19,
            content_type="table",
            required=True,
            description="Data table (0.53\", 2.45\", 12.28\"x4.53\")"
        ),
    },
    match_criteria=MatchCriteria(
        has_title=40,
        has_table=100,
    ),
    priority=40,
)

# ============================================================================
# IMAGE LAYOUTS (17-26)
# ============================================================================

LAYOUT_17_PICTURE_WITH_PULLOUT = LayoutRecipe(
    name="Picture with Pullout",
    index=17,
    category="image",
    description="Large background image with text pullout on left. Creates dramatic visual impact.",
    content_slots={
        "title": ContentSlot(
            ph_idx=0,
            content_type="title",
            required=True,
            max_chars=60,
            description="Small title on left"
        ),
        "text": ContentSlot(
            ph_idx=31,
            content_type="body",
            required=False,
            max_chars=200,
            description="Left text content"
        ),
        "background_image": ContentSlot(
            ph_idx=20,
            content_type="image",
            required=True,
            description="Full background image"
        ),
        "pullout_image": ContentSlot(
            ph_idx=10,
            content_type="image",
            required=False,
            description="Optional overlay image"
        ),
    },
    match_criteria=MatchCriteria(
        has_single_image=80,
        image_heavy=50,
    ),
    priority=65,
)

LAYOUT_18_PICTURE_WITH_CAPTION = LayoutRecipe(
    name="Picture with Caption",
    index=18,
    category="image",
    description="Image with caption below. Clean layout for showcasing a single image with explanatory text.",
    content_slots={
        "title": ContentSlot(
            ph_idx=0,
            content_type="title",
            required=True,
            description="Slide title"
        ),
        "subtitle": ContentSlot(
            ph_idx=31,
            content_type="subtitle",
            required=False,
            description="Optional subtitle"
        ),
        "image": ContentSlot(
            ph_idx=32,
            content_type="image",
            required=True,
            description="Main image (0.52\", 2.47\", 12.28\"x4.51\")"
        ),
        "caption": ContentSlot(
            ph_idx=2,
            content_type="body",
            required=False,
            max_chars=150,
            description="Caption at bottom (0.39\"h)"
        ),
        "footer": ContentSlot(
            ph_idx=17,
            content_type="body",
            required=False,
            description="Footer text"
        ),
        "slide_number": ContentSlot(
            ph_idx=18,
            content_type="body",
            required=False,
            description="Slide number"
        ),
    },
    match_criteria=MatchCriteria(
        has_title=30,
        has_single_image=90,
        image_heavy=60,
    ),
    priority=60,
)

LAYOUT_19_TEXT_WITH_IMAGE_TWO_THIRDS = LayoutRecipe(
    name="Text with Image Two Thirds",
    index=19,
    category="image",
    description="Left sidebar (1/3) with text, right main area (2/3) with full-height image.",
    content_slots={
        "title": ContentSlot(
            ph_idx=0,
            content_type="title",
            required=True,
            description="Slide title"
        ),
        "subtitle": ContentSlot(
            ph_idx=31,
            content_type="subtitle",
            required=False,
            description="Optional subtitle"
        ),
        "text_content": ContentSlot(
            ph_idx=10,
            content_type="object",
            required=True,
            description="Left text content"
        ),
        "image": ContentSlot(
            ph_idx=34,
            content_type="image",
            required=True,
            description="Right image (2/3 width, full-height)"
        ),
    },
    match_criteria=MatchCriteria(
        has_title=30,
        has_body_text=40,
        has_single_image=70,
    ),
    priority=61,
)

LAYOUT_20_TEXT_WITH_IMAGE_HALF = LayoutRecipe(
    name="Text with Image Half",
    index=20,
    category="image",
    description="Left half with text, right half with full-height image. Balanced layout.",
    content_slots={
        "title": ContentSlot(
            ph_idx=0,
            content_type="title",
            required=True,
            description="Slide title"
        ),
        "subtitle": ContentSlot(
            ph_idx=31,
            content_type="subtitle",
            required=False,
            description="Optional subtitle"
        ),
        "text_content": ContentSlot(
            ph_idx=10,
            content_type="object",
            required=True,
            description="Left text content (1/2 width)"
        ),
        "image": ContentSlot(
            ph_idx=32,
            content_type="image",
            required=True,
            description="Right image (1/2 width, full-height)"
        ),
    },
    match_criteria=MatchCriteria(
        has_title=30,
        has_body_text=40,
        has_single_image=70,
        has_two_columns=60,
    ),
    priority=62,
)

LAYOUT_21_TEXT_WITH_IMAGE_ONE_THIRD = LayoutRecipe(
    name="Text with Image One Third",
    index=21,
    category="image",
    description="Left main area (2/3) with text, right sidebar (1/3) with full-height image.",
    content_slots={
        "title": ContentSlot(
            ph_idx=0,
            content_type="title",
            required=True,
            description="Slide title"
        ),
        "subtitle": ContentSlot(
            ph_idx=31,
            content_type="subtitle",
            required=False,
            description="Optional subtitle"
        ),
        "text_content": ContentSlot(
            ph_idx=10,
            content_type="object",
            required=True,
            description="Left text content (2/3 width)"
        ),
        "image": ContentSlot(
            ph_idx=34,
            content_type="image",
            required=True,
            description="Right image (1/3 width, full-height)"
        ),
    },
    match_criteria=MatchCriteria(
        has_title=30,
        has_body_text=40,
        has_single_image=60,
    ),
    priority=63,
)

LAYOUT_22_TEXT_WITH_IMAGE_TWO_THIRDS_ALT = LayoutRecipe(
    name="Text with Image Two Thirds Alt",
    index=22,
    category="image",
    description="Alternative arrangement: left (1/3) title/subtitle, main text below, right (2/3) image.",
    content_slots={
        "title": ContentSlot(
            ph_idx=0,
            content_type="title",
            required=True,
            description="Slide title"
        ),
        "subtitle": ContentSlot(
            ph_idx=31,
            content_type="subtitle",
            required=False,
            description="Optional subtitle"
        ),
        "text_content": ContentSlot(
            ph_idx=33,
            content_type="object",
            required=True,
            description="Left text content"
        ),
        "image": ContentSlot(
            ph_idx=34,
            content_type="image",
            required=True,
            description="Right image (2/3 width)"
        ),
    },
    match_criteria=MatchCriteria(
        has_title=30,
        has_body_text=40,
        has_single_image=70,
    ),
    priority=64,
)

LAYOUT_23_TEXT_WITH_IMAGE_HALF_ALT = LayoutRecipe(
    name="Text with Image Half Alt",
    index=23,
    category="image",
    description="Alternative: left text below title/subtitle, right (1/2) image.",
    content_slots={
        "title": ContentSlot(
            ph_idx=0,
            content_type="title",
            required=True,
            description="Slide title"
        ),
        "subtitle": ContentSlot(
            ph_idx=31,
            content_type="subtitle",
            required=False,
            description="Optional subtitle"
        ),
        "text_content": ContentSlot(
            ph_idx=33,
            content_type="object",
            required=True,
            description="Left text content"
        ),
        "image": ContentSlot(
            ph_idx=32,
            content_type="image",
            required=True,
            description="Right image (1/2 width)"
        ),
    },
    match_criteria=MatchCriteria(
        has_title=30,
        has_body_text=40,
        has_single_image=70,
        has_two_columns=60,
    ),
    priority=65,
)

LAYOUT_24_TEXT_WITH_IMAGE_ONE_THIRD_ALT = LayoutRecipe(
    name="Text with Image One Third Alt",
    index=24,
    category="image",
    description="Alternative: left text area with title/subtitle, right (1/3) image.",
    content_slots={
        "title": ContentSlot(
            ph_idx=0,
            content_type="title",
            required=True,
            description="Slide title"
        ),
        "subtitle": ContentSlot(
            ph_idx=31,
            content_type="subtitle",
            required=False,
            description="Optional subtitle"
        ),
        "text_content": ContentSlot(
            ph_idx=10,
            content_type="object",
            required=True,
            description="Left text content (2/3 width)"
        ),
        "image": ContentSlot(
            ph_idx=34,
            content_type="image",
            required=True,
            description="Right image (1/3 width)"
        ),
    },
    match_criteria=MatchCriteria(
        has_title=30,
        has_body_text=40,
        has_single_image=60,
    ),
    priority=66,
)

LAYOUT_25_TEXT_WITH_FOUR_IMAGES = LayoutRecipe(
    name="Text with 4 Images",
    index=25,
    category="image",
    description="Left sidebar with text, right side with 2x2 grid of images.",
    content_slots={
        "title": ContentSlot(
            ph_idx=0,
            content_type="title",
            required=True,
            description="Slide title"
        ),
        "text_content": ContentSlot(
            ph_idx=10,
            content_type="object",
            required=True,
            description="Left text area"
        ),
        "image_1": ContentSlot(
            ph_idx=12,
            content_type="image",
            required=True,
            description="Top-left image in grid"
        ),
        "image_2": ContentSlot(
            ph_idx=16,
            content_type="image",
            required=True,
            description="Top-right image in grid"
        ),
        "image_3": ContentSlot(
            ph_idx=17,
            content_type="image",
            required=True,
            description="Bottom-left image in grid"
        ),
        "image_4": ContentSlot(
            ph_idx=18,
            content_type="image",
            required=True,
            description="Bottom-right image in grid"
        ),
    },
    match_criteria=MatchCriteria(
        has_title=30,
        has_body_text=20,
        has_multiple_images=100,
        image_count_range=(4, 4),
        image_heavy=70,
    ),
    priority=67,
)

LAYOUT_26_THREE_COLUMN_TEXT_AND_IMAGES = LayoutRecipe(
    name="Three Column Text & Images",
    index=26,
    category="image",
    description="Three-column layout with text blocks and images below each column.",
    content_slots={
        "title": ContentSlot(
            ph_idx=0,
            content_type="title",
            required=True,
            description="Slide title"
        ),
        "subtitle": ContentSlot(
            ph_idx=53,
            content_type="subtitle",
            required=False,
            description="Optional subtitle"
        ),
        "text_1": ContentSlot(
            ph_idx=10,
            content_type="object",
            required=True,
            description="Left column text"
        ),
        "text_2": ContentSlot(
            ph_idx=34,
            content_type="object",
            required=True,
            description="Centre column text"
        ),
        "text_3": ContentSlot(
            ph_idx=35,
            content_type="object",
            required=True,
            description="Right column text"
        ),
        "image_1": ContentSlot(
            ph_idx=14,
            content_type="image",
            required=True,
            description="Left column image"
        ),
        "image_2": ContentSlot(
            ph_idx=16,
            content_type="image",
            required=True,
            description="Centre column image"
        ),
        "image_3": ContentSlot(
            ph_idx=18,
            content_type="image",
            required=True,
            description="Right column image"
        ),
    },
    match_criteria=MatchCriteria(
        has_title=40,
        has_three_columns=100,
        has_multiple_images=80,
        image_count_range=(3, 3),
    ),
    priority=68,
)

LAYOUT_41_IMAGE_COLLAGE = LayoutRecipe(
    name="Image Collage",
    index=41,
    category="image",
    description="Gallery layout for multiple images in collage arrangement with optional caption.",
    content_slots={
        "title": ContentSlot(
            ph_idx=0,
            content_type="title",
            required=True,
            description="Slide title"
        ),
        "image_1": ContentSlot(
            ph_idx=14,
            content_type="image",
            required=True,
            description="Collage image 1"
        ),
        "image_2": ContentSlot(
            ph_idx=21,
            content_type="image",
            required=True,
            description="Collage image 2"
        ),
        "image_3": ContentSlot(
            ph_idx=39,
            content_type="image",
            required=True,
            description="Collage image 3"
        ),
        "image_4": ContentSlot(
            ph_idx=40,
            content_type="image",
            required=True,
            description="Collage image 4"
        ),
        "caption": ContentSlot(
            ph_idx=38,
            content_type="body",
            required=False,
            description="Optional caption text"
        ),
    },
    match_criteria=MatchCriteria(
        has_title=30,
        has_multiple_images=100,
        image_count_range=(4, 4),
        image_heavy=80,
    ),
    priority=69,
)

# ============================================================================
# VISUAL EMPHASIS LAYOUTS (29-37)
# ============================================================================

LAYOUT_29_GRAPH_WITH_DARK_PURPLE_BLOCK = LayoutRecipe(
    name="Graph with Dark Purple Block",
    index=29,
    category="content",
    description="Left chart/graph with right dark purple callout block. Use for highlighting key metrics.",
    content_slots={
        "title": ContentSlot(
            ph_idx=0,
            content_type="title",
            required=True,
            description="Slide title"
        ),
        "subtitle": ContentSlot(
            ph_idx=31,
            content_type="subtitle",
            required=False,
            description="Optional subtitle"
        ),
        "chart": ContentSlot(
            ph_idx=10,
            content_type="object",
            required=True,
            description="Left chart/graph"
        ),
        "callout": ContentSlot(
            ph_idx=14,
            content_type="body",
            required=True,
            description="Right dark purple callout block"
        ),
    },
    match_criteria=MatchCriteria(
        has_title=40,
        has_single_image=50,
    ),
    priority=70,
)

LAYOUT_30_GRAPH_WITH_NEUTRAL_BLOCK = LayoutRecipe(
    name="Graph with Neutral Block",
    index=30,
    category="content",
    description="Left chart with right neutral-coloured callout block.",
    content_slots={
        "title": ContentSlot(
            ph_idx=0,
            content_type="title",
            required=True,
            description="Slide title"
        ),
        "subtitle": ContentSlot(
            ph_idx=31,
            content_type="subtitle",
            required=False,
            description="Optional subtitle"
        ),
        "chart": ContentSlot(
            ph_idx=10,
            content_type="object",
            required=True,
            description="Left chart/graph"
        ),
        "callout": ContentSlot(
            ph_idx=14,
            content_type="body",
            required=True,
            description="Right neutral callout block"
        ),
    },
    match_criteria=MatchCriteria(
        has_title=40,
        has_single_image=50,
    ),
    priority=71,
)

LAYOUT_31_GRAPH_WITH_GREY_BLOCK = LayoutRecipe(
    name="Graph with Grey Block",
    index=31,
    category="content",
    description="Left chart with right grey callout block.",
    content_slots={
        "title": ContentSlot(
            ph_idx=0,
            content_type="title",
            required=True,
            description="Slide title"
        ),
        "subtitle": ContentSlot(
            ph_idx=31,
            content_type="subtitle",
            required=False,
            description="Optional subtitle"
        ),
        "chart": ContentSlot(
            ph_idx=10,
            content_type="object",
            required=True,
            description="Left chart/graph"
        ),
        "callout": ContentSlot(
            ph_idx=14,
            content_type="body",
            required=True,
            description="Right grey callout block"
        ),
    },
    match_criteria=MatchCriteria(
        has_title=40,
        has_single_image=50,
    ),
    priority=72,
)

LAYOUT_32_TEXT_WITH_DARK_PURPLE_BLOCK = LayoutRecipe(
    name="Text with Dark Purple Block",
    index=32,
    category="content",
    description="Left text content with right dark purple emphasis block.",
    content_slots={
        "title": ContentSlot(
            ph_idx=0,
            content_type="title",
            required=True,
            description="Slide title"
        ),
        "subtitle": ContentSlot(
            ph_idx=31,
            content_type="subtitle",
            required=False,
            description="Optional subtitle"
        ),
        "text_content": ContentSlot(
            ph_idx=1,
            content_type="object",
            required=True,
            description="Left body text"
        ),
        "callout": ContentSlot(
            ph_idx=14,
            content_type="body",
            required=True,
            description="Right dark purple callout block"
        ),
    },
    match_criteria=MatchCriteria(
        has_title=40,
        has_body_text=60,
    ),
    priority=73,
)

LAYOUT_33_TEXT_WITH_NEUTRAL_BLOCK = LayoutRecipe(
    name="Text with Neutral Block",
    index=33,
    category="content",
    description="Left text content with right neutral callout block.",
    content_slots={
        "title": ContentSlot(
            ph_idx=0,
            content_type="title",
            required=True,
            description="Slide title"
        ),
        "subtitle": ContentSlot(
            ph_idx=31,
            content_type="subtitle",
            required=False,
            description="Optional subtitle"
        ),
        "text_content": ContentSlot(
            ph_idx=1,
            content_type="object",
            required=True,
            description="Left body text"
        ),
        "callout": ContentSlot(
            ph_idx=14,
            content_type="body",
            required=True,
            description="Right neutral callout block"
        ),
    },
    match_criteria=MatchCriteria(
        has_title=40,
        has_body_text=60,
    ),
    priority=74,
)

LAYOUT_34_TEXT_WITH_GREY_BLOCK = LayoutRecipe(
    name="Text with Grey Block",
    index=34,
    category="content",
    description="Left text content with right grey emphasis block.",
    content_slots={
        "title": ContentSlot(
            ph_idx=0,
            content_type="title",
            required=True,
            description="Slide title"
        ),
        "subtitle": ContentSlot(
            ph_idx=31,
            content_type="subtitle",
            required=False,
            description="Optional subtitle"
        ),
        "text_content": ContentSlot(
            ph_idx=1,
            content_type="object",
            required=True,
            description="Left body text"
        ),
        "callout": ContentSlot(
            ph_idx=14,
            content_type="body",
            required=True,
            description="Right grey callout block"
        ),
    },
    match_criteria=MatchCriteria(
        has_title=40,
        has_body_text=60,
    ),
    priority=75,
)

LAYOUT_35_THREE_PULLOUTS = LayoutRecipe(
    name="Three Pullouts",
    index=35,
    category="content",
    description="Three key point callout blocks arranged horizontally. Use for highlighting three key takeaways.",
    content_slots={
        "title": ContentSlot(
            ph_idx=0,
            content_type="title",
            required=True,
            description="Slide title"
        ),
        "subtitle": ContentSlot(
            ph_idx=34,
            content_type="subtitle",
            required=False,
            description="Optional subtitle"
        ),
        "callout_1": ContentSlot(
            ph_idx=19,
            content_type="body",
            required=True,
            description="Left callout block"
        ),
        "callout_2": ContentSlot(
            ph_idx=20,
            content_type="body",
            required=True,
            description="Centre callout block"
        ),
        "callout_3": ContentSlot(
            ph_idx=21,
            content_type="body",
            required=True,
            description="Right callout block"
        ),
    },
    match_criteria=MatchCriteria(
        has_title=40,
        has_three_columns=80,
        has_body_text=50,
    ),
    priority=76,
)

LAYOUT_36_MULTI_LAYOUT_1 = LayoutRecipe(
    name="Multi-layout 1",
    index=36,
    category="content",
    description="Complex multi-element layout with top row of two elements and bottom row of two elements.",
    content_slots={
        "title": ContentSlot(
            ph_idx=0,
            content_type="title",
            required=True,
            description="Slide title"
        ),
        "subtitle": ContentSlot(
            ph_idx=34,
            content_type="subtitle",
            required=False,
            description="Optional subtitle"
        ),
        "element_top_left": ContentSlot(
            ph_idx=31,
            content_type="object",
            required=True,
            description="Top-left element"
        ),
        "element_top_right": ContentSlot(
            ph_idx=36,
            content_type="body",
            required=True,
            description="Top-right element"
        ),
        "element_bottom_left": ContentSlot(
            ph_idx=28,
            content_type="body",
            required=True,
            description="Bottom-left element"
        ),
        "element_bottom_right": ContentSlot(
            ph_idx=35,
            content_type="object",
            required=True,
            description="Bottom-right element"
        ),
    },
    match_criteria=MatchCriteria(
        has_title=40,
        complex_layout_penalty=30,
    ),
    priority=77,
)

LAYOUT_37_MULTI_LAYOUT_2 = LayoutRecipe(
    name="Multi-layout 2",
    index=37,
    category="content",
    description="Alternative multi-element layout with asymmetric element placement.",
    content_slots={
        "title": ContentSlot(
            ph_idx=0,
            content_type="title",
            required=True,
            description="Slide title"
        ),
        "subtitle": ContentSlot(
            ph_idx=31,
            content_type="subtitle",
            required=False,
            description="Optional subtitle"
        ),
        "element_top_left": ContentSlot(
            ph_idx=10,
            content_type="object",
            required=True,
            description="Top-left element"
        ),
        "element_bottom_left": ContentSlot(
            ph_idx=36,
            content_type="object",
            required=True,
            description="Bottom-left element"
        ),
        "element_right_1": ContentSlot(
            ph_idx=19,
            content_type="object",
            required=True,
            description="Right side element 1"
        ),
        "element_right_2": ContentSlot(
            ph_idx=20,
            content_type="object",
            required=True,
            description="Right side element 2"
        ),
        "label_left": ContentSlot(
            ph_idx=26,
            content_type="body",
            required=False,
            description="Left label"
        ),
        "label_right": ContentSlot(
            ph_idx=35,
            content_type="body",
            required=False,
            description="Right label"
        ),
    },
    match_criteria=MatchCriteria(
        has_title=40,
        complex_layout_penalty=30,
    ),
    priority=78,
)

# ============================================================================
# QUOTE LAYOUTS (38-39)
# ============================================================================

LAYOUT_38_QUOTE_1 = LayoutRecipe(
    name="Quote 1",
    index=38,
    category="quote",
    description="Full-screen background image with quote text and attribution overlay.",
    content_slots={
        "background": ContentSlot(
            ph_idx=10,
            content_type="image",
            required=True,
            description="Full-screen background image"
        ),
        "quote_text": ContentSlot(
            ph_idx=16,
            content_type="body",
            required=True,
            max_chars=200,
            description="Main quote text"
        ),
        "attribution": ContentSlot(
            ph_idx=17,
            content_type="body",
            required=True,
            max_chars=100,
            description="Quote attribution"
        ),
    },
    match_criteria=MatchCriteria(
        has_quote_pattern=100,
        has_single_image=80,
        image_heavy=60,
    ),
    priority=80,
)

LAYOUT_39_QUOTE_2 = LayoutRecipe(
    name="Quote 2",
    index=39,
    category="quote",
    description="Alternative quote layout with different background style.",
    content_slots={
        "background": ContentSlot(
            ph_idx=10,
            content_type="image",
            required=True,
            description="Full-screen background image"
        ),
        "quote_text": ContentSlot(
            ph_idx=16,
            content_type="body",
            required=True,
            max_chars=200,
            description="Main quote text"
        ),
        "attribution": ContentSlot(
            ph_idx=17,
            content_type="body",
            required=True,
            max_chars=100,
            description="Quote attribution"
        ),
    },
    match_criteria=MatchCriteria(
        has_quote_pattern=100,
        has_single_image=80,
        image_heavy=60,
    ),
    priority=81,
)

# ============================================================================
# ENDING/CLOSING LAYOUTS (42-43)
# ============================================================================

LAYOUT_42_THANK_YOU = LayoutRecipe(
    name="Thank You",
    index=42,
    category="ending",
    description="Thank you slide with contact information. Use as final slide.",
    content_slots={
        "title": ContentSlot(
            ph_idx=0,
            content_type="title",
            required=True,
            max_chars=50,
            description="Main message (usually 'Thank You')"
        ),
        "name": ContentSlot(
            ph_idx=10,
            content_type="body",
            required=True,
            max_chars=80,
            description="Contact person name"
        ),
        "title_role": ContentSlot(
            ph_idx=16,
            content_type="body",
            required=False,
            max_chars=80,
            description="Job title/role"
        ),
        "email": ContentSlot(
            ph_idx=17,
            content_type="body",
            required=True,
            max_chars=80,
            description="Email address"
        ),
        "phone": ContentSlot(
            ph_idx=18,
            content_type="body",
            required=False,
            max_chars=30,
            description="Phone number"
        ),
    },
    match_criteria=MatchCriteria(
        is_last_slide=100,
        minimal_content=50,
    ),
    priority=5,
)

LAYOUT_43_TITLE_ONLY = LayoutRecipe(
    name="Title Only",
    index=43,
    category="special",
    description="Simple layout with title only. Use for section breaks or minimal slides.",
    content_slots={
        "title": ContentSlot(
            ph_idx=0,
            content_type="title",
            required=True,
            description="Slide title"
        ),
        "footer": ContentSlot(
            ph_idx=17,
            content_type="body",
            required=False,
            description="Footer text"
        ),
        "slide_number": ContentSlot(
            ph_idx=18,
            content_type="body",
            required=False,
            description="Slide number"
        ),
    },
    match_criteria=MatchCriteria(
        has_title=40,
        minimal_content=80,
    ),
    priority=85,
)

LAYOUT_44_BLANK_BRANDED = LayoutRecipe(
    name="Blank Branded",
    index=44,
    category="blank",
    description="Completely blank slide with only footer and branding. For custom content.",
    content_slots={
        "footer": ContentSlot(
            ph_idx=10,
            content_type="body",
            required=False,
            description="Footer text"
        ),
        "slide_number": ContentSlot(
            ph_idx=11,
            content_type="body",
            required=False,
            description="Slide number"
        ),
    },
    match_criteria=MatchCriteria(
        minimal_content=100,
    ),
    priority=90,
)

# ============================================================================
# SKIPPED LAYOUTS (complex, too specific, or off-screen placeholders)
# ============================================================================

# Layout 3: Contents 1 - Skip (27 placeholders, complex grid TOC)
# Layout 4: Contents 2 - Skip (complex table of contents)
# Layout 9-12: Skip (very specific arrangements, complex icon grids)
# Layout 27-28: Skip (complex icon/order layouts)
# Layout 45: Process Diagram - Skip (off-screen text labels)

# ============================================================================
# BUILD RECIPE DICTIONARY AND CATEGORIES
# ============================================================================

RECIPES = {
    # Cover layouts
    "Cover 1": LAYOUT_0_COVER_1,
    "Cover 2": LAYOUT_1_COVER_2,
    "Cover 3": LAYOUT_2_COVER_3,
    # Section divider
    "Section Divider": LAYOUT_5_SECTION_DIVIDER,
    # Main content layouts
    "Title and Content": LAYOUT_6_TITLE_AND_CONTENT,
    "Two Content": LAYOUT_7_TWO_CONTENT,
    "Three Content": LAYOUT_8_THREE_CONTENT,
    "Title, Subtitle, 2 Graphs": LAYOUT_13_TWO_GRAPHS,
    "Two Content Horizontal": LAYOUT_14_TWO_CONTENT_HORIZONTAL,
    "One Third Two Third": LAYOUT_15_ONE_THIRD_TWO_THIRD,
    "Two Third One Third": LAYOUT_16_TWO_THIRD_ONE_THIRD,
    "Title and Table": LAYOUT_40_TITLE_AND_TABLE,
    # Image layouts
    "Picture with Pullout": LAYOUT_17_PICTURE_WITH_PULLOUT,
    "Picture with Caption": LAYOUT_18_PICTURE_WITH_CAPTION,
    "Text with Image Two Thirds": LAYOUT_19_TEXT_WITH_IMAGE_TWO_THIRDS,
    "Text with Image Half": LAYOUT_20_TEXT_WITH_IMAGE_HALF,
    "Text with Image One Third": LAYOUT_21_TEXT_WITH_IMAGE_ONE_THIRD,
    "Text with Image Two Thirds Alt": LAYOUT_22_TEXT_WITH_IMAGE_TWO_THIRDS_ALT,
    "Text with Image Half Alt": LAYOUT_23_TEXT_WITH_IMAGE_HALF_ALT,
    "Text with Image One Third Alt": LAYOUT_24_TEXT_WITH_IMAGE_ONE_THIRD_ALT,
    "Text with 4 Images": LAYOUT_25_TEXT_WITH_FOUR_IMAGES,
    "Three Column Text & Images": LAYOUT_26_THREE_COLUMN_TEXT_AND_IMAGES,
    "Image Collage": LAYOUT_41_IMAGE_COLLAGE,
    # Visual emphasis layouts
    "Graph with Dark Purple Block": LAYOUT_29_GRAPH_WITH_DARK_PURPLE_BLOCK,
    "Graph with Neutral Block": LAYOUT_30_GRAPH_WITH_NEUTRAL_BLOCK,
    "Graph with Grey Block": LAYOUT_31_GRAPH_WITH_GREY_BLOCK,
    "Text with Dark Purple Block": LAYOUT_32_TEXT_WITH_DARK_PURPLE_BLOCK,
    "Text with Neutral Block": LAYOUT_33_TEXT_WITH_NEUTRAL_BLOCK,
    "Text with Grey Block": LAYOUT_34_TEXT_WITH_GREY_BLOCK,
    "Three Pullouts": LAYOUT_35_THREE_PULLOUTS,
    "Multi-layout 1": LAYOUT_36_MULTI_LAYOUT_1,
    "Multi-layout 2": LAYOUT_37_MULTI_LAYOUT_2,
    # Quote layouts
    "Quote 1": LAYOUT_38_QUOTE_1,
    "Quote 2": LAYOUT_39_QUOTE_2,
    # Ending layouts
    "Thank You": LAYOUT_42_THANK_YOU,
    "Title Only": LAYOUT_43_TITLE_ONLY,
    "Blank Branded": LAYOUT_44_BLANK_BRANDED,
}

LAYOUT_CATEGORIES = {
    "cover": [
        LAYOUT_0_COVER_1,
        LAYOUT_1_COVER_2,
        LAYOUT_2_COVER_3,
    ],
    "divider": [
        LAYOUT_5_SECTION_DIVIDER,
    ],
    "content": [
        LAYOUT_6_TITLE_AND_CONTENT,
        LAYOUT_7_TWO_CONTENT,
        LAYOUT_8_THREE_CONTENT,
        LAYOUT_13_TWO_GRAPHS,
        LAYOUT_14_TWO_CONTENT_HORIZONTAL,
        LAYOUT_15_ONE_THIRD_TWO_THIRD,
        LAYOUT_16_TWO_THIRD_ONE_THIRD,
        LAYOUT_32_TEXT_WITH_DARK_PURPLE_BLOCK,
        LAYOUT_33_TEXT_WITH_NEUTRAL_BLOCK,
        LAYOUT_34_TEXT_WITH_GREY_BLOCK,
        LAYOUT_35_THREE_PULLOUTS,
        LAYOUT_36_MULTI_LAYOUT_1,
        LAYOUT_37_MULTI_LAYOUT_2,
    ],
    "image": [
        LAYOUT_17_PICTURE_WITH_PULLOUT,
        LAYOUT_18_PICTURE_WITH_CAPTION,
        LAYOUT_19_TEXT_WITH_IMAGE_TWO_THIRDS,
        LAYOUT_20_TEXT_WITH_IMAGE_HALF,
        LAYOUT_21_TEXT_WITH_IMAGE_ONE_THIRD,
        LAYOUT_22_TEXT_WITH_IMAGE_TWO_THIRDS_ALT,
        LAYOUT_23_TEXT_WITH_IMAGE_HALF_ALT,
        LAYOUT_24_TEXT_WITH_IMAGE_ONE_THIRD_ALT,
        LAYOUT_25_TEXT_WITH_FOUR_IMAGES,
        LAYOUT_26_THREE_COLUMN_TEXT_AND_IMAGES,
        LAYOUT_41_IMAGE_COLLAGE,
    ],
    "table": [
        LAYOUT_40_TITLE_AND_TABLE,
    ],
    "quote": [
        LAYOUT_38_QUOTE_1,
        LAYOUT_39_QUOTE_2,
    ],
    "ending": [
        LAYOUT_42_THANK_YOU,
    ],
    "special": [
        LAYOUT_43_TITLE_ONLY,
        LAYOUT_29_GRAPH_WITH_DARK_PURPLE_BLOCK,
        LAYOUT_30_GRAPH_WITH_NEUTRAL_BLOCK,
        LAYOUT_31_GRAPH_WITH_GREY_BLOCK,
    ],
    "blank": [
        LAYOUT_44_BLANK_BRANDED,
    ],
}

# Most commonly useful layouts - the ones that handle ~85% of typical slides
COMMON_LAYOUTS = [
    LAYOUT_6_TITLE_AND_CONTENT,      # Workhorse - handles most content
    LAYOUT_7_TWO_CONTENT,             # Common comparison layout
    LAYOUT_40_TITLE_AND_TABLE,        # Data tables
    LAYOUT_8_THREE_CONTENT,           # Three-way comparisons
    LAYOUT_20_TEXT_WITH_IMAGE_HALF,   # Image + text (balanced)
    LAYOUT_18_PICTURE_WITH_CAPTION,   # Single image focus
    LAYOUT_19_TEXT_WITH_IMAGE_TWO_THIRDS,  # Image-dominant
    LAYOUT_14_TWO_CONTENT_HORIZONTAL,      # Vertical stacking
    LAYOUT_35_THREE_PULLOUTS,              # Key points
    LAYOUT_0_COVER_1,                 # First slide (cover)
    LAYOUT_42_THANK_YOU,              # Last slide
    LAYOUT_43_TITLE_ONLY,             # Simple section breaks
    LAYOUT_5_SECTION_DIVIDER,         # Formal section breaks
    LAYOUT_32_TEXT_WITH_DARK_PURPLE_BLOCK,  # Emphasis
    LAYOUT_26_THREE_COLUMN_TEXT_AND_IMAGES, # Multi-column showcase
]


def score_layout_match(layout: LayoutRecipe, content_analysis: dict) -> int:
    """
    Score how well a layout matches the given content.

    Higher score = better match. This function should be called during
    slide analysis to determine the best layout for rebuilding.

    Args:
        layout: LayoutRecipe to score
        content_analysis: Dict with keys like:
            - is_first_slide: bool
            - is_last_slide: bool
            - has_title: bool
            - has_body_text: bool
            - num_body_blocks: int (count of text blocks)
            - has_images: bool
            - image_count: int
            - has_table: bool
            - slide_position: int (0-indexed position in deck)
            - deck_size: int (total slides)
            - is_mostly_text: bool
            - is_mostly_image: bool
            - has_quote_pattern: bool

    Returns:
        Integer score (higher = better match)
    """
    score = 0
    criteria = layout.match_criteria

    # Position-based scoring
    if content_analysis.get("is_first_slide"):
        score += criteria.is_first_slide
    if content_analysis.get("is_last_slide"):
        score += criteria.is_last_slide

    # Content type scoring
    if content_analysis.get("has_title"):
        score += criteria.has_title
    if content_analysis.get("has_subtitle"):
        score += criteria.has_subtitle
    if content_analysis.get("has_body_text"):
        score += criteria.has_body_text
    if content_analysis.get("has_table"):
        score += criteria.has_table

    # Image scoring
    image_count = content_analysis.get("image_count", 0)
    if image_count == 1:
        score += criteria.has_single_image
    elif image_count >= 2:
        score += criteria.has_multiple_images

    # Exact image count matching
    if criteria.image_count_range:
        min_imgs, max_imgs = criteria.image_count_range
        if min_imgs <= image_count <= max_imgs:
            score += 20
        else:
            score -= 30

    # Content volume scoring
    if content_analysis.get("is_mostly_text"):
        score += criteria.text_heavy
    if content_analysis.get("is_mostly_image"):
        score += criteria.image_heavy

    # Column-based scoring
    num_content_blocks = content_analysis.get("num_content_blocks", 0)
    if num_content_blocks == 2:
        score += criteria.has_two_columns
    elif num_content_blocks == 3:
        score += criteria.has_three_columns

    # Body text count scoring
    if criteria.body_text_count_range:
        min_blocks, max_blocks = criteria.body_text_count_range
        if min_blocks <= num_content_blocks <= max_blocks:
            score += 15

    # Section break detection
    if content_analysis.get("is_section_break"):
        score += criteria.is_section_break

    # Quote pattern
    if content_analysis.get("has_quote_pattern"):
        score += criteria.has_quote_pattern

    # Minimal content (blank/title-only slides)
    if content_analysis.get("is_minimal_content"):
        score += criteria.minimal_content

    # Complex layout penalty for simple content
    if content_analysis.get("is_minimal_content"):
        score -= criteria.complex_layout_penalty

    return score
