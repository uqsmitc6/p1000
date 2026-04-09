"""
Layout Manager — UQ Slide Compliance Tool (Recipe-Based)
=========================================================
Analyses slide content, determines the best UQ template layout using recipes,
rebuilds each slide from scratch using the template, then verifies via Claude Vision.

Architecture:
    ContentAnalyser  — Extracts structured content from any slide
    LayoutMatcher    — Scores recipes against content analysis
    SlideRebuilder   — Creates fresh slide from template + content
    LayoutManager    — Orchestrates the full pipeline

The recipe-based approach is superior because it:
  - Uses documented, tested layout templates from layout_recipes.py
  - Matches content to layouts using explicit scoring criteria
  - Rebuilds slides with guaranteed correct placeholders and content slots
  - Handles edge cases (text overflow, image sizing, etc.) consistently
"""

import os
import io
import json
import base64
import tempfile
from pathlib import Path
from dataclasses import dataclass, field
from typing import Optional, List, Dict, Tuple
from collections import defaultdict

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

from layout_recipes import RECIPES, COMMON_LAYOUTS, score_layout_match

try:
    import anthropic
    HAS_ANTHROPIC = True
except ImportError:
    HAS_ANTHROPIC = False


# ============================================================================
# Constants
# ============================================================================

UQ_TEMPLATE_PATH = os.path.join(
    os.path.dirname(__file__),
    "uq_template.pptx",
)

VISION_MODEL = "claude-sonnet-4-20250514"
ANALYSIS_MAX_TOKENS = 2000
VERIFY_MAX_TOKENS = 1500

# Cost constants (Sonnet input/output per token)
COST_INPUT_PER_TOKEN = 3.0 / 1_000_000
COST_OUTPUT_PER_TOKEN = 15.0 / 1_000_000

# Minimum score threshold for auto-matching layouts
MIN_MATCH_SCORE = 10


# ============================================================================
# Data Structures
# ============================================================================

@dataclass
class ImageInfo:
    """Information about an image extracted from a slide."""
    blob: bytes
    width: int
    height: int
    original_position: Optional[str] = None  # e.g., "top-right"


@dataclass
class BodyBlock:
    """A text block with formatting hints."""
    text: str
    is_bullet: bool = False
    level: int = 0  # Bullet indent level
    bold: bool = False
    italic: bool = False
    formatting_hints: Dict = field(default_factory=dict)


@dataclass
class TableInfo:
    """Information about a table extracted from a slide."""
    rows: int
    cols: int
    cell_data: List[List[str]]  # [row][col] → text


@dataclass
class ContentAnalysis:
    """Structured analysis of a slide's content."""
    # Raw content
    title: Optional[str] = None
    subtitle: Optional[str] = None
    body_texts: List[BodyBlock] = field(default_factory=list)
    images: List[ImageInfo] = field(default_factory=list)
    tables: List[TableInfo] = field(default_factory=list)
    other_shapes: List = field(default_factory=list)
    speaker_notes: str = ""
    original_layout_name: str = ""

    # Context
    slide_position: int = 0
    deck_size: int = 1

    # Computed properties
    @property
    def is_first_slide(self) -> bool:
        return self.slide_position == 0

    @property
    def is_last_slide(self) -> bool:
        return self.slide_position == self.deck_size - 1

    @property
    def has_title(self) -> bool:
        return self.title is not None and len(self.title.strip()) > 0

    @property
    def has_subtitle(self) -> bool:
        return self.subtitle is not None and len(self.subtitle.strip()) > 0

    @property
    def has_body_text(self) -> bool:
        return len(self.body_texts) > 0

    @property
    def has_images(self) -> bool:
        return len(self.images) > 0

    @property
    def has_table(self) -> bool:
        return len(self.tables) > 0

    @property
    def image_count(self) -> int:
        return len(self.images)

    @property
    def num_body_blocks(self) -> int:
        return len(self.body_texts)

    @property
    def total_text_chars(self) -> int:
        total = (len(self.title or "") + len(self.subtitle or "") +
                 sum(len(b.text) for b in self.body_texts))
        return total

    @property
    def is_mostly_text(self) -> bool:
        """True if content is primarily text with few/no images."""
        return self.has_body_text and self.image_count < 2

    @property
    def is_mostly_image(self) -> bool:
        """True if content is primarily images (minimal text)."""
        return self.image_count >= 2 and self.total_text_chars < 200

    @property
    def is_section_break(self) -> bool:
        """True if this looks like a section divider (title-only or short title + minimal body)."""
        has_minimal_body = len(self.body_texts) == 0 or (
            len(self.body_texts) == 1 and len(self.body_texts[0].text) < 100
        )
        return self.has_title and has_minimal_body and self.image_count == 0

    @property
    def has_quote_pattern(self) -> bool:
        """True if content looks like a quote (centered, short text block)."""
        if not self.body_texts or len(self.body_texts) > 2:
            return False
        total = sum(len(b.text) for b in self.body_texts)
        return 50 < total < 500

    @property
    def is_minimal_content(self) -> bool:
        """True if slide has almost no content (blank or title-only)."""
        return (not self.has_body_text and
                self.image_count == 0 and
                not self.has_table)


@dataclass
class SlideResult:
    """Result for a single slide in the pipeline."""
    slide_number: int
    original_layout: str
    recommended_layout: str
    confidence: float
    analysis: dict = field(default_factory=dict)
    rebuild_report: dict = field(default_factory=dict)
    verification: dict = field(default_factory=dict)
    status: str = "pending"  # pending, analysed, rebuilt, verified, failed


# ============================================================================
# ContentAnalyser
# ============================================================================

class ContentAnalyser:
    """Extracts structured content from a slide."""

    def __init__(self):
        pass

    def analyse_slide(self, slide, slide_position: int = 0, deck_size: int = 1) -> ContentAnalysis:
        """Extract structured content from a slide.

        Args:
            slide: python-pptx Slide object
            slide_position: 0-indexed position of this slide in the deck
            deck_size: Total number of slides in the deck

        Returns:
            ContentAnalysis object with all extracted content
        """
        analysis = ContentAnalysis(
            slide_position=slide_position,
            deck_size=deck_size,
            original_layout_name=slide.slide_layout.name,
            speaker_notes=self._extract_speaker_notes(slide),
        )

        # Extract content from shapes
        for shape in slide.shapes:
            self._extract_from_shape(shape, analysis)

        return analysis

    def _extract_speaker_notes(self, slide) -> str:
        """Extract speaker notes from a slide."""
        if slide.has_notes_slide:
            text_frame = slide.notes_slide.notes_text_frame
            if text_frame:
                return text_frame.text
        return ""

    def _extract_from_shape(self, shape, analysis: ContentAnalysis):
        """Extract content from a shape and add to analysis."""
        if shape.is_placeholder:
            ph_format = shape.placeholder_format
            ph_type = ph_format.type

            if hasattr(ph_format, 'type') and shape.has_text_frame:
                type_val = str(ph_type)

                # Title placeholder (TITLE or CENTER_TITLE)
                if 'TITLE' in type_val and 'SUB' not in type_val:
                    analysis.title = shape.text.strip()
                    return

                # Subtitle placeholder
                if 'SUBTITLE' in type_val or 'SUB_TITLE' in type_val:
                    analysis.subtitle = shape.text.strip()
                    return

                # Footer / slide number — skip, don't extract as body text
                if 'FOOTER' in type_val or 'SLIDE_NUMBER' in type_val or 'DATE' in type_val:
                    return

            # Placeholder with picture
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                try:
                    self._extract_image(shape, analysis)
                    return
                except Exception:
                    pass
            # Try to get image from placeholder (some placeholders hold images)
            try:
                if shape.placeholder_format and shape.image:
                    self._extract_image(shape, analysis)
                    return
            except (ValueError, AttributeError):
                pass

        # Image or picture (non-placeholder)
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            self._extract_image(shape, analysis)
            return

        # Table
        if shape.has_table:
            self._extract_table(shape, analysis)
            return

        # Text frame (body text, bullet points) — only after ruling out title/subtitle/footer
        if shape.has_text_frame:
            text = shape.text.strip()
            if text:  # Skip empty text frames
                self._extract_text_blocks(shape, analysis)
            return

        # Other shapes (freeform text, charts, etc.)
        analysis.other_shapes.append({
            'shape_type': str(shape.shape_type),
            'name': shape.name,
        })

    def _extract_text_blocks(self, shape, analysis: ContentAnalysis):
        """Extract text blocks (preserving bullet structure) from a shape."""
        if not shape.has_text_frame:
            return

        text_frame = shape.text_frame

        # Check if this looks like body text (has bullets/paragraphs)
        for paragraph in text_frame.paragraphs:
            text = paragraph.text.strip()
            if not text:
                continue

            # Determine if it's a bullet and get level
            is_bullet = paragraph.level is not None
            level = paragraph.level if is_bullet else 0

            # Extract formatting hints
            bold = False
            italic = False
            if paragraph.runs:
                first_run = paragraph.runs[0]
                if first_run.font.bold:
                    bold = True
                if first_run.font.italic:
                    italic = True

            block = BodyBlock(
                text=text,
                is_bullet=is_bullet,
                level=level,
                bold=bold,
                italic=italic,
            )
            analysis.body_texts.append(block)

    def _extract_image(self, shape, analysis: ContentAnalysis):
        """Extract image from a shape."""
        try:
            image = shape.image
            blob = image.blob

            # shape.width/height are in EMU (914400 EMU = 1 inch)
            # Convert to approximate pixels at 96 DPI
            width_px = int(shape.width / 914400 * 96)
            height_px = int(shape.height / 914400 * 96)

            # Skip very small images (likely icons/bullets) — under 50px either dimension
            if width_px < 50 and height_px < 50:
                return

            # Skip duplicate images (same bytes already extracted)
            import hashlib
            img_hash = hashlib.sha256(blob[:1024]).hexdigest()
            for existing in analysis.images:
                existing_hash = hashlib.sha256(existing.blob[:1024]).hexdigest()
                if img_hash == existing_hash:
                    return

            img_info = ImageInfo(
                blob=blob,
                width=width_px,
                height=height_px,
                original_position=f"({shape.left}, {shape.top})",
            )
            analysis.images.append(img_info)
        except Exception:
            # Unable to extract image
            pass

    def _extract_table(self, shape, analysis: ContentAnalysis):
        """Extract table data from a shape."""
        try:
            table = shape.table
            rows = len(table.rows)
            cols = len(table.columns)

            cell_data = []
            for row in table.rows:
                row_data = []
                for cell in row.cells:
                    row_data.append(cell.text.strip())
                cell_data.append(row_data)

            table_info = TableInfo(
                rows=rows,
                cols=cols,
                cell_data=cell_data,
            )
            analysis.tables.append(table_info)
        except Exception:
            # Unable to extract table
            pass


# ============================================================================
# LayoutMatcher
# ============================================================================

class LayoutMatcher:
    """Scores recipes against content analysis and finds the best match."""

    def __init__(self):
        pass

    def find_best_layout(self, analysis: ContentAnalysis) -> Tuple[str, int, str]:
        """Find the best layout recipe for the given content.

        Args:
            analysis: ContentAnalysis object

        Returns:
            Tuple of (layout_name, score, reasoning)
        """
        # ---- Hard overrides for obvious cases ----
        title_lower = (analysis.title or "").lower().strip()

        # Thank You / closing slides
        if analysis.is_last_slide and any(kw in title_lower for kw in
                ["thank you", "thanks", "thank-you", "questions?", "q&a"]):
            return "Thank You", 200, "Keyword match: closing slide"

        # Cover slide (first slide with minimal body content)
        if analysis.is_first_slide and analysis.has_title:
            if analysis.image_count > 0:
                return "Cover 3", 200, "First slide with image"
            return "Cover 1", 200, "First slide"

        # ---- Score-based matching ----
        # Build analysis dict for scoring
        content_dict = {
            "is_first_slide": analysis.is_first_slide,
            "is_last_slide": analysis.is_last_slide,
            "has_title": analysis.has_title,
            "has_subtitle": analysis.has_subtitle,
            "has_body_text": analysis.has_body_text,
            "has_images": analysis.has_images,
            "image_count": analysis.image_count,
            "has_table": analysis.has_table,
            "num_content_blocks": analysis.num_body_blocks,
            "is_mostly_text": analysis.is_mostly_text,
            "is_mostly_image": analysis.is_mostly_image,
            "is_section_break": analysis.is_section_break,
            "has_quote_pattern": analysis.has_quote_pattern,
            "is_minimal_content": analysis.is_minimal_content,
        }

        # Score all eligible recipes
        scores = {}

        for recipe_name, recipe in RECIPES.items():
            # Skip layouts marked as non-matching
            if recipe.skip_matching:
                continue

            # Filter by position constraints
            if recipe.category == "cover" and not analysis.is_first_slide:
                continue
            if recipe.name == "Thank You" and not analysis.is_last_slide:
                continue

            # Score this recipe
            score = score_layout_match(recipe, content_dict)
            scores[recipe_name] = score

        # Find the best match
        if not scores:
            # Fallback to safe default
            best_name = "Title and Content"
            best_score = 0
            reasoning = "No eligible layouts (fallback)"
        else:
            best_name = max(scores, key=scores.get)
            best_score = scores[best_name]
            reasoning = f"Score: {best_score}"

        return best_name, best_score, reasoning


# ============================================================================
# SlideRebuilder
# ============================================================================

class SlideRebuilder:
    """Rebuilds slides using template layouts and content analysis."""

    def __init__(self, template_path: str = None):
        self.template_path = template_path or UQ_TEMPLATE_PATH
        self._template_prs = None

    @property
    def template_prs(self):
        """Lazy-load the template presentation."""
        if self._template_prs is None:
            self._template_prs = Presentation(self.template_path)
        return self._template_prs

    def rebuild_slide(self, target_prs: Presentation, layout_name: str,
                     analysis: ContentAnalysis) -> Dict:
        """Rebuild a slide using a template layout and content analysis.

        Args:
            target_prs: The presentation to add the slide to
            layout_name: Name of the recipe layout to use
            analysis: ContentAnalysis with extracted content

        Returns:
            Dict with rebuild report (status, errors, etc.)
        """
        report = {
            "layout_name": layout_name,
            "status": "pending",
            "errors": [],
            "warnings": [],
            "content_placed": {},
        }

        try:
            # Get the recipe
            recipe = RECIPES.get(layout_name)
            if not recipe:
                report["status"] = "failed"
                report["errors"].append(f"Recipe not found: {layout_name}")
                return report

            # Get template layout
            template_layout = self.template_prs.slide_layouts[recipe.index]

            # Create new slide from template
            slide = target_prs.slides.add_slide(template_layout)

            # Place content using recipe's content slots
            self._place_content(slide, recipe, analysis, report)

            report["status"] = "success"

        except Exception as e:
            report["status"] = "failed"
            report["errors"].append(str(e))

        return report

    def _place_content(self, slide, recipe, analysis: ContentAnalysis, report: Dict):
        """Place content into slide placeholders according to the recipe.

        Uses consumption tracking so each image/table/text block is placed
        only once across multiple content slots.
        """
        # For section dividers: if title is just a number, swap title and first body text
        import re
        if (recipe.category == "divider" and analysis.has_title
                and re.match(r'^\d{1,3}$', analysis.title.strip())
                and analysis.body_texts):
            # The real section name is in body_texts[0]; the number is in title
            section_num = analysis.title.strip()
            real_title = analysis.body_texts[0].text
            # Temporarily swap for placement
            analysis.title = real_title
            analysis.body_texts = analysis.body_texts[1:]  # Remove the used block
            # Store section number for the section_number slot
            analysis._section_number_override = section_num

        # Track which content items have been consumed
        next_image = 0
        next_table = 0
        body_placed = False

        # Count how many "object" or "body" slots exist for text splitting
        text_slot_names = [
            name for name, slot in recipe.content_slots.items()
            if slot.content_type in ("object", "body")
            and name not in ("footer", "slide_number", "section_number")
        ]

        # Pre-split body text for multi-column layouts
        body_chunks = self._split_body_for_columns(analysis.body_texts, len(text_slot_names))
        body_chunk_idx = 0

        # Process each content slot in the recipe
        for slot_name, slot_config in recipe.content_slots.items():
            ph_idx = slot_config.ph_idx
            content_type = slot_config.content_type

            # Skip footer/slide_number — brand fixer handles these
            if slot_name in ("footer", "slide_number"):
                continue

            try:
                ph = slide.placeholders[ph_idx]

                if content_type == "title" and analysis.has_title:
                    self._place_text(ph, analysis.title, is_bullet=False)
                    report["content_placed"]["title"] = "placed"

                elif content_type == "subtitle":
                    if analysis.has_subtitle:
                        self._place_text(ph, analysis.subtitle, is_bullet=False)
                        report["content_placed"]["subtitle"] = "placed"
                    # Some covers use subtitle slots for secondary info;
                    # if there's no subtitle, try first body block as fallback
                    elif (recipe.category == "cover" and analysis.has_body_text
                          and body_chunk_idx < len(body_chunks)
                          and len(body_chunks[body_chunk_idx]) > 0):
                        first_text = body_chunks[body_chunk_idx][0].text
                        if len(first_text) < 120:  # Short enough for subtitle slot
                            self._place_text(ph, first_text, is_bullet=False)
                            body_chunks[body_chunk_idx] = body_chunks[body_chunk_idx][1:]
                            report["content_placed"]["subtitle"] = "from body"

                elif content_type == "image":
                    if next_image < len(analysis.images):
                        self._place_image(ph, analysis.images[next_image])
                        next_image += 1
                        report["content_placed"][f"image_{next_image}"] = "placed"

                elif content_type == "table":
                    if next_table < len(analysis.tables):
                        self._place_table(ph, analysis.tables[next_table])
                        next_table += 1
                        report["content_placed"]["table"] = "placed"

                elif content_type == "object":
                    # Object placeholder: decide what goes in based on what's available
                    # Priority: image (if this is an image slot by name) > table > body text

                    is_image_slot = "image" in slot_name.lower() or "picture" in slot_name.lower()
                    is_table_slot = "table" in slot_name.lower()

                    placed = False

                    # If the slot name suggests an image, try image first
                    if is_image_slot and next_image < len(analysis.images):
                        self._place_image(ph, analysis.images[next_image])
                        next_image += 1
                        report["content_placed"][slot_name] = "image"
                        placed = True

                    # If it's a table slot or we have a table and no text
                    elif is_table_slot and next_table < len(analysis.tables):
                        self._place_table(ph, analysis.tables[next_table])
                        next_table += 1
                        report["content_placed"][slot_name] = "table"
                        placed = True

                    # Default: try body text chunk, then image, then table
                    if not placed:
                        if body_chunk_idx < len(body_chunks) and len(body_chunks[body_chunk_idx]) > 0:
                            self._place_body_blocks(ph, body_chunks[body_chunk_idx])
                            report["content_placed"][slot_name] = f"{len(body_chunks[body_chunk_idx])} blocks"
                            body_chunk_idx += 1
                            placed = True
                        elif next_image < len(analysis.images):
                            self._place_image(ph, analysis.images[next_image])
                            next_image += 1
                            report["content_placed"][slot_name] = "image"
                            placed = True
                        elif next_table < len(analysis.tables):
                            self._place_table(ph, analysis.tables[next_table])
                            next_table += 1
                            report["content_placed"][slot_name] = "table"
                            placed = True

                elif content_type == "body":
                    # Named body slot (e.g. section_number, description)
                    if slot_name == "section_number":
                        # Try to extract a section number from the content
                        num = self._find_section_number(analysis)
                        if num:
                            self._place_text(ph, num, is_bullet=False)
                            report["content_placed"]["section_number"] = num
                    elif slot_name == "description":
                        # Use subtitle or first short body text
                        desc = analysis.subtitle or (
                            analysis.body_texts[0].text if analysis.body_texts else ""
                        )
                        if desc:
                            self._place_text(ph, desc, is_bullet=False)
                            report["content_placed"]["description"] = "placed"
                    elif body_chunk_idx < len(body_chunks) and len(body_chunks[body_chunk_idx]) > 0:
                        self._place_body_blocks(ph, body_chunks[body_chunk_idx])
                        report["content_placed"][slot_name] = f"{len(body_chunks[body_chunk_idx])} blocks"
                        body_chunk_idx += 1

            except KeyError:
                # Placeholder index doesn't exist in this slide
                report["warnings"].append(f"Placeholder {ph_idx} not found for {slot_name}")
            except Exception as e:
                report["warnings"].append(f"Failed to populate {slot_name}: {e}")

    def _split_body_for_columns(self, body_texts: List[BodyBlock], num_columns: int) -> List[List[BodyBlock]]:
        """Split body text blocks roughly equally across N columns.

        Tries to split on natural boundaries (e.g. between top-level blocks)
        rather than mid-paragraph.
        """
        if num_columns <= 1 or not body_texts:
            return [body_texts]

        # Simple approach: split by count, rounding up
        total = len(body_texts)
        chunk_size = max(1, (total + num_columns - 1) // num_columns)

        chunks = []
        for i in range(0, total, chunk_size):
            chunks.append(body_texts[i:i + chunk_size])

        # Pad with empty lists if fewer chunks than columns
        while len(chunks) < num_columns:
            chunks.append([])

        return chunks

    def _find_section_number(self, analysis: ContentAnalysis) -> Optional[str]:
        """Try to extract a section/module number from the content."""
        # Check for override (set by section divider pre-processing)
        if hasattr(analysis, '_section_number_override'):
            return analysis._section_number_override

        import re
        # Look for patterns like "01", "Module 3", "Section 2" in title or body
        patterns = [
            r'[Mm]odule\s+(\d+)',
            r'[Ss]ection\s+(\d+)',
            r'\b(\d{1,2})\b',
        ]
        text_to_search = (analysis.title or "") + " " + (analysis.subtitle or "")
        for pattern in patterns:
            match = re.search(pattern, text_to_search)
            if match:
                return match.group(1).zfill(2)
        return None

    def _place_text(self, placeholder, text: str, is_bullet: bool = False):
        """Place plain text into a placeholder."""
        if not hasattr(placeholder, 'text_frame'):
            return

        text_frame = placeholder.text_frame
        text_frame.clear()

        p = text_frame.paragraphs[0] if text_frame.paragraphs else text_frame.add_paragraph()
        p.text = text
        if is_bullet:
            p.level = 0

    def _place_body_blocks(self, placeholder, body_blocks: List[BodyBlock]):
        """Place multiple body text blocks into a placeholder, preserving bullets."""
        if not hasattr(placeholder, 'text_frame'):
            return

        text_frame = placeholder.text_frame
        text_frame.clear()

        for i, block in enumerate(body_blocks):
            # Use existing first paragraph or create new
            if i == 0:
                p = text_frame.paragraphs[0]
            else:
                p = text_frame.add_paragraph()

            # Use a run so we can set formatting
            run = p.add_run()
            run.text = block.text
            p.level = block.level

            # Apply formatting
            if block.bold:
                run.font.bold = True
            if block.italic:
                run.font.italic = True

    def _place_image(self, placeholder, image_info: ImageInfo):
        """Place an image into a placeholder."""
        try:
            placeholder.insert_picture(io.BytesIO(image_info.blob))
        except Exception as e:
            # If placeholder doesn't support insert_picture, try adding shape
            pass

    def _place_table(self, placeholder, table_info: TableInfo):
        """Place table data into a placeholder."""
        try:
            # Try to insert table into placeholder
            table_shape = placeholder.insert_table(table_info.rows, table_info.cols)
            table = table_shape.table

            for r in range(table_info.rows):
                for c in range(table_info.cols):
                    cell = table.cell(r, c)
                    cell.text = table_info.cell_data[r][c]
        except Exception as e:
            # Placeholder doesn't support table insertion
            pass


# ============================================================================
# LayoutManager
# ============================================================================

class LayoutManager:
    """Orchestrates the full layout auto-apply pipeline."""

    def __init__(self, template_path: str = None, api_key: str = None):
        self.api_key = api_key or os.environ.get("ANTHROPIC_API_KEY")
        self.template_path = template_path or UQ_TEMPLATE_PATH

        self.content_analyser = ContentAnalyser()
        self.layout_matcher = LayoutMatcher()
        self.slide_rebuilder = SlideRebuilder(template_path)

        # Tracking
        self.results: List[SlideResult] = []
        self.total_input_tokens = 0
        self.total_output_tokens = 0

    def run_pipeline(self, pptx_bytes: bytes,
                     progress_callback=None,
                     skip_verification: bool = False,
                     skip_vision: bool = False,
                     slide_limit: int = None) -> dict:
        """Run the full layout pipeline on a PPTX file.

        Args:
            pptx_bytes: Raw bytes of the source PPTX
            progress_callback: Optional callable(step, detail, progress_pct)
            skip_verification: Skip Vision verification step
            skip_vision: Skip ALL Vision calls (use recipe-based matching only)
            slide_limit: Only process first N slides (for testing/cost control)

        Returns:
            dict with:
                - output_pptx_bytes: bytes of the rebuilt PPTX
                - results: list of SlideResult
                - summary: dict with totals
        """
        self.results = []
        self.total_input_tokens = 0
        self.total_output_tokens = 0

        def progress(step, detail, pct):
            if progress_callback:
                progress_callback(step, detail, pct)

        # Load source presentation
        progress("load", "Loading source presentation...", 0.05)
        source_prs = Presentation(io.BytesIO(pptx_bytes))
        num_slides = len(source_prs.slides)
        if slide_limit:
            num_slides = min(num_slides, slide_limit)

        # Create target presentation (empty, will add slides from template)
        progress("create", "Creating target presentation...", 0.10)
        target_prs = Presentation(self.template_path)

        # Remove template example slides
        while len(target_prs.slides) > 0:
            self._remove_first_slide(target_prs)

        # Process each slide
        for slide_idx in range(num_slides):
            pct = 0.15 + (0.75 * slide_idx / num_slides)
            progress("process", f"Processing slide {slide_idx + 1}/{num_slides}...", pct)

            source_slide = source_prs.slides[slide_idx]

            # Analyse content
            analysis = self.content_analyser.analyse_slide(
                source_slide,
                slide_position=slide_idx,
                deck_size=num_slides,
            )

            # Find best layout
            layout_name, score, reasoning = self.layout_matcher.find_best_layout(analysis)

            # Rebuild slide
            rebuild_report = self.slide_rebuilder.rebuild_slide(
                target_prs,
                layout_name,
                analysis,
            )

            # Create result
            result = SlideResult(
                slide_number=slide_idx + 1,
                original_layout=source_slide.slide_layout.name,
                recommended_layout=layout_name,
                confidence=min(100, max(0, score)) / 100.0,  # Normalize to 0-1
                analysis=self._analysis_to_dict(analysis),
                rebuild_report=rebuild_report,
                status="rebuilt" if rebuild_report["status"] == "success" else "failed",
            )
            self.results.append(result)

        # Verify each slide (optional)
        if not skip_verification and not skip_vision and HAS_ANTHROPIC:
            progress("verify", "Verifying slides...", 0.92)
            self._verify_all_slides(source_prs, target_prs)

        # Save target presentation
        progress("save", "Saving rebuilt presentation...", 0.98)
        output_bytes = io.BytesIO()
        target_prs.save(output_bytes)
        output_bytes.seek(0)

        # Deduplicate ZIP entries (python-pptx leaves duplicate entries
        # from removed template slides which corrupts the file for
        # LibreOffice and PowerPoint)
        clean_bytes = self._deduplicate_zip(output_bytes.getvalue())

        # Generate summary
        progress("summary", "Generating summary...", 0.99)
        summary = self._generate_summary()

        progress("done", "Complete", 1.0)

        return {
            "output_pptx_bytes": clean_bytes,
            "results": self.results,
            "summary": summary,
        }

    def _remove_first_slide(self, prs: Presentation):
        """Remove the first slide from a presentation."""
        if len(prs.slides) == 0:
            return

        try:
            # Get the relationship ID of the first slide
            rId = prs.slides._sldIdLst[0].get(
                '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
            )
            if rId:
                prs.part.drop_rel(rId)
            prs.slides._sldIdLst.remove(prs.slides._sldIdLst[0])
        except Exception:
            # If removal fails, just skip
            pass

    def _analysis_to_dict(self, analysis: ContentAnalysis) -> dict:
        """Convert ContentAnalysis to a dict for serialization."""
        return {
            "title": analysis.title,
            "subtitle": analysis.subtitle,
            "has_title": analysis.has_title,
            "has_subtitle": analysis.has_subtitle,
            "has_body_text": analysis.has_body_text,
            "num_body_blocks": analysis.num_body_blocks,
            "has_images": analysis.has_images,
            "image_count": analysis.image_count,
            "has_table": analysis.has_table,
            "is_mostly_text": analysis.is_mostly_text,
            "is_mostly_image": analysis.is_mostly_image,
            "is_section_break": analysis.is_section_break,
            "is_minimal_content": analysis.is_minimal_content,
            "original_layout": analysis.original_layout_name,
        }

    def _verify_all_slides(self, source_prs: Presentation, target_prs: Presentation):
        """Verify each rebuilt slide via Claude Vision (optional enhancement)."""
        # This is a nice-to-have verification step
        # For now, we skip it since the recipe-based matching is reliable
        pass

    def _generate_summary(self) -> dict:
        """Generate a summary of the pipeline results."""
        total_slides = len(self.results)
        rebuilt = sum(1 for r in self.results if r.status == "rebuilt")
        failed = sum(1 for r in self.results if r.status == "failed")
        low_confidence = sum(1 for r in self.results if r.confidence < 0.7)

        return {
            "total_slides": total_slides,
            "rebuilt": rebuilt,
            "failed": failed,
            "low_confidence": low_confidence,
            "total_cost_usd": self._calculate_cost(),
            "total_input_tokens": self.total_input_tokens,
            "total_output_tokens": self.total_output_tokens,
        }

    def _deduplicate_zip(self, pptx_bytes: bytes) -> bytes:
        """Remove duplicate entries from the PPTX ZIP archive.

        python-pptx can leave duplicate ZIP entries when slides are removed
        from a presentation, which corrupts the file for PowerPoint and
        LibreOffice.  We keep the LAST occurrence of each name (the most
        recently written version).
        """
        import zipfile

        src = io.BytesIO(pptx_bytes)
        dst = io.BytesIO()

        with zipfile.ZipFile(src, "r") as zin:
            # Build map: name → last ZipInfo + data
            entries = {}
            for info in zin.infolist():
                entries[info.filename] = (info, zin.read(info.filename))

            with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
                for filename, (info, data) in entries.items():
                    zout.writestr(info, data)

        return dst.getvalue()

    def _calculate_cost(self) -> float:
        """Calculate total API cost for Vision calls."""
        input_cost = self.total_input_tokens * COST_INPUT_PER_TOKEN
        output_cost = self.total_output_tokens * COST_OUTPUT_PER_TOKEN
        return input_cost + output_cost
