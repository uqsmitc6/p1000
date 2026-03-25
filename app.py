#!/usr/bin/env python3
"""
UQ Slide Compliance Tool — Web UI
===================================
Streamlit interface for the brand fixer and image audit tools.
Designed for deployment on Streamlit Community Cloud.

Usage (local):
  streamlit run app.py

Deployment:
  Push to GitHub → connect to share.streamlit.io
  Set ANTHROPIC_API_KEY in Streamlit Cloud secrets
"""

import io
import json
import os
import sys
import tempfile
import time
from pathlib import Path
from collections import defaultdict

import streamlit as st

# Ensure our scripts are importable
sys.path.insert(0, str(Path(__file__).parent))

from pptx import Presentation
from brand_fixer import BrandFixer
from image_audit import extract_images, classify_image, generate_html_report


# ─── API Key from secrets ──────────────────────────────────────────────

def get_api_key():
    """Retrieve API key from Streamlit secrets (set in Cloud dashboard)."""
    try:
        return st.secrets["ANTHROPIC_API_KEY"]
    except (KeyError, FileNotFoundError):
        return os.environ.get("ANTHROPIC_API_KEY")


# ─── Page Config ───────────────────────────────────────────────────────

st.set_page_config(
    page_title="UQ Slide Compliance Tool",
    page_icon="🟣",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─── Custom CSS ────────────────────────────────────────────────────────

st.markdown("""
<style>
    /* UQ Purple branding */
    .stApp > header { background-color: #51247A; }

    .uq-header {
        background: linear-gradient(135deg, #51247A 0%, #962A8B 100%);
        color: white;
        padding: 1.5rem 2rem;
        border-radius: 8px;
        margin-bottom: 1.5rem;
    }
    .uq-header h1 { color: white; margin: 0 0 0.3rem 0; font-size: 1.8rem; }
    .uq-header p { color: rgba(255,255,255,0.85); margin: 0; font-size: 0.95rem; }

    /* Stat cards */
    .stat-card {
        background: white;
        border-radius: 8px;
        padding: 1rem 1.2rem;
        text-align: center;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }
    .stat-card .number { font-size: 1.8rem; font-weight: bold; line-height: 1.2; }
    .stat-card .label { font-size: 0.8rem; color: #666; margin-top: 0.2rem; }

    /* Risk colours */
    .risk-critical { color: #DC2626; }
    .risk-high { color: #EA580C; }
    .risk-medium { color: #D97706; }
    .risk-low { color: #16A34A; }
    .risk-clear { color: #059669; }

    /* Change list */
    .change-item {
        padding: 0.4rem 0.6rem;
        margin: 0.2rem 0;
        border-radius: 4px;
        font-size: 0.85rem;
        border-left: 3px solid #D7D1CC;
    }
    .change-font { border-left-color: #4085C6; background: #F0F7FF; }
    .change-colour { border-left-color: #51247A; background: #F5F0FA; }
    .change-colour_flagged { border-left-color: #D97706; background: #FFFBEB; }
    .change-table { border-left-color: #16A34A; background: #F0FDF4; }
    .change-footer { border-left-color: #962A8B; background: #FDF0FA; }
    .change-heading_size { border-left-color: #E62645; background: #FFF0F3; }
    .change-bullet { border-left-color: #FBB800; background: #FFFDE7; }

    /* Hide deploy button and Streamlit branding */
    .stDeployButton { display: none; }
    #MainMenu { visibility: hidden; }
    footer { visibility: hidden; }

    /* Info box */
    .info-box {
        background: #F5F0FA;
        border: 1px solid #D7D1CC;
        border-radius: 8px;
        padding: 1rem 1.2rem;
        margin: 1rem 0;
        font-size: 0.9rem;
    }
    .info-box strong { color: #51247A; }
</style>
""", unsafe_allow_html=True)


# ─── Header ────────────────────────────────────────────────────────────

st.markdown("""
<div class="uq-header">
    <h1>UQ Slide Compliance Tool</h1>
    <p>Brand formatting fixer &amp; image copyright audit for executive education decks</p>
</div>
""", unsafe_allow_html=True)


# ─── Sidebar (minimal — just info) ────────────────────────────────────

with st.sidebar:
    st.markdown("### How to use")
    st.markdown(
        "**Brand Fixer** — Upload a `.pptx`, click the button, "
        "download your fixed file. No API key needed."
    )
    st.markdown(
        "**Image Audit** — Upload a `.pptx` and the tool will extract "
        "every image and classify its copyright risk using AI."
    )
    st.markdown("---")
    st.markdown(
        "**Tips:**\n"
        "- Start with 'Extract only' to preview images before running the full audit\n"
        "- For large decks, set a limit (e.g. 10) to test before running all images\n"
        "- The brand fixer flags uncertain colours for you to review manually"
    )
    st.markdown("---")

    # API status indicator
    api_key = get_api_key()
    if api_key:
        st.success("API key configured", icon="✅")
    else:
        st.warning("API key not configured — image audit classification unavailable", icon="⚠️")

    st.markdown(
        "<small style='color: #999;'>Built for UQ Business School<br/>"
        "Learning Design Team</small>",
        unsafe_allow_html=True,
    )


# ─── Main Tabs ─────────────────────────────────────────────────────────

tab1, tab2 = st.tabs(["Brand Fixer", "Image Audit"])


# ═══════════════════════════════════════════════════════════════════════
# TAB 1: Brand Fixer
# ═══════════════════════════════════════════════════════════════════════

with tab1:
    st.markdown("#### Upload a slide deck to fix brand formatting")

    st.markdown("""
    <div class="info-box">
        <strong>What this does:</strong> Corrects fonts to Arial, fixes text colours
        to the UQ palette, restyles tables with UQ brand colours, normalises heading
        sizes, and standardises bullets. Your original file is never modified —
        you'll download a new fixed copy.
    </div>
    """, unsafe_allow_html=True)

    uploaded_file = st.file_uploader(
        "Choose a .pptx file",
        type=["pptx"],
        key="brand_upload",
        help="Upload the slide deck you want to fix",
    )

    if uploaded_file is not None:
        # Show file info
        file_size_mb = len(uploaded_file.getvalue()) / (1024 * 1024)
        st.caption(f"Uploaded: **{uploaded_file.name}** ({file_size_mb:.1f} MB)")

        if st.button("Fix Brand Formatting", type="primary", key="run_brand"):
            # Save uploaded file to temp
            with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
                tmp.write(uploaded_file.getvalue())
                tmp_path = tmp.name

            try:
                # Load and fix
                prs = Presentation(tmp_path)
                fixer = BrandFixer(prs, report=True)

                # Progress updates
                progress = st.progress(0, text="Fixing fonts...")
                fixer.fix_fonts()
                progress.progress(17, text="Fixing text colours...")
                fixer.fix_colours()
                progress.progress(34, text="Restyling tables...")
                fixer.fix_tables()
                progress.progress(50, text="Standardising footers...")
                fixer.fix_footers()
                progress.progress(67, text="Normalising heading sizes...")
                fixer.fix_heading_sizes()
                progress.progress(84, text="Fixing bullet styles...")
                fixer.fix_bullets()
                progress.progress(100, text="Done!")

                # Save fixed file to buffer
                output_buffer = io.BytesIO()
                prs.save(output_buffer)
                output_buffer.seek(0)

                report = fixer.generate_report()
                total = report["total_changes"]

                # ── Summary stats ──
                st.markdown("---")
                st.markdown("### Results")

                if total == 0:
                    st.success(
                        "No changes needed — this deck is already brand-compliant!",
                        icon="✅",
                    )
                else:
                    stats = report["summary"]
                    cols = st.columns(7)
                    stat_items = [
                        ("Font fixes", stats.get("font", 0), "#4085C6"),
                        ("Colour fixes", stats.get("colour", 0), "#51247A"),
                        ("Flagged", stats.get("colour_flagged", 0), "#D97706"),
                        ("Tables", stats.get("table", 0), "#16A34A"),
                        ("Footers", stats.get("footer", 0), "#962A8B"),
                        ("Headings", stats.get("heading_size", 0), "#E62645"),
                        ("Bullets", stats.get("bullet", 0), "#FBB800"),
                    ]
                    for col, (label, count, colour) in zip(cols, stat_items):
                        col.markdown(
                            f"<div class='stat-card'>"
                            f"<div class='number' style='color:{colour};'>{count}</div>"
                            f"<div class='label'>{label}</div></div>",
                            unsafe_allow_html=True,
                        )

                    st.markdown(f"**{total} total changes** across {len(prs.slides)} slides")

                    # ── Download button ──
                    fixed_name = uploaded_file.name.replace(".pptx", "_FIXED.pptx")
                    st.download_button(
                        label=f"Download {fixed_name}",
                        data=output_buffer,
                        file_name=fixed_name,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        type="primary",
                    )

                    # ── Flagged items (show first — these need attention) ──
                    flagged = [
                        c for c in report["changes"]
                        if c["category"] == "colour_flagged"
                    ]
                    if flagged:
                        with st.expander(
                            f"⚠️  {len(flagged)} colours flagged for review",
                            expanded=True,
                        ):
                            st.markdown(
                                "These non-UQ colours were **not auto-corrected** "
                                "because they may be intentional accent colours. "
                                "Please check these slides and decide whether to "
                                "keep or change them:"
                            )
                            # Deduplicate by colour + slide for cleaner display
                            seen = set()
                            for c in flagged:
                                key = (c["slide"], c["detail"])
                                if key not in seen:
                                    seen.add(key)
                                    st.markdown(
                                        f"<div class='change-item change-colour_flagged'>"
                                        f"Slide {c['slide']}: {c['detail']}</div>",
                                        unsafe_allow_html=True,
                                    )

                    # ── Detailed changes (collapsed) ──
                    if report["changes"]:
                        with st.expander(
                            f"View all {total} changes", expanded=False
                        ):
                            by_slide = defaultdict(list)
                            for change in report["changes"]:
                                by_slide[change["slide"]].append(change)

                            for slide_num in sorted(by_slide.keys()):
                                changes = by_slide[slide_num]
                                st.markdown(
                                    f"**Slide {slide_num}** "
                                    f"({len(changes)} change{'s' if len(changes) != 1 else ''})"
                                )
                                for c in changes:
                                    cat = c["category"]
                                    st.markdown(
                                        f"<div class='change-item change-{cat}'>"
                                        f"{c['detail']}</div>",
                                        unsafe_allow_html=True,
                                    )

            except Exception as e:
                st.error(f"Something went wrong: {e}")
                with st.expander("Error details"):
                    import traceback
                    st.code(traceback.format_exc())
            finally:
                try:
                    os.unlink(tmp_path)
                except Exception:
                    pass


# ═══════════════════════════════════════════════════════════════════════
# TAB 2: Image Audit
# ═══════════════════════════════════════════════════════════════════════

with tab2:
    st.markdown("#### Upload a slide deck to audit images for copyright risk")

    st.markdown("""
    <div class="info-box">
        <strong>What this does:</strong> Extracts every image from the deck and
        uses AI to classify each one by type (stock photo, screenshot, diagram, etc.)
        and copyright risk level (Critical → Clear). You'll get a downloadable
        report showing which images need attention.<br/><br/>
        <strong>Why it matters:</strong> Executive education programs are commercial —
        they fall outside the statutory education licence. Every uncleared copyrighted
        image is genuine legal exposure.
    </div>
    """, unsafe_allow_html=True)

    audit_file = st.file_uploader(
        "Choose a .pptx file",
        type=["pptx"],
        key="audit_upload",
        help="Upload the slide deck you want to audit",
    )

    if audit_file is not None:
        file_size_mb = len(audit_file.getvalue()) / (1024 * 1024)
        st.caption(f"Uploaded: **{audit_file.name}** ({file_size_mb:.1f} MB)")

        col_a, col_b = st.columns(2)
        with col_a:
            limit = st.number_input(
                "Limit images (0 = all)",
                min_value=0,
                max_value=500,
                value=0,
                help="Limit to first N images — useful for testing. Set to 0 to audit all.",
            )
        with col_b:
            extract_only = st.checkbox(
                "Extract only (no AI classification)",
                help="Just see what images are in the deck without running the AI. Free and fast.",
            )

        if st.button("Run Image Audit", type="primary", key="run_audit"):
            # Validate API key if doing classification
            if not extract_only:
                api_key = get_api_key()
                if not api_key:
                    st.error(
                        "The API key hasn't been configured yet. "
                        "Ask Sean to set it up, or tick 'Extract only' to preview images without AI."
                    )
                    st.stop()

            with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
                tmp.write(audit_file.getvalue())
                tmp_path = tmp.name

            try:
                # ── Extract ──
                with st.spinner("Extracting images from presentation..."):
                    with tempfile.TemporaryDirectory() as extract_dir:
                        img_limit = limit if limit > 0 else None
                        images = extract_images(
                            tmp_path, output_dir=extract_dir, limit=img_limit
                        )

                if not images:
                    st.info("No images found in this presentation.")
                    st.stop()

                st.success(f"Found **{len(images)}** unique images across {len(set(i['slide_number'] for i in images))} slides")

                # ── Extract-only mode ──
                if extract_only:
                    st.markdown("---")
                    st.markdown("### Extracted Images")
                    st.caption("Showing all images found in the deck. Run with AI classification to get risk ratings.")

                    for img_info in images:
                        col1, col2 = st.columns([1, 3])
                        with col1:
                            try:
                                from PIL import Image as PILImage
                                pil_img = PILImage.open(io.BytesIO(img_info["image_bytes"]))
                                st.image(pil_img, width=180)
                            except Exception:
                                st.caption("*(preview unavailable)*")
                        with col2:
                            st.markdown(
                                f"**Slide {img_info['slide_number']}** — "
                                f"`{img_info['filename']}`"
                            )
                            st.caption(
                                f"{img_info['width']}×{img_info['height']}px "
                                f"  |  {img_info['content_type']}"
                            )
                            if img_info.get("slide_title"):
                                st.caption(f"Slide title: {img_info['slide_title']}")
                        st.markdown("---")

                # ── Full AI classification ──
                else:
                    st.markdown("---")

                    import anthropic
                    client = anthropic.Anthropic(api_key=api_key)

                    classifications = []
                    progress = st.progress(0)
                    status_text = st.empty()

                    for i, img_info in enumerate(images):
                        status_text.markdown(
                            f"Classifying image **{i+1}/{len(images)}**: "
                            f"`{img_info['filename']}` (slide {img_info['slide_number']})"
                        )
                        progress.progress(i / len(images))

                        result = classify_image(client, img_info)
                        classifications.append(result)

                        # Brief pause for rate limiting
                        if i < len(images) - 1:
                            time.sleep(0.3)

                    progress.progress(1.0)
                    status_text.empty()
                    st.success("Classification complete!")

                    # ── Risk Summary ──
                    risk_counts = defaultdict(int)
                    for cls in classifications:
                        if "error" not in cls:
                            risk_counts[cls.get("risk_level", "UNKNOWN")] += 1

                    st.markdown("### Risk Summary")

                    risk_items = [
                        ("Critical", risk_counts.get("CRITICAL", 0), "critical", "#DC2626"),
                        ("High", risk_counts.get("HIGH", 0), "high", "#EA580C"),
                        ("Medium", risk_counts.get("MEDIUM", 0), "medium", "#D97706"),
                        ("Low", risk_counts.get("LOW", 0), "low", "#16A34A"),
                        ("Clear", risk_counts.get("CLEAR", 0), "clear", "#059669"),
                    ]

                    cols = st.columns(5)
                    for col, (label, count, cls_name, colour) in zip(cols, risk_items):
                        col.markdown(
                            f"<div class='stat-card'>"
                            f"<div class='number' style='color:{colour};'>{count}</div>"
                            f"<div class='label'>{label}</div></div>",
                            unsafe_allow_html=True,
                        )

                    # Key message based on results
                    critical_high = risk_counts.get("CRITICAL", 0) + risk_counts.get("HIGH", 0)
                    if critical_high > 0:
                        st.warning(
                            f"**{critical_high} image{'s' if critical_high != 1 else ''} "
                            f"flagged as Critical or High risk** — "
                            f"these likely need replacement or licence verification.",
                            icon="⚠️",
                        )

                    # ── Downloads ──
                    st.markdown("---")
                    st.markdown("### Download Report")

                    dl_col1, dl_col2 = st.columns(2)

                    # Generate HTML report
                    with tempfile.TemporaryDirectory() as report_dir:
                        for img_info in images:
                            img_path = Path(report_dir) / "images" / img_info["filename"]
                            img_path.parent.mkdir(exist_ok=True)
                            img_path.write_bytes(img_info["image_bytes"])

                        report_path = Path(report_dir) / "report.html"

                        images_clean = []
                        for img in images:
                            images_clean.append(
                                {k: v for k, v in img.items() if k != "image_bytes"}
                            )

                        summary = generate_html_report(
                            images_clean, classifications,
                            audit_file.name, str(report_path), "images",
                        )

                        html_content = report_path.read_text()

                    with dl_col1:
                        report_name = audit_file.name.replace(".pptx", "_audit_report.html")
                        st.download_button(
                            label="Download HTML Report",
                            data=html_content,
                            file_name=report_name,
                            mime="text/html",
                            type="primary",
                        )

                    with dl_col2:
                        json_data = {
                            "metadata": {
                                "input_file": audit_file.name,
                                "total_images": len(images),
                            },
                            "summary": {"risk_counts": dict(risk_counts)},
                            "images": [
                                {
                                    **{k: v for k, v in img.items() if k != "image_bytes"},
                                    **cls,
                                }
                                for img, cls in zip(images, classifications)
                            ],
                        }
                        json_name = audit_file.name.replace(".pptx", "_audit_data.json")
                        st.download_button(
                            label="Download JSON Data",
                            data=json.dumps(json_data, indent=2, default=str),
                            file_name=json_name,
                            mime="application/json",
                        )

                    # ── Image cards ──
                    st.markdown("---")
                    st.markdown("### Image Details")
                    st.caption("Sorted by risk level (highest first)")

                    risk_order = {
                        "CRITICAL": 0, "HIGH": 1, "MEDIUM": 2,
                        "LOW": 3, "CLEAR": 4, "UNKNOWN": 5,
                    }
                    paired = sorted(
                        zip(images, classifications),
                        key=lambda x: risk_order.get(
                            x[1].get("risk_level", "UNKNOWN"), 5
                        ),
                    )

                    # Risk filter
                    show_levels = st.multiselect(
                        "Filter by risk level:",
                        ["CRITICAL", "HIGH", "MEDIUM", "LOW", "CLEAR"],
                        default=["CRITICAL", "HIGH", "MEDIUM", "LOW", "CLEAR"],
                    )

                    for img_info, cls in paired:
                        risk_level = cls.get("risk_level", "UNKNOWN")
                        if risk_level not in show_levels:
                            continue

                        risk_css = risk_level.lower() if risk_level in risk_order else "unknown"

                        col1, col2 = st.columns([1, 3])
                        with col1:
                            try:
                                from PIL import Image as PILImage
                                pil_img = PILImage.open(
                                    io.BytesIO(img_info["image_bytes"])
                                )
                                st.image(pil_img, width=200)
                            except Exception:
                                st.caption("*(preview unavailable)*")

                        with col2:
                            risk_display = cls.get("risk_level", "ERROR")
                            img_type = cls.get("image_type", "ERROR")
                            action = cls.get("recommended_action", "N/A")

                            # Header line
                            risk_colour = {
                                "CRITICAL": "#DC2626", "HIGH": "#EA580C",
                                "MEDIUM": "#D97706", "LOW": "#16A34A",
                                "CLEAR": "#059669",
                            }.get(risk_display, "#6B7280")

                            st.markdown(
                                f"**Slide {img_info['slide_number']}** &nbsp; "
                                f"<span style='background:{risk_colour};color:white;"
                                f"padding:2px 8px;border-radius:4px;font-size:0.8rem;'>"
                                f"{risk_display}</span> &nbsp; "
                                f"<span style='background:#E5E7EB;padding:2px 8px;"
                                f"border-radius:4px;font-size:0.8rem;'>{img_type}</span>",
                                unsafe_allow_html=True,
                            )

                            if "error" in cls:
                                st.error(cls["error"])
                            else:
                                st.markdown(cls.get("reasoning", ""))

                                if cls.get("content_description"):
                                    st.caption(f"Content: {cls['content_description']}")

                                st.caption(f"Recommended action: **{action}**")

                                flags = []
                                if cls.get("watermark_text"):
                                    flags.append(f"Watermark: {cls['watermark_text']}")
                                if cls.get("copyright_notice"):
                                    flags.append(f"Copyright: {cls['copyright_notice']}")
                                if cls.get("brand_visible"):
                                    flags.append(f"Brand: {cls['brand_visible']}")
                                if flags:
                                    st.warning(" | ".join(flags))

                        st.markdown("---")

            except Exception as e:
                st.error(f"Something went wrong: {e}")
                with st.expander("Error details"):
                    import traceback
                    st.code(traceback.format_exc())
            finally:
                try:
                    os.unlink(tmp_path)
                except Exception:
                    pass
