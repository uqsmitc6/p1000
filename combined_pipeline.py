#!/usr/bin/env python3
"""
UQ Slide Compliance Tool — Combined Pipeline
===============================================
Runs all three compliance checks in sequence:
  1. Brand Fixer — fonts, colours, tables, bullets, headings
  2. Reference Checker — APA 7 citations, attributions, cross-refs
  3. Image Audit — copyright risk classification

Produces a single fixed PPTX (with brand + ref fixes applied) and a
unified compliance summary report.

Usage:
    python combined_pipeline.py input.pptx --output fixed.pptx
"""

import io
import json
import base64
import tempfile
from collections import defaultdict
from datetime import datetime
from pathlib import Path

from pptx import Presentation

from brand_fixer import BrandFixer
from ref_checker import RefChecker
from image_audit import extract_images, classify_image, generate_html_report


def run_pipeline(pptx_bytes, filename, api_key=None, image_limit=None,
                 skip_image_audit=False, progress_callback=None):
    """Run the full compliance pipeline on a PPTX file.

    Args:
        pptx_bytes: Raw bytes of the input PPTX file
        filename: Original filename (for reporting)
        api_key: Anthropic API key (required for image audit classification)
        image_limit: Limit number of images to classify (0 or None = all)
        skip_image_audit: If True, skip image audit entirely
        progress_callback: Optional callable(percent, message) for progress updates

    Returns:
        dict with keys:
            output_bytes: Fixed PPTX file bytes
            brand_report: Brand fixer report dict
            ref_report: Reference checker report dict
            image_report: Image audit report dict (or None if skipped)
            image_html: Self-contained HTML report string (or None)
            summary: Unified summary dict
    """
    def progress(pct, msg):
        if progress_callback:
            progress_callback(pct, msg)

    # ── Step 1: Brand Fixer ──────────────────────────────────────────
    progress(0, "Step 1/3: Fixing brand formatting...")

    # Load presentation from bytes
    prs = Presentation(io.BytesIO(pptx_bytes))
    fixer = BrandFixer(prs, report=True)

    progress(5, "Fixing fonts...")
    fixer.fix_fonts()
    progress(10, "Fixing text colours...")
    fixer.fix_colours()
    progress(15, "Restyling tables...")
    fixer.fix_tables()
    progress(18, "Standardising footers...")
    fixer.fix_footers()
    progress(20, "Normalising heading sizes...")
    fixer.fix_heading_sizes()
    progress(21, "Checking body text sizes...")
    fixer.flag_body_text_sizes()
    progress(22, "Fixing bullet styles...")
    fixer.fix_bullets()

    brand_report = fixer.generate_report()
    progress(25, f"Brand fixer: {brand_report['total_changes']} changes")

    # ── Step 2: Reference Checker ────────────────────────────────────
    progress(25, "Step 2/3: Checking references & attributions...")

    checker = RefChecker(prs, report=True)
    progress(28, "Scanning citations...")
    checker.scan_citations()
    progress(31, "Scanning reference lists...")
    checker.scan_references()
    progress(34, "Checking image attributions...")
    checker.scan_attributions()
    progress(37, "Cross-referencing...")
    checker.cross_reference()
    progress(39, "Applying reference fixes...")
    checker.fix_attributions()
    checker.fix_citations()

    ref_report = checker.generate_report()
    progress(40, f"Reference checker: {ref_report['total_issues']} issues, {ref_report['total_changes']} fixes")

    # ── Save fixed PPTX ─────────────────────────────────────────────
    output_buffer = io.BytesIO()
    prs.save(output_buffer)
    output_bytes = output_buffer.getvalue()

    # ── Step 3: Image Audit ──────────────────────────────────────────
    image_report = None
    image_html = None
    image_data = None

    if not skip_image_audit:
        progress(40, "Step 3/3: Extracting images...")

        # Extract images from the FIXED file (so we audit what the LDO will actually use)
        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
            tmp.write(output_bytes)
            tmp_path = tmp.name

        try:
            with tempfile.TemporaryDirectory() as extract_dir:
                img_limit = image_limit if image_limit and image_limit > 0 else None
                images = extract_images(tmp_path, output_dir=extract_dir, limit=img_limit)

            if images:
                total_images = len(images)
                progress(45, f"Found {total_images} images")

                if api_key:
                    import anthropic
                    import time
                    client = anthropic.Anthropic(api_key=api_key)

                    classifications = []
                    for i, img_info in enumerate(images):
                        pct = 45 + int((i / total_images) * 45)
                        progress(pct, f"Classifying image {i+1}/{total_images}: {img_info['filename']}")

                        result = classify_image(client, img_info)
                        classifications.append(result)

                        if i < total_images - 1:
                            time.sleep(0.3)

                    # Generate self-contained HTML report
                    image_bytes_map = {
                        img["filename"]: img["image_bytes"]
                        for img in images if "image_bytes" in img
                    }
                    images_clean = [
                        {k: v for k, v in img.items() if k != "image_bytes"}
                        for img in images
                    ]

                    with tempfile.NamedTemporaryFile(
                        suffix=".html", delete=False, mode="w"
                    ) as htmp:
                        html_path = htmp.name

                    generate_html_report(
                        images_clean, classifications,
                        filename, html_path, None,
                        image_bytes_map=image_bytes_map,
                    )
                    image_html = Path(html_path).read_text()

                    # Build summary
                    risk_counts = defaultdict(int)
                    for cls in classifications:
                        if "error" not in cls:
                            risk_counts[cls.get("risk_level", "UNKNOWN")] += 1

                    # Calculate cost
                    from image_audit import COST_PER_M_INPUT, COST_PER_M_OUTPUT
                    total_input_tokens = sum(
                        c.get("api_tokens_used", {}).get("input", 0)
                        for c in classifications
                    )
                    total_output_tokens = sum(
                        c.get("api_tokens_used", {}).get("output", 0)
                        for c in classifications
                    )
                    cost_usd = (
                        (total_input_tokens / 1_000_000) * COST_PER_M_INPUT
                        + (total_output_tokens / 1_000_000) * COST_PER_M_OUTPUT
                    )

                    image_report = {
                        "total_images": total_images,
                        "risk_counts": dict(risk_counts),
                        "classifications": classifications,
                        "images": images_clean,
                        "cost_usd": round(cost_usd, 4),
                        "tokens": {
                            "input": total_input_tokens,
                            "output": total_output_tokens,
                        },
                    }

                    # Build JSON data for download
                    image_data = {
                        "metadata": {
                            "input_file": filename,
                            "total_images": total_images,
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
                else:
                    # No API key — extraction only
                    image_report = {
                        "total_images": total_images,
                        "risk_counts": {},
                        "note": "No API key provided — images extracted but not classified",
                    }
            else:
                image_report = {
                    "total_images": 0,
                    "risk_counts": {},
                    "note": "No images found in presentation",
                }
        finally:
            import os
            try:
                os.unlink(tmp_path)
            except Exception:
                pass

    progress(95, "Building summary...")

    # ── Unified Summary ──────────────────────────────────────────────
    num_slides = len(prs.slides)
    summary = {
        "filename": filename,
        "num_slides": num_slides,
        "generated": datetime.now().isoformat(),
        "brand": {
            "total_changes": brand_report["total_changes"],
            "summary": brand_report["summary"],
        },
        "references": {
            "total_issues": ref_report["total_issues"],
            "total_changes": ref_report["total_changes"],
            "citations_found": ref_report["summary"]["citations_found"],
            "references_found": ref_report["summary"]["references_found"],
        },
        "images": {
            "total_images": image_report["total_images"] if image_report else 0,
            "risk_counts": image_report.get("risk_counts", {}) if image_report else {},
        } if image_report else None,
    }

    progress(100, "Done!")

    return {
        "output_bytes": output_bytes,
        "brand_report": brand_report,
        "ref_report": ref_report,
        "image_report": image_report,
        "image_html": image_html,
        "image_data": image_data,
        "summary": summary,
    }
