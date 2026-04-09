# Project Log — UQ Slide Compliance Tool

> **Last updated:** 2026-04-09
> **Project:** Two-part automated tool for UQ Business School executive education slide deck compliance — brand formatting fixer + copyright image audit
> **Tech stack:** Python (python-pptx, Pillow, lxml), Claude Vision API (anthropic SDK), Streamlit UI, HTML reporting
> **Repository:** Cowork project folder
> **Live app:** https://p10004u.streamlit.app/

---

## Current State

All four components are **built and functionally tested**, plus a **Streamlit web UI** (`app.py`) wraps all tools for non-technical users — now with four tabs: Brand Fixer, Image Audit, Reference Checker, and Full Compliance Check. App is **live on Streamlit Community Cloud** with API key configured. The combined pipeline runs all three checks in sequence from a single upload. Self-contained HTML image audit reports now embed images as base64 data URIs.

1. ~~**Too many "review manually" flags**~~ — FIXED. Expanded near-black colour detection (explicit hex list + luminance/saturation check). Flags reduced from 55 → 27, with remaining flags being genuine accent colours (greens, oranges, reds, blues) that do need manual review.
2. ~~**Download button cleared the page**~~ — FIXED. Results now stored in `st.session_state` so they persist across Streamlit reruns triggered by download buttons.
3. **Slide backgrounds/masters not applied** — OPEN. The brand fixer only modifies text-level properties (fonts, colours, bullets). It does not change slide backgrounds, layouts, or masters. This is architecturally significant — see Open Items for discussion.

### Key Architecture Decisions in Effect

- "Auto-correct with smart rules" for colour handling — auto-fix obvious cases (e.g. near-black→#2B1D37), skip text on dark backgrounds, flag non-UQ accent colours for manual review — decided 2026-03-25
- Both Step 1 (brand fixer) and Step 2 (image audit) developed in same session
- python-pptx is the core library for PPTX manipulation (not XML unpacking) — per project instructions
- Never modify the original file — always output a new file
- Err on the side of flagging — false positives preferred over false negatives
- Theme font codes (+mj-lt, +mn-lt) are left alone — they resolve to Arial via the theme and don't need fixing — decided 2026-03-25
- Table restyling: first row always treated as header (UQ Purple bg, white text), alternating rows white/#D7D1CC — decided 2026-03-25
- Bullet normalisation: only non-standard characters are changed; standard bullets (•, –, ▪, ►, etc.) are left as-is — decided 2026-03-25
- Image audit uses Claude claude-sonnet-4-20250514 for classification — balances cost vs accuracy — decided 2026-03-25

---

## Open Items

| # | Category | Description | Since | Priority | Notes |
|---|----------|-------------|-------|----------|-------|
| 1 | DONE | ~~Test Step 2 AI classification with live API key~~ | 2026-03-25 | — | Tested via Streamlit Cloud on BEFORE deck. 67 images classified. Cost: USD $0.51. Working correctly. |
| 2 | DECIDED | ~~Slide backgrounds/masters — scope and approach~~ | 2026-03-25 | — | **Decision: Leave as out-of-scope for automated tool.** Confirmed that placeholder indices match between BEFORE deck and UQ template — when LDO manually selects the correct layout, text snaps into position correctly. Manual layout selection is a small price vs the risk of automated layout remapping. |
| 3 | DONE | ~~Refine image classifier for Adobe Stock awareness~~ | 2026-03-25 | — | Classification prompt updated with licensing context (Adobe Stock, Shutterstock, Microsoft Stock). Added `_detect_attribution()` function to scan slide text for attribution patterns and pass results to the classifier. New fields: `attribution_found` in response, `ADD_ATTRIBUTION` as recommended action. Risk guide updated: stock photos with attribution → LOW, without → VERIFY_LICENCE. |
| 4 | TODO | Build image replacement tool (future component) | 2026-03-25 | Medium | The image audit identifies problems; a future component should help resolve them. Scope TBD — could range from suggesting free alternatives to auto-replacing with Adobe Stock search. Sean has flagged this to Evie and Sarah as a planned component. |
| 5 | DONE | ~~Tune classification prompt based on first real run~~ | 2026-03-25 | — | Prompt significantly expanded in Session 3: added licensing context, attribution detection, new `ADD_ATTRIBUTION` action, refined risk guide to distinguish licensed vs unlicensed stock photos. Should substantially reduce false high/critical flags on Adobe Stock images. Needs live re-test to confirm improvement. |
| 6 | NOTE | BEFORE deck is available as PPTX (148 slides) not just PDF — original instructions said PDF only but PPTX was uploaded | 2026-03-25 | Low | Can extract images directly from PPTX |
| 7 | NOTE | Slide 12 in BEFORE deck has a linked/external image (not embedded) — script handles this gracefully with a warning | 2026-03-25 | Low | The shape is "Graphic 56" — likely an SVG or EMF |
| 8 | DONE | ~~Add `--footer-text` CLI option to brand_fixer.py~~ | 2026-03-25 | — | Added `--footer-text` CLI arg and `footer_text` param to BrandFixer constructor. Sets all FOOTER placeholder text to specified string. Tested: 146 footers updated on AFTER deck with "UQ Business School". |
| 9 | DONE | ~~Add body text size check~~ | 2026-03-25 | — | Added `flag_body_text_sizes()` method. Flags body text outside 12–24pt range (flag-only, no auto-fix). Excludes titles, footers, and attribution text (detected via regex). 45 flags on BEFORE deck (8–10pt text). Added as step 6 in `fix_all()`, integrated into Streamlit UI and combined pipeline. |
| 10 | NOTE | Image attribution convention in AFTER deck | 2026-03-25 | — | LDOs use on-slide text boxes with format "Source: Adobe Stock {ID}" or "Image licensed through Adobe Stock: {ID}". Also: "Source: {creator}, CC BY-SA 4.0, via Wikimedia Commons", "Source: Public domain, via Wikimedia Commons", "Images: Microsoft Stock", "Image licensed through Shutterstock: {ID}". Some slides also note attribution in speaker notes (slides 127, 144). |
| 11 | TODO | Build layout auto-apply tool | 2026-04-09 | High | Analyse each slide's content (title, body, images, tables) and automatically assign the best-matching UQ template layout. Addresses feedback F1 (cover/subsection formatting) and F2 (text alignment). Sean chose analyse-and-auto-apply approach. Requires: unpacking all UQ template layouts, building content-to-layout matching rules per slide type, handling edge cases (multi-image slides, tables, etc.). This is the next major feature. |
| 12 | TODO | Explore Streamlit upload size limit | 2026-04-09 | Low | Evie/Sarah asked about increasing the upload limit. Current Streamlit config may restrict large decks. Check .streamlit/config.toml settings. |

---

## Session History

### Session 4 — 2026-04-09

**Focus:** Triage Evie & Sarah's testing feedback, implement quick fixes, plan layout auto-apply tool.

**Feedback received (Slide Deck Automation Test Document.docx):**

Issues triaged and categorised:
| # | Feedback Item | Status | Action |
|---|---------------|--------|--------|
| F1 | Cover/subsection slides don't format properly after layout change | PLANNED | Part of layout auto-apply tool (see Open Item #11) |
| F2 | Text alignment on blank branded slides | PLANNED | Part of layout auto-apply tool |
| F3 | Title colours reset to black after layout change | FIXED v1.4 | Title colours now explicitly set to UQ Purple (#51247A) even when theme-inherited |
| F4 | Table header text black on purple background (unreadable) | FIXED v1.4 | Table header text now explicitly set to white, even when theme-inherited |
| F5 | Slide master currency check | DEFERRED | Sean: "bridge too far for now — if they update the official slide deck I can update the tool" |
| F6 | File upload size limit | NOTE | Streamlit Cloud default 200MB; config.toml can increase |
| F7 | Data security / IP concerns | NOTE | Sean: "no" — Streamlit Cloud, API calls to Anthropic only |
| F8 | Prioritise diagrams/figures over decorative images | FIXED v1.4 | Classification prompt updated with priority guide |
| F9 | Gov/corporate website screenshots generally OK | FIXED v1.4 | Prompt updated: gov/corporate screenshots LOW if acknowledged, MEDIUM if not |
| F10 | Auto-format references/attributions | ALREADY DONE | Reference Checker tab (Session 2b) — Evie/Sarah may not have seen it |
| F11 | Flag missing image attributions | ALREADY DONE | Reference Checker tab flags these |
| F12 | Generate alt text for images | FIXED v1.4 | Added `alt_text` field to classification prompt; shown in HTML report and Streamlit UI |

**Outcomes:**
- **Title colour fix**: Titles with theme-inherited colour now explicitly set to UQ Purple. Previously skipped (colour was None), causing titles to appear black after LDO applies new layout. 133 titles fixed on BEFORE deck.
- **Table header text fix**: Header row text now explicitly set to white even when theme-inherited. Previously purple background was applied but text stayed inherited (black), making it unreadable. 9 header text runs fixed on BEFORE deck.
- **Image audit prompt improvements**: Added priority guide (published diagrams > decorative photos), government/corporate screenshot handling, alt text generation field.
- **Alt text generation**: Classification prompt now generates suggested alt text (max 125 chars) for each image. Shown in both HTML report and Streamlit image cards. No extra API cost — same call.
- **Version number**: Added `APP_VERSION` to app.py header. v1.3.0 → v1.4.0.
- **Live app URL**: Recorded as https://p10004u.streamlit.app/

**Key decision: Layout auto-apply tool**
Sean wants to build a tool that analyses each slide's content and automatically assigns the best-matching UQ template layout. This addresses feedback items F1 and F2. Approach: analyse → auto-apply (not just recommend). This is the next major feature — see Open Item #11.

**Changes to codebase:**
- `brand_fixer.py`: Title colours explicitly set for theme-inherited text; table header text explicitly set to white; both fixes prevent colour issues after layout changes
- `image_audit.py`: Classification prompt updated with priority guide, gov/corporate screenshot rules, and alt text generation field; HTML report shows alt text
- `app.py`: Version bumped to 1.4.0; version number shown in header; alt text shown in image cards
- `PROJECT_LOG.md`: Full feedback triage, Session 4 history, new open item #11

---

### Session 3 — 2026-03-26

**Focus:** Tick off remaining open items: Adobe Stock classifier refinement, footer text, body text size check, combined pipeline testing.

**Outcomes:**
- **Image classifier prompt overhaul (#3 + #5):**
  - Added licensing context to the classification prompt: UQ holds Adobe Stock, Shutterstock, and Microsoft 365 licences
  - Created `_detect_attribution()` function with 17 regex patterns to scan slide text for attribution mentions before sending to the API
  - Added `detected_attribution` field to the context template so the classifier sees attribution text alongside the image
  - Added new response field `attribution_found` and new recommended action `ADD_ATTRIBUTION`
  - Updated risk guide: stock photos with attribution → LOW, without → VERIFY_LICENCE (not HIGH/CRITICAL)
  - Tested on AFTER deck: 3 out of first 20 images had attribution correctly detected
- **Footer text feature (#8):**
  - Added `footer_text` parameter to `BrandFixer.__init__()` and `--footer-text` CLI arg
  - Only sets text on FOOTER-type placeholders (not slide numbers/dates)
  - Tested: 146 footers updated on AFTER deck
- **Body text size check (#9):**
  - Added `flag_body_text_sizes()` method — flags body text outside 12–24pt range
  - Flag-only (no auto-correction) since small text may be intentional
  - Excludes titles, footers, and attribution text (via regex filter for "source:", "adobe stock", "cc by", etc.)
  - 45 flags on BEFORE deck (8–10pt text), 78 with original 14pt threshold → lowered to 12pt for less noise
  - Integrated into `fix_all()`, Streamlit UI (new stat card + expandable section), and combined pipeline
- **Combined pipeline tested on BEFORE deck:**
  - 148 slides: 246 brand changes (including 45 body size flags), 73 ref issues, 10 ref fixes — PASS

**Changes to codebase:**
- `image_audit.py`: Overhauled classification prompt, added `_detect_attribution()`, updated context template with `detected_attribution` field
- `brand_fixer.py`: Added `footer_text` param + CLI arg, added `flag_body_text_sizes()` method, added `import re`, updated `fix_all()` and `print_summary()` categories
- `app.py`: Added body size CSS class, updated Brand Fixer stat cards (7→8 columns), added body size flagged expandable section, updated progress steps
- `combined_pipeline.py`: Added `flag_body_text_sizes()` call to pipeline
- `PROJECT_LOG.md`: Updated open items #3, #5, #8, #9 to DONE, added Session 3 history

---

### Session 2c — 2026-03-25/26

**Focus:** Two overnight features — combined pipeline mode and self-contained HTML image audit reports.

**Outcomes:**
- Built `combined_pipeline.py` — orchestrates brand fixer → reference checker → image audit in a single `run_pipeline()` call:
  - Takes raw PPTX bytes, returns fixed PPTX bytes + all three reports + unified summary
  - Applies brand fixes and ref fixes to the same Presentation object before saving
  - Extracts images from the FIXED file (not original) so audit reflects what LDO will actually use
  - Supports `skip_image_audit` flag and `image_limit` for cost-controlled runs
  - Progress callback for real-time UI updates
- Added "Full Compliance Check" as fourth tab in Streamlit app:
  - Single file upload → runs all three checks in sequence
  - Downloads: fixed PPTX (_COMPLIANT.pptx), JSON compliance report, self-contained HTML image audit
  - Collapsible detail sections for each component (brand, refs, images)
  - Overview stat cards showing slides, brand fixes, ref fixes, images audited
- Made HTML image audit reports self-contained:
  - Modified `generate_html_report()` in `image_audit.py` to accept optional `image_bytes_map` parameter
  - When provided, images are embedded as base64 data URIs instead of file path references
  - HTML report can be downloaded and opened anywhere — no separate image folder needed
  - Updated both the standalone Image Audit tab and the combined pipeline to use this approach
- All features tested end-to-end:
  - Self-contained HTML: 3 test images embedded as data URIs, no file references (PASS)
  - Combined pipeline (skip images): 141 brand changes, 60 ref issues, 12 fixes on AFTER deck (PASS)
  - Combined pipeline (extract only, no API): Images extracted but not classified (PASS)

**Changes to codebase:**
- `combined_pipeline.py`: Created — ~250 lines, orchestrates all three tools
- `app.py`: Added fourth tab "Full Compliance Check", added `from combined_pipeline import run_pipeline`, updated sidebar docs
- `image_audit.py`: Modified `generate_html_report()` to support `image_bytes_map` for base64 embedding
- `PROJECT_LOG.md`: Updated with Session 2c outcomes

---

### Session 2b — 2026-03-25

**Focus:** Build APA 7 Reference & Attribution Checker (new component).

**Outcomes:**
- Built `ref_checker.py` — comprehensive reference and attribution checker with auto-fix capability:
  - **In-text citation detection**: Parenthetical `(Author, Year)`, narrative `Author (Year)`, and `Adapted from Author (Year)` patterns. Validates APA 7 format (& vs and, et al. period, comma placement).
  - **Reference list detection**: Identifies reference slides by title, extracts individual entries, validates APA 7 format (DOI format, "Retrieved from" deprecation, publisher location, edition format).
  - **Image attribution detection**: Recognises Adobe Stock, Shutterstock, Microsoft Stock, Wikimedia Commons, public domain, Flickr, and generic "Source:" attributions. Standardises all to "Source: {Provider} {ID}" format.
  - **Cross-referencing**: Matches citations against reference list entries by surname+year. Flags orphaned citations (cited but not in refs) and orphaned references (in list but never cited). Handles multi-author matching and partial surname matching.
  - **Missing attribution detection**: Flags slides with images but no attribution text. Filters out likely decorative images (icons, logos, small images) to reduce noise.
  - **Auto-fix**: Modifies PPTX to standardise attributions and fix citation formatting. Produces change report alongside issues report.
- Added "Reference Checker" as third tab in Streamlit app with full UI (progress bar, stat cards, categorised expandable issue lists, download buttons for fixed file + JSON report).
- Tested on both decks:
  - AFTER deck: 19 citations found, 27 references across 4 ref slides, 60 issues (31 warnings, 29 info), 12 auto-fixes applied
  - BEFORE deck: 10 citations found, 0 parseable references (ref slide has URLs not APA format), 73 issues (64 warnings), 55 missing attributions

**Decisions made:**
- Standardise all Adobe Stock attributions to "Source: Adobe Stock {ID}" format (Sean's explicit choice)
- Flag slides with images but no attribution, filtering out decorative/small images
- Cross-referencing: orphaned citations are warnings, orphaned references are info (many refs may be "further reading")
- Concatenated text (e.g., reference + attribution jammed together) flagged for manual review rather than auto-fixed

**Changes to codebase:**
- `ref_checker.py`: Created — ~650 lines, APA 7 checking, attribution standardisation, cross-referencing, auto-fix
- `app.py`: Added third tab "Reference Checker", import of RefChecker, new CSS classes for issue categories
- `PROJECT_LOG.md`: Updated with Session 2b outcomes

---

### Session 2 — 2026-03-25

**Focus:** Fix three issues identified from live testing before sharing with Evie and Sarah.

**Outcomes:**
- **Near-black colour threshold** — Expanded detection from a simple Euclidean distance check (`< 60` from black) to a three-pronged approach: explicit hex list of 20 common near-black values (#111111 through #545454), the original distance check, and a luminance+saturation heuristic (max channel < 100, saturation < 30). Result: flagged colours dropped from 55 → 27. All 27 remaining flags are genuine accent colours (greens, oranges, reds, blues) that do need human review.
- **Download clearing the page** — Refactored both Brand Fixer and Image Audit tabs to store results in `st.session_state`. Processing happens inside the button click, results are rendered from session state outside the click block. Download buttons now work without clearing the displayed results.
- **Slide backgrounds** — Identified as architecturally significant. The brand fixer only modifies text-level properties within shapes. Slide backgrounds, layouts, and masters are not touched. This explains why the BEFORE_FIXED output shows white backgrounds on section divider slides that should have UQ purple. Options to discuss with Sean.

**Decisions made:**
- Near-black auto-correction: Use triple heuristic (explicit list + distance + luminance/saturation) rather than just increasing the distance threshold — this is more predictable and avoids catching saturated dark colours like dark reds or dark greens.
- Slide backgrounds/masters: **Out of scope for automated tool.** Confirmed via programmatic analysis that the BEFORE deck's placeholder indices (PH 0=title, PH 10=body, PH 31=subtitle, etc.) match the UQ template's indices. When an LDO manually reassigns a slide layout in PowerPoint, text content repositions correctly. Manual layout selection is a design judgement call that doesn't lend itself well to automation.

**Key findings — Image audit:**
- Image audit tested live on BEFORE deck via Streamlit Cloud: 67 images classified, USD $0.51 total cost (~0.76c per image).
- Results: 14 critical, 43 high, 5 medium, 5 low, 0 clear. High count likely inflated because classifier doesn't know UQ holds Adobe Stock licence.
- AFTER deck analysis reveals the LDO attribution convention: on-slide text boxes with "Source: Adobe Stock {ID}" or similar. Also Shutterstock, Microsoft Stock, Wikimedia Commons, and public domain attributions found. This convention should inform how the classifier treats stock photos.
- Sean confirmed image audit is an identification tool at this stage — image replacement is a future component.

**Key findings — Placeholder structure:**
- BEFORE deck layouts that match UQ template names exactly: Cover 1, Title and Content, Two Content, Three content layout, Quote 2, Text with Image Half/Alt, Text with Image One Third/Alt, Contents 1/2, Title Only, Thank You, Icons & Text, Order, Title Subtitle 2 Graphs.
- Unmatched layouts are mostly "1_", "2_", "3_" prefixed duplicates (from an older template version) plus Section Divider 2 (template has "Section Divider" without the "2").
- Placeholder indices are consistent between decks — PH 0 (title), PH 10 (body/content), PH 31 (subtitle), PH 17/18 (footer/slide number). This confirms text will reposition correctly when layouts are manually reassigned.

**Changes to codebase:**
- `brand_fixer.py`: Added NEAR_BLACK_COLOURS constant, updated `_fix_text_frame_colours()` with expanded near-black detection logic
- `app.py`: Refactored both tabs to use `st.session_state` for result persistence
- `PROJECT_LOG.md`: Updated with Session 2 outcomes

---

### Session 1 — 2026-03-25

**Focus:** Full Phase 1 build — both scripts from zero to tested.

**Outcomes:**
- Read and understood all reference documents (project instructions, setup checklist, session protocol, deep research report, PPTX skill docs)
- Analysed all 3 PowerPoint files programmatically:
  - UQ Template: 63 slides, 46 layouts, Arial font, all UQ brand colours confirmed (#51247A, #962A8B, #2B1D37, #D7D1CC, #E62645, #4085C6, #FBB800)
  - BEFORE deck: 148 slides, 75 images, 66 layouts (messy), 1 linked/external image on slide 12
  - AFTER deck: 155 slides, 16 layouts (cleaned), 150 non-Arial font instances remaining
- Built `brand_fixer.py` (Step 1) — 6 fix categories:
  1. Font normalisation: all non-Arial → Arial (preserves bold/italic/size, resolves theme font codes correctly)
  2. Text colour correction: smart rules — near-black→#2B1D37, titles→#51247A, respects white-on-dark, flags non-UQ accents
  3. Table restyling: header row #51247A with white text, alternating rows white/#D7D1CC
  4. Footer standardisation: font → Arial, oversized footers → 10pt
  5. Heading size normalisation: titles clamped to 28–44pt range
  6. Bullet consistency: non-standard bullet chars normalised
- Tested brand_fixer.py on AFTER deck:
  - **141 total changes**: 49 font fixes, 30 colour corrections, 35 flagged for review, 7 tables restyled, 1 heading size fix, 19 bullet normalisations
  - **Verification passed**: Re-scan of FIXED file shows 0 non-Arial fonts (171 Arial + 161 theme font references, all correct)
  - Fonts fixed: Helvetica Neue (slide 19), Times New Roman (slide 43), minion-pro (slide 71), Aptos (slide 88), Helvetica (slide 139)
  - Colours flagged: #7F7F7F grey (slide 22), #00B050 green and #EB602B orange (slide 32) — likely intentional accent colours
- Built `image_audit.py` (Step 2) — full pipeline:
  - Image extraction with deduplication (SHA-256 hash), metadata harvesting (alt text, shape name, slide context, EXIF copyright, hyperlinks)
  - Claude Vision API classification with structured JSON response (11 image types, 5 risk levels, 5 recommended actions)
  - Professional HTML report with UQ branding, risk summary bar, sortable image cards with thumbnails
  - JSON data export alongside HTML
  - `--no-classify` mode for extraction-only testing
  - `--limit N` for cost-controlled test runs
- Tested image extraction on BEFORE deck (20 images): all extracted correctly, proper naming (`slide{N}_image{M}.{ext}`), deduplication working, TIFF/PNG/JPG all handled

**Decisions made:**
- Colour handling: Auto-correct with smart rules — see Decision Log for detail
- Theme font codes (+mj-lt/+mn-lt): Leave alone — they resolve to Arial via theme, touching them risks breaking theme inheritance
- Table restyling: Always treat first row as header — simple heuristic, may need refinement for tables without headers
- Image classification model: claude-sonnet-4-20250514 — cheaper than Opus, still excellent at vision tasks
- Bullet normalisation: Conservative approach — only fix truly non-standard chars (e.g. §), leave common bullets (•, –, ▪) alone

**Issues encountered:**
- Slide 12 in BEFORE deck: shape "Graphic 56" is type PICTURE but has no embedded image data — it's a linked/external graphic (likely SVG/EMF). `python-pptx` raises `ValueError("no embedded image")`. The script catches this gracefully and continues.
- File deletion in Cowork sandbox required explicit permission grant via `allow_cowork_file_delete` tool.

**Changes to codebase:**
- `brand_fixer.py`: Created — 450 lines, fully functional, tested
- `image_audit.py`: Created — 550 lines, extraction tested, API classification ready pending key
- `PROJECT_LOG.md`: Updated with full session outcomes
- `Powerpoints/Advanced Change Management AFTER_FIXED.pptx`: Generated test output from brand_fixer
- `Powerpoints/Advanced Change Management AFTER_FIXED.json`: Detailed change report (JSON)

**Landmines / Watch out for:**
- The AFTER deck has 155 slides (more than BEFORE's 148) — the project instructions said "~90 after" which doesn't match. The AFTER file may be a different version than described.
- Table restyling assumes first row is always a header — may produce incorrect results on tables where the first row is data, not headers. Visual QA needed.
- Colour flagging reports 35 items on the AFTER deck — many are #7F7F7F grey which might be intentionally used for de-emphasised text. The LDOs should decide if grey is acceptable or needs to be added to the approved palette.
- `image_audit.py` resizes images to max 1024px before API call to control costs — very small details (fine print watermarks, tiny copyright symbols) might be lost. Consider increasing to 2048px if classification accuracy seems low.
- The extraction deduplicates by image content hash — if the same image appears on multiple slides, only the first instance is reported. This is intentional (reduces API calls) but the report should note that duplicates were skipped.

---

## Decision Log

| Date | Decision | Rationale | Alternatives Considered | Status |
|------|----------|-----------|------------------------|--------|
| 2026-03-25 | Auto-correct colours with smart rules | Saves LDO time vs flag-only; smart rules avoid breaking white-on-dark text; non-UQ accents flagged for manual review rather than auto-corrected | Flag only; fix all blindly; fix safe + flag ambiguous | Active |
| 2026-03-25 | Use python-pptx for brand fixer (not XML unpacking) | Project instructions specify this approach; simpler, more maintainable | XML manipulation via unpack/edit/pack workflow | Active |
| 2026-03-25 | Leave theme font codes (+mj-lt, +mn-lt) alone | They resolve to Arial via the theme. Overwriting them with literal "Arial" could break theme inheritance and cause issues if the template theme font ever changes. | Force all to literal "Arial" | Active |
| 2026-03-25 | First table row = header (always) | Simple, covers 95%+ of real-world cases. Tables without headers are rare in academic presentations. | Detect header heuristically by formatting; require user flag | Active |
| 2026-03-25 | Use claude-sonnet-4-20250514 for image classification | Best cost/quality balance for vision tasks. At ~$3/1M input tokens + $15/1M output tokens, a 75-image deck costs ~$0.50-1.00. | Claude Opus (more accurate but 5x cost); Claude Haiku (cheaper but less reliable on nuanced classification) | Active |
| 2026-03-25 | Conservative bullet normalisation | Only fix truly weird characters (§). LDOs may have intentional bullet styles. More aggressive normalisation risks breaking visual consistency. | Normalise all to one character; leave all alone | Active |
| 2026-03-25 | Image deduplication by SHA-256 hash | Avoids classifying identical images multiple times (saves API cost). First occurrence is reported with its slide number. | Report every instance (more complete but more expensive); deduplicate by visual similarity (complex) | Active |
| 2026-03-25 | Slide backgrounds/masters left out of scope | Placeholder indices match between BEFORE deck and UQ template. LDO manually selects correct layout; text repositions automatically. Automating layout selection carries high risk of breaking content positioning and is ultimately a design judgement. | Auto-apply matching layouts by name; flag non-matching backgrounds only; full template remapping | Active |

---

## Known Issues & Gotchas

- **PPTX theme font codes (+mj-lt, +mn-lt)** — These are not real font names. They're references to the theme's major/minor font (both Arial in UQ template). The brand fixer correctly ignores them. First encountered: 2026-03-25.
- **AFTER deck slide count mismatch** — Project instructions say "~90 after" but actual AFTER file has 155 slides. May be a different version of the deck. First encountered: 2026-03-25.
- **Linked/external images** — Some shapes report as PICTURE type but have no embedded image data (e.g., linked SVGs or EMFs). python-pptx raises `ValueError("no embedded image")`. Both scripts handle this gracefully. First encountered: 2026-03-25, slide 12 of BEFORE deck.
- **Cowork file deletion** — Deleting files in the mounted workspace folder requires an explicit permission grant via `allow_cowork_file_delete`. First encountered: 2026-03-25.

---

## Dependencies & Environment

| Dependency | Version | Purpose | Notes |
|-----------|---------|---------|-------|
| python-pptx | latest (installed) | PPTX reading and manipulation | Core library for both steps |
| Pillow | latest (installed) | Image processing | For image extraction, resizing, EXIF reading |
| anthropic | latest (installed) | Claude Vision API calls | For Step 2 image classification |
| lxml | latest (installed) | XML manipulation | For table cell fill operations in brand_fixer |

### Environment Variables

| Variable | Purpose | Where Set | Notes |
|----------|---------|-----------|-------|
| ANTHROPIC_API_KEY | Claude Vision API auth | User to configure at runtime | Needed for Step 2 image audit; not needed for `--no-classify` mode |

---

## File Map

```
P-1000/
├── PROJECT_LOG.md                                        — This file (living project log)
├── COWORK_SESSION_PROTOCOL.md                            — Session continuity instructions
├── cowork_project_instructions.md                        — Full project spec and requirements
├── cowork_setup_checklist.md                             — Setup guide
├── deep_research_report.md                               — Copyright compliance research
├── app.py                                                — [DONE] Streamlit web UI (deploy to Streamlit Cloud)
├── requirements.txt                                      — Python dependencies
├── .streamlit/
│   └── config.toml                                       — UQ theme colours + upload size limit
├── .gitignore                                            — Keeps secrets and test data out of GitHub
├── HOW_TO_RUN.md                                         — Deployment guide (GitHub → Streamlit Cloud)
├── brand_fixer.py                                        — [DONE] Step 1 brand compliance script
├── image_audit.py                                        — [DONE] Step 2 image audit script (needs API key for classification)
├── ref_checker.py                                        — [DONE] Step 3 APA 7 reference & attribution checker
├── combined_pipeline.py                                  — [DONE] Combined pipeline (brand + ref + image in one run)
├── Powerpoints/
│   ├── UQ PPT Template - February 2026.pptx              — Canonical UQ template (46 layouts)
│   ├── Advanced Change Management BEFORE.pptx            — Example messy deck (148 slides, 75 images)
│   ├── Advanced Change Management AFTER.pptx             — Example cleaned deck (155 slides, some rogue fonts)
│   ├── Advanced Change Management AFTER_FIXED.pptx       — [OUTPUT] Brand-fixed version of AFTER deck
│   └── Advanced Change Management AFTER_FIXED.json       — [OUTPUT] Detailed change report for AFTER_FIXED
```

---

## Changelog

- **2026-04-09** — BUGFIX + FEATURE — Session 4: Fixed title colour (explicitly set UQ Purple for theme-inherited titles). Fixed table header text (explicitly set white on purple headers). Updated image classifier: priority guide for diagrams vs decorative, gov/corporate screenshot handling, alt text generation. Added version number to app header (v1.4.0). Triaged Evie & Sarah's testing feedback (12 items).
- **2026-03-26 09:00** — FEATURE — Session 3: Overhauled image classification prompt for Adobe Stock/Shutterstock/Microsoft Stock licence awareness. Added `_detect_attribution()` for pre-classification attribution scanning. Added `--footer-text` CLI option and `footer_text` param. Added `flag_body_text_sizes()` body text size check (12–24pt, flag-only). All integrated into Streamlit UI and combined pipeline.
- **2026-03-26 08:00** — FEATURE — Session 2c: Built combined pipeline (`combined_pipeline.py`) running brand fixer → ref checker → image audit in one go. Added as fourth Streamlit tab "Full Compliance Check". Made HTML image audit reports self-contained with base64 embedded images. All features tested end-to-end.
- **2026-03-25 20:00** — FEATURE — Session 2b: Built APA 7 Reference & Attribution Checker (`ref_checker.py`). Scans citations, references, image attributions. Auto-fixes attribution formatting and citation style. Cross-references citations vs reference list. Added as third Streamlit tab. Tested on both BEFORE and AFTER decks.
- **2026-03-25 19:00** — UPDATE — Session 2 continued: Image audit tested live ($0.51 for 67 images). Adobe Stock licence identified as key refinement needed — classifier currently over-flags stock photos. Slide background issue resolved as out-of-scope (placeholder indices confirmed matching; manual layout selection works). AFTER deck attribution conventions documented. Project log updated for Session 3 handoff.
- **2026-03-25 18:30** — CHANGE — Session 2 fixes: (1) Expanded near-black colour detection in `brand_fixer.py` — added explicit NEAR_BLACK_COLOURS set + luminance/saturation heuristic, flags reduced 55→27. (2) Refactored both tabs in `app.py` to use `st.session_state` for result persistence — download buttons no longer clear the page. (3) Updated PROJECT_LOG.md with Session 2 outcomes.
- **2026-03-25 17:55** — CHANGE — Refactored `app.py` for Streamlit Community Cloud deployment. API key via `st.secrets`, added `.streamlit/config.toml` (UQ theme), `.gitignore`, updated `HOW_TO_RUN.md` with full deployment guide. Added risk-level filter to image audit tab, improved error messaging for missing API key.
- **2026-03-25 17:45** — CHANGE — Built Streamlit web UI (`app.py`) wrapping both tools. Added `requirements.txt` and `HOW_TO_RUN.md` for Evie and Sarah.
- **2026-03-25 17:35** — CHANGE — Updated PROJECT_LOG.md with full Session 1 outcomes, decisions, and issues.
- **2026-03-25 17:30** — CHANGE — Built `image_audit.py`. Tested image extraction on BEFORE deck (20 images). AI classification pending API key.
- **2026-03-25 17:20** — CHANGE — Built `brand_fixer.py`. Tested on AFTER deck: 141 changes, 0 non-Arial fonts remaining. Verification passed.
- **2026-03-25 14:00** — NOTE — Project log created. All reference files read and analysed. Ready to begin development.
