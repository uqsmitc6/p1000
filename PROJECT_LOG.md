# Project Log — UQ Slide Compliance Tool

> **Last updated:** 2026-03-25 18:30
> **Project:** Two-part automated tool for UQ Business School executive education slide deck compliance — brand formatting fixer + copyright image audit
> **Tech stack:** Python (python-pptx, Pillow, lxml), Claude Vision API (anthropic SDK), Streamlit UI, HTML reporting
> **Repository:** Cowork project folder

---

## Current State

Both scripts are **built and functionally tested**, plus a **Streamlit web UI** (`app.py`) wraps both tools for non-technical users. App is **live on Streamlit Community Cloud** with API key configured. Sean tested the brand fixer on the BEFORE deck live — 201 changes made. Three issues were identified from live testing and two have been fixed in Session 2:

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
| 1 | TODO | Test Step 2 AI classification with live API key | 2026-03-25 | High | User will provide ANTHROPIC_API_KEY; run `image_audit.py` on BEFORE deck with `--limit 10` |
| 2 | TODO | Slide backgrounds/masters — scope and approach | 2026-03-25 | High | Brand fixer only fixes text-level formatting; backgrounds/layouts are untouched. Need to decide: (a) apply UQ template layouts automatically, (b) flag slides with non-UQ backgrounds, or (c) leave as out-of-scope for automated tool. See Session 2 notes. |
| 3 | TODO | Tune classification prompt after first test run | 2026-03-25 | Medium | May need to adjust category definitions or risk thresholds based on results |
| 4 | NOTE | BEFORE deck is available as PPTX (148 slides) not just PDF — original instructions said PDF only but PPTX was uploaded | 2026-03-25 | Low | Can extract images directly from PPTX |
| 5 | NOTE | Slide 12 in BEFORE deck has a linked/external image (not embedded) — script handles this gracefully with a warning | 2026-03-25 | Low | The shape is "Graphic 56" — likely an SVG or EMF |
| 6 | TODO | Add `--footer-text` CLI option to brand_fixer.py | 2026-03-25 | Low | Currently standardises footer formatting but doesn't set specific text content |
| 7 | TODO | Consider adding a body text size check (14–20pt range flagging) | 2026-03-25 | Low | Currently only checks title sizes; body text size normalisation deferred |

---

## Session History

### Session 2 — 2026-03-25

**Focus:** Fix three issues identified from live testing before sharing with Evie and Sarah.

**Outcomes:**
- **Near-black colour threshold** — Expanded detection from a simple Euclidean distance check (`< 60` from black) to a three-pronged approach: explicit hex list of 20 common near-black values (#111111 through #545454), the original distance check, and a luminance+saturation heuristic (max channel < 100, saturation < 30). Result: flagged colours dropped from 55 → 27. All 27 remaining flags are genuine accent colours (greens, oranges, reds, blues) that do need human review.
- **Download clearing the page** — Refactored both Brand Fixer and Image Audit tabs to store results in `st.session_state`. Processing happens inside the button click, results are rendered from session state outside the click block. Download buttons now work without clearing the displayed results.
- **Slide backgrounds** — Identified as architecturally significant. The brand fixer only modifies text-level properties within shapes. Slide backgrounds, layouts, and masters are not touched. This explains why the BEFORE_FIXED output shows white backgrounds on section divider slides that should have UQ purple. Options to discuss with Sean.

**Decisions made:**
- Near-black auto-correction: Use triple heuristic (explicit list + distance + luminance/saturation) rather than just increasing the distance threshold — this is more predictable and avoids catching saturated dark colours like dark reds or dark greens.

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
├── Powerpoints/
│   ├── UQ PPT Template - February 2026.pptx              — Canonical UQ template (46 layouts)
│   ├── Advanced Change Management BEFORE.pptx            — Example messy deck (148 slides, 75 images)
│   ├── Advanced Change Management AFTER.pptx             — Example cleaned deck (155 slides, some rogue fonts)
│   ├── Advanced Change Management AFTER_FIXED.pptx       — [OUTPUT] Brand-fixed version of AFTER deck
│   └── Advanced Change Management AFTER_FIXED.json       — [OUTPUT] Detailed change report for AFTER_FIXED
```

---

## Changelog

- **2026-03-25 18:30** — CHANGE — Session 2 fixes: (1) Expanded near-black colour detection in `brand_fixer.py` — added explicit NEAR_BLACK_COLOURS set + luminance/saturation heuristic, flags reduced 55→27. (2) Refactored both tabs in `app.py` to use `st.session_state` for result persistence — download buttons no longer clear the page. (3) Updated PROJECT_LOG.md with Session 2 outcomes.
- **2026-03-25 17:55** — CHANGE — Refactored `app.py` for Streamlit Community Cloud deployment. API key via `st.secrets`, added `.streamlit/config.toml` (UQ theme), `.gitignore`, updated `HOW_TO_RUN.md` with full deployment guide. Added risk-level filter to image audit tab, improved error messaging for missing API key.
- **2026-03-25 17:45** — CHANGE — Built Streamlit web UI (`app.py`) wrapping both tools. Added `requirements.txt` and `HOW_TO_RUN.md` for Evie and Sarah.
- **2026-03-25 17:35** — CHANGE — Updated PROJECT_LOG.md with full Session 1 outcomes, decisions, and issues.
- **2026-03-25 17:30** — CHANGE — Built `image_audit.py`. Tested image extraction on BEFORE deck (20 images). AI classification pending API key.
- **2026-03-25 17:20** — CHANGE — Built `brand_fixer.py`. Tested on AFTER deck: 141 changes, 0 non-Arial fonts remaining. Verification passed.
- **2026-03-25 14:00** — NOTE — Project log created. All reference files read and analysed. Ready to begin development.
