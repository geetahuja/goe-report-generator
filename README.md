# GOE Lean Report Generator

Generates "Intermediate/Completion Stage" reports formatted to match the
**GOE-164 golden sample** exactly: same cover layout, watermark, header/footer,
Index, page-fit rules, and embedded-image tables — built from the same kind
of data folder you already collect per unit (KPI PDF, Savings PDF, kaizen
PDF(s), owner photo, and a "Units Text Information.txt" file).

## What this is (and isn't)

This is **not** a hosted website you upload files into from a browser —
that's not something this chat environment can serve persistently. It's a
**local pipeline** that runs on your own PC, either via one double-click
(Windows) or a couple of terminal commands (Mac/Linux). Once set up, you
genuinely just drop your data into a folder and run it — no JSON-editing
required for anything the source files already tell it.

## Fastest path: Windows, one click

1. Unzip this folder anywhere (e.g. `Downloads\GOE_Report_Generator`).
2. Install Node.js LTS from nodejs.org and Python from python.org, if you
   don't already have them (tick "Add python.exe to PATH" during the Python
   install). Restart any open Command Prompt/PowerShell windows afterward.
3. Inside this folder, create a folder named `raw_data`. Put your unit data
   in it — see "Data layout" below.
4. Double-click **`RUN_REPORT_GENERATOR.bat`**.

It will install what it needs the first time, convert your PDFs/photos into
report-ready images, parse `Units Text Information.txt`, build a
`config_auto.json` with everything it could figure out on its own filled in,
pause so you can fill in the handful of fields that need real writing
(marked `[EDIT: ...]`), then generate `out\report.docx` when you press a key.

## Mac/Linux, or if you'd rather run it by hand

```bash
npm install                              # first time only
pip install pillow --break-system-packages   # first time only

python3 scripts/prepare_assets.py raw_data/ prepared_assets/ --dpi 220
python3 scripts/parse_units_txt.py "raw_data/Units Text Information.txt" prepared_assets/units_parsed.json
python3 -c "
from scripts.prepare_assets import pdf_to_pngs
from pathlib import Path
pdf_to_pngs(Path('raw_data/Cumulative Outcome Table for GOE.pdf'), Path('prepared_assets/goe_cumulative'), 220)
"
python3 scripts/build_config.py prepared_assets/units_parsed.json prepared_assets/assets_manifest.json config_auto.json \
    --cumulative-image prepared_assets/goe_cumulative-1.png \
    --period-start "10 August 2025" --period-end "16 June 2026" \
    --auditors "Name 1, Name 2, Name 3" --only-uploaded

# edit config_auto.json - fill in anything marked [EDIT: ...]

node engine/generate_report.js config_auto.json out/report.docx
```

## Data layout (`raw_data/`)

```
raw_data/
  Units Text Information.txt
  Cumulative Outcome Table for GOE XXX.pdf
  1_Idea Tools/
    KPI Summary Table for Idea Tools.pdf
    Lean savings table for Idea Tools.pdf
    Idea Tools_kaizens.pdf
    Idea Tools Owner Image.jpeg
  2_Unit Two/
    ...
```

Leading numbers on the unit folders (`1_`, `2_`, ...) control the order
units appear in the report — this matches the naming convention already used
in the data you send me. `Units Text Information.txt` should follow the
same "Label : value" format shown in `sample_output/` — that's what gets
auto-parsed into unit profile facts (address, products, process, owner,
lean coordinator, turnover) and the GOE code / consultant name.

## What gets auto-filled vs. what still needs writing

**Auto-filled from your files**, no typing required:
- GOE code, consultant name (from the .txt file)
- Every unit's name, address, products, process, owner name, lean
  coordinator + contact, 3-year turnover (from the .txt file)
- KPI table, Lean Savings table, Cumulative Outcome table (as images)
- Top 3 kaizen sheets **ranked by actual monetary savings** found on each
  sheet (Rs. or ".../-" figures) — not just file order
- A starting set of 3 recommended lean tools per unit, picked by matching
  keywords in that unit's process description (e.g. "welding" leads to Poka
  Yoke, "punching/bending" leads to SMED) so it isn't the same 3 tools for
  every unit

**Still needs a human (or a follow-up chat with me) to write:**
- Testimonial body paragraphs (owner's voice, specific challenges/impact)
- Consultant's observations narrative
- GOE-wide "Unique Features/USP" paragraph (<=80 words - the engine hard-caps
  it, but write it tight so nothing gets cut mid-sentence)
- Period of implementation, auditor names (pass as command-line flags, or
  fill in after `build_config.py` runs)
- MSME DI contact name/designation for the Acknowledgement page

Every one of these is left in `config_auto.json` as a clearly marked
`[EDIT: ...]` placeholder - search the file for `[EDIT` to find all of them.
The easiest way to fill them in well: paste the unit's PPT/description to me
in chat and ask me to draft that unit's testimonial/observations in the
golden report's tone, then paste my draft into the matching field.

**Not filled in at all, on purpose:** "B. Lean Implementation Overview"'s
four bullets (Product Family included / Lean Tools Applied / Key Good
Practices / Before & After Photos) and "D. Top 3 Kaizens in detail:" are
left blank in GOE-164 itself - the report just points to the unit's PPT
via a note on the heading line (see "What's already handled" below). There
is nothing to write here; `config.json`'s `overview` block is kept only for
backward compatibility and isn't rendered.

## Kaizen filtering

The script reads the text of every page in each kaizen PDF and only keeps
pages that actually look like a kaizen sheet (checks for markers like
"Kaizen No.", "Counter Measure", "Root Cause", etc., and explicitly excludes
"Training Attendance Record" style pages). This is what stops training
slides or attendance sheets from ending up in the report. If a real kaizen
sheet ever gets dropped by mistake, open `scripts/prepare_assets.py` and
adjust `KAIZEN_MARKERS` / `KAIZEN_MIN_HITS`.

## What's already handled, matching your punch list

- Watermark is a proper full "(c) AB Associates. All Rights Reserved." on
  every page, not clipped
- Disclaimer sits at the true bottom of the cover page, not floating mid-page
- The cover page uses the same footer as every other page (no missing first-page footer)
- Footer gradient bar is full-width / centered, "Page X of Y"
- KPI, Lean Savings, and Cumulative Outcome tables are embedded as images
- Only real kaizen sheets are used, ranked by savings, training/attendance
  pages filtered out automatically
- Index/TOC follows the same structure as GOE-164 (Word TOC field) and is compacted so the complete Index stays on one page
- "1. Introduction of GOE" (heading + 1.1/1.2/1.3 + USP + cumulative table +
  note) is kept together so it lands on one page
- USP of GOE is hard-capped at 80 words
- "B. Lean Implementation Overview" flows directly after "A. Unit Profile" -
  no forced page break, no separated section, no heading/table stranded
  alone on the page below a large gap (previously an over-chained keepNext
  could force the whole B->C->Lean Savings block onto a fresh page even
  when B itself had plenty of room right under A)
- "B" and "D" match GOE-164 exactly: both are blank (label + colon only,
  or just the heading) with a "(Kindly refer the attached unit PPT for
  this section/where all kaizens are explained in detail)" note on the
  heading line - nothing is auto-written into these sections
- Section headings stay glued to the start of their own content (KPI table,
  first kaizen sheet) - just not chained onward into the section after them
- Lean Savings summary is kept with Section D and the first kaizen sheet, rather than being stranded on its own page
- The KPI table image is capped at ~4.3in tall so it reliably lands on the
  same page as A + B above it (matching GOE-164, where A + B + the KPI
  table are always on one page) instead of a page break landing before C
- Kaizen sheets are enlarged (2 per page), no caption paragraph above them
- Font sizing matches GOE-164 exactly, measured directly from its XML:
  Calibri Light 20pt for H1 section titles, Calibri Light 14pt for H2
  ("2.1 UNIT NAME" etc.), Calibri Bold 12pt for the A/B/C/D/E/F
  sub-headings, Calibri 11pt for body/bullets, Calibri 16pt (no bold) for
  every cover-page line, Calibri Bold 14pt for "Acknowledgement" with its
  body text at Calibri 12pt, Calibri Light 16pt for "Index", and Calibri
  9pt for the header/footer text

## Pre-flight validation (report generation will refuse to run if...)

`engine/generate_report.js` now checks `config.json` **before** touching
`out/report.docx`, and refuses to generate anything if it finds:

- Any of the 4 branding assets (`watermark`, `abLogo`, `coverLogosBand`,
  `footerBar`) missing from `config.json` or pointing at a file that
  doesn't exist. This is what guarantees the MSME / QCI / MCLS + AB
  Associates cover-page logos are on **every** report this tool produces
  - a bad or missing logo path now stops the run with a clear message
  instead of silently shipping a report with a gap on the cover.
- Any leftover `[EDIT: ...]` placeholder anywhere in the config - the
  exact field path and text are printed so you know exactly what still
  needs writing. This is what stops an unfinished paragraph (most often
  in Section F, "Consultant's closing observations & recommendations")
  from ending up in a "final" .docx.
- Any unit whose Section F is structurally incomplete - empty
  `observations`, empty `recommendationsIntro`/`recommendationsClosing`,
  or a `recommendations` array that isn't exactly 3 entries.

If you genuinely need a draft with known gaps (e.g. to review pagination
before the consultant's narrative is ready), pass `--force`:
`node engine/generate_report.js config.json out/draft.docx --force`.
This still generates the file, but only after printing every gap it
found, so a draft can't be mistaken for a finished report.

## Known limitations

- **Index/TOC only populates in real Microsoft Word.** If you preview via
  LibreOffice the Index will look blank - that's a LibreOffice limitation
  (it doesn't evaluate TOC fields on conversion), not a bug in the report.
  In Word, press Ctrl+A then F9 if it doesn't auto-update on open.
- Pagination uses Word's "keep with next" / "keep lines together" paragraph
  properties, not pixel-exact frames. This reliably keeps headings glued to
  their tables/images, but with unusually long testimonials or observations
  text you may occasionally want to glance through the final PDF and nudge
  a paragraph - the same touch-up you'd do by hand in Word.
- `build_config.py`'s recommendation-picker is a simple keyword heuristic,
  meant to give a plausible starting point, not a substitute for a
  consultant reviewing the unit.
- The static program logos (Ministry of MSME / QCI / MCLS) and the AB
  Associates logo in `assets/` were extracted from the GOE-164 sample. If a
  future scheme uses different logos, replace those PNG files.

## Files

```
RUN_REPORT_GENERATOR.bat    one-click Windows launcher (see above)
engine/generate_report.js   the docx assembly engine
scripts/prepare_assets.py   PDF/photo -> print-ready PNG converter + kaizen filter + savings ranker
scripts/parse_units_txt.py  parses "Units Text Information.txt" into structured JSON
scripts/build_config.py     merges parsed txt + prepared images into a ready-to-use config.json
assets/                     static logos, watermark, footer bar (from GOE-164)
sample_output/              a full worked example built from real GOE-213 data
  GOE213_report.docx
  config_goe213_used.json
```


### Fix for `ENOENT ... assets\cover_logos_band.png`

The generator now resolves branding assets relative to both `config_auto.json` and the bundled generator assets. This makes the report generation resilient if `config_auto.json` is copied or generated in a different folder from the project root. The supplied `assets\cover_logos_band.png` is still the authoritative bundled file.

### `assets_backup/`

A verified duplicate of `watermark.png`, `ab_logo.png`, `cover_logos_band.png`,
and `footer_bar.png` (byte-identical to the copies in `assets/` - confirmed
by MD5). This is **not** wired into the generator and doesn't need to be -
it's just insurance in case `assets/` is ever accidentally edited, moved, or
deleted. If that happens, copy these four files back into `assets/`.

Note: if a report ever comes out with a missing logo/watermark, it is *not*
because these source image files are damaged or wrong - they've always been
byte-identical to this backup. That failure mode is the `docx` packing
library occasionally dropping one image during file assembly, which is what
the verify-and-retry step in `engine/generate_report.js` (see "Pre-flight
validation" above) now catches automatically.
