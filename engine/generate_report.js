/**
 * generate_report.js
 * -------------------
 * Builds an "Intermediate Stage Completion Report" .docx that follows the
 * exact formatting/layout rules of the GOE-164 golden sample:
 *   - fixed watermark "© AB Associates. All Rights Reserved." on every page
 *   - running header (title + AB Associates logo + rule) on every page
 *   - footer: full-width centered gradient bar + left report name / right page X of Y
 *   - disclaimer pinned to the bottom of the cover page (not floating mid-page)
 *   - KPI / Lean-savings / Cumulative-outcome tables embedded as IMAGES
 *   - auto-generated Index (Word TOC field, same 3-level structure as golden report)
 *   - GOE-level outcomes (heading+table+note) kept together
 *   - USP of GOE capped at 80 words
 *   - Lean Savings table has NO heading/label above it (sits directly under the
 *     KPI table, matching the golden report); kaizen sheets enlarged, no caption
 *     paragraph placed above them
 *   - only real kaizen-sheet images are used (asset pipeline already filtered out
 *     training/attendance pages)
 *
 * Usage:  node generate_report.js <config.json> <output.docx>
 */
const fs = require("fs");
const path = require("path");
const { imageSize } = require("image-size");
const sizeOf = (file) => imageSize(fs.readFileSync(file));
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  HorizontalPositionRelativeFrom, HorizontalPositionAlign,
  VerticalPositionRelativeFrom, VerticalPositionAlign, TextWrappingType,
  PageNumber, TableOfContents, LevelFormat, convertInchesToTwip, ShadingType,
  VerticalAlign, PageBreak, SectionType, FrameAnchorType, FrameWrap,
} = require("docx");

const CFG_PATH = process.argv[2];
const OUT_PATH = process.argv[3] || "out/report.docx";
if (!CFG_PATH) {
  console.error("Usage: node generate_report.js <config.json> <output.docx>");
  process.exit(1);
}
const ROOT = path.dirname(path.resolve(CFG_PATH));
const cfg = JSON.parse(fs.readFileSync(CFG_PATH, "utf8"));
const FORCE = process.argv.includes("--force");

// ---------- pre-flight validation -----------------------------------------
// This is the guardrail that stops the two classes of defect that have
// previously reached a "final" report: (1) a cover/branding logo silently
// missing because its asset path was wrong/blank, and (2) unfinished
// [EDIT: ...] placeholder text - or a Section F with an empty
// Observations/Recommendations block - getting baked into the delivered
// .docx as if it were finished copy. Both are treated as fatal errors
// (not warnings) unless --force is passed, because a report with either
// problem does not match the GOE-164 standard and should never leave this
// tool looking "done".
function validateConfig(cfg) {
  const errors = [];

  // 1) Every branding/cover asset must resolve to a real file on disk.
  //    These four images are what put the MSME / QCI / MCLS + AB Associates
  //    branding on every page - if any path is missing, wrong, or blank,
  //    fail loudly here instead of silently shipping a report with a gap
  //    where a logo should be.
  const REQUIRED_ASSETS = {
    "assets.watermark": cfg.assets && cfg.assets.watermark,
    "assets.abLogo": cfg.assets && cfg.assets.abLogo,
    "assets.coverLogosBand": cfg.assets && cfg.assets.coverLogosBand,
    "assets.footerBar": cfg.assets && cfg.assets.footerBar,
  };
  for (const [key, relPath] of Object.entries(REQUIRED_ASSETS)) {
    if (!relPath) {
      errors.push(`${key} is missing from config.json - the cover page will be missing its MSME/QCI/MCLS or AB Associates branding.`);
      continue;
    }
    const abs = resolveAssetRaw(relPath);
    if (!fs.existsSync(abs)) {
      errors.push(`${key} points to "${relPath}" which does not exist (resolved to ${abs}). Fix the path or restore the file in assets/.`);
    }
  }

  // 2) No unresolved [EDIT: ...] placeholder may reach the final report.
  //    Walk the whole config tree and report every one, with its exact
  //    location, so it can be filled in rather than silently shipped.
  function walk(node, pathStr) {
    if (typeof node === "string") {
      if (node.includes("[EDIT")) {
        errors.push(`Unfilled placeholder at ${pathStr}: "${node}"`);
      }
    } else if (Array.isArray(node)) {
      node.forEach((v, i) => walk(v, `${pathStr}[${i}]`));
    } else if (node && typeof node === "object") {
      for (const [k, v] of Object.entries(node)) walk(v, pathStr ? `${pathStr}.${k}` : k);
    }
  }
  walk(cfg, "");

  // 3) Every unit's Section F (Observations + Recommendations) must be
  //    structurally complete - this is what was found missing/truncated
  //    in past reports (an "Observations:" block with no "Recommendations:"
  //    block underneath it at all).
  (cfg.units || []).forEach((u, i) => {
    const who = `units[${i}] (${u.name || "unnamed unit"})`;
    if (!Array.isArray(u.observations) || u.observations.length === 0) {
      errors.push(`${who}: "observations" is empty - Section F needs at least one Observations paragraph.`);
    }
    if (!Array.isArray(u.recommendationsIntro) || u.recommendationsIntro.length === 0) {
      errors.push(`${who}: "recommendationsIntro" is empty - Section F's Recommendations block needs an opening line.`);
    }
    if (!Array.isArray(u.recommendations) || u.recommendations.length !== 3) {
      errors.push(`${who}: "recommendations" must have exactly 3 entries (found ${Array.isArray(u.recommendations) ? u.recommendations.length : 0}).`);
    }
    if (!Array.isArray(u.recommendationsClosing) || u.recommendationsClosing.length === 0) {
      errors.push(`${who}: "recommendationsClosing" is empty - Section F needs a closing paragraph.`);
    }
  });

  // 4) The GOE-wide "Overall Outcome (for all units, cumulative)" table
  //    under 1.3 GOE-level Outcomes is mandatory. This used to only be
  //    caught deep inside document assembly (goeIntroSection(), well after
  //    validateConfig() had already passed), which meant a run could churn
  //    through building most of the document before failing, and the error
  //    didn't say *where* it had looked. Checking it here, up front, with
  //    full diagnostics, catches a stale/relocated config or a deleted
  //    prepared_assets folder immediately instead of deep into a build.
  {
    const runtimeCumDir = path.resolve(ROOT, "prepared_assets");
    const searched = [];
    const cumImages = discoverCumulativeImages();
    searched.push(`prepared_assets/ next to config.json (${runtimeCumDir}) - ${fs.existsSync(runtimeCumDir) ? "exists" : "does NOT exist"}`);
    searched.push(`cfg.cumulativeOutcomeImages = ${JSON.stringify(cfg.cumulativeOutcomeImages || null)}`);
    searched.push(`cfg.cumulativeOutcomeImage = ${JSON.stringify(cfg.cumulativeOutcomeImage || null)}`);
    if (cumImages.length === 0) {
      errors.push(
        `GOE cumulative outcome table: no rendered page(s) found - the "1.3 GOE-level Outcomes" ` +
        `table would come out blank. Looked in:\n      ` + searched.join("\n      ") +
        `\n    Fix: put the GOE-wide "Cumulative Outcome" table source (PDF/DOCX/XLSX/image) anywhere ` +
        `inside raw_data, then re-run scripts/find_cumulative_source.py + scripts/render_cumulative_source.py ` +
        `(or just re-run RUN_REPORT_GENERATOR.bat from scratch) before generating again.`
      );
    }
  }

  if (errors.length) {
    console.error(`\n${errors.length} problem(s) found - report NOT generated:\n`);
    errors.forEach((e) => console.error("  - " + e));
    if (FORCE) {
      console.error("\n--force was passed: generating anyway. This report is NOT final - it will contain the gaps listed above.\n");
    } else {
      console.error("\nFix these in config.json and re-run. (Pass --force to generate a draft anyway - not recommended for submission.)\n");
      process.exit(1);
    }
  }
}

function resolveAssetRaw(p) {
  if (!p) return p;
  if (path.isAbsolute(p)) return p;

  // Normal case: paths in config.json are relative to the config file.
  const fromConfig = path.resolve(ROOT, p);
  if (fs.existsSync(fromConfig)) return fromConfig;

  // Recovery case: config_auto.json may have been moved/copied while the
  // generator itself still has its bundled assets.  Resolve relative to
  // engine/generate_report.js as a fallback. This prevents a valid bundled
  // asset such as cover_logos_band.png from failing only because config.json
  // lives in a different folder.
  const fromEngine = path.resolve(__dirname, "..", p);
  if (fs.existsSync(fromEngine)) return fromEngine;

  // Also support configs that contain a plain asset filename/path such as
  // "assets/cover_logos_band.png" while the project root is one level above
  // the config directory.
  const fromProject = path.resolve(__dirname, "..", "..", p);
  if (fs.existsSync(fromProject)) return fromProject;

  // Return the original resolved path so the existing validation/error
  // message still reports exactly what path could not be found.
  return fromConfig;
}

validateConfig(cfg);

// ---------- constants matching the golden report -------------------------
// Values below were measured directly from GOE-164's document.xml/styles.xml
// (not eyeballed), so headings, spacing and fonts render pixel-for-pixel the
// same as the golden sample:
//   H1  "1. Introduction of GOE" / "2. Unit-Wise Summary"  -> Heading1 style:
//        Calibri Light, 20pt (sz 40), color 2E5395, NOT bold, spacing before ~4pt, line=single
//   H2  "1.1 ..." / "2.1 UNIT NAME"                         -> Heading2/3 style:
//        Calibri Light, 14pt (sz 28), color black,   NOT bold, spacing before ~9pt, line=single
//   H3  "A. Unit Profile" / "B." / "C." / "D." / "E." / "F." -> Heading4 style:
//        Calibri (regular), 12pt (sz 24), color black, BOLD,  spacing before ~9pt, line=single
//   Body/bullets/labels                                      -> BodyText/ListParagraph/Heading6:
//        Calibri, 11pt (sz 22), spacing BEFORE ~8-9pt (not after), line=single (NOT 1.15x)
const BLUE = "2E5395";           // H1 section-title blue - exact color sampled from GOE-164 styles.xml
const BAR_GREY = "BFBFBF";
const FONT = "Calibri";
const FONT_LIGHT = "Calibri Light"; // used for Heading1/Heading2-equivalent headings, per golden doc
const PAGE_W = 11906, PAGE_H = 16838; // A4 in twips
const MARGIN = 1080;             // ~0.75"
const HEADER_DIST = 460;
const FOOTER_DIST = 460;
const CONTENT_W = PAGE_W - MARGIN * 2; // usable width in twips

function resolveAsset(p) {
  if (!p) return null;
  return resolveAssetRaw(p);
}

// Cumulative outcome tables are a multi-page source. Treat every rendered
// page as mandatory and discover the sequence from prepared_assets even when
// config_auto.json contains only the first-page pointer.
function discoverCumulativeImages() {
  const found = [];
  const add = (p) => { if (p && fs.existsSync(p)) found.push(path.resolve(p)); };
  // The authoritative runtime location is prepared_assets. This avoids a
  // stale/relative config pointer ever collapsing the cumulative source.
  const runtimeCumDir = path.resolve(ROOT, 'prepared_assets');
  if (fs.existsSync(runtimeCumDir)) {
    for (const name of fs.readdirSync(runtimeCumDir)) {
      if (/^(goe_cumulative|cumulative[_-]?outcome|overall[_-]?outcome)[_-]?\d*\.(png|jpg|jpeg)$/i.test(name)) add(path.join(runtimeCumDir, name));
    }
  }
  if (Array.isArray(cfg.cumulativeOutcomeImages)) {
    for (const p of cfg.cumulativeOutcomeImages) add(resolveAsset(p));
  }
  if (cfg.cumulativeOutcomeImage) {
    const first = resolveAsset(cfg.cumulativeOutcomeImage);
    add(first);
    if (first) {
      const m = first.match(/^(.*?)(?:-\d+)?\.(png|jpg|jpeg)$/i);
      if (m) {
        for (let n = 1; n <= 100; n++) add(`${m[1]}-${n}.${m[2]}`);
      }
    }
  }
  const preparedDirs = [
    path.resolve(ROOT, 'prepared_assets'),
    path.resolve(__dirname, '..', 'prepared_assets'),
  ];
  for (const dir of preparedDirs) {
    if (!fs.existsSync(dir)) continue;
    for (const name of fs.readdirSync(dir)) {
      if (/^(goe_cumulative|cumulative[_-]?outcome|overall[_-]?outcome)[_-]?\d*\.(png|jpg|jpeg)$/i.test(name)) add(path.join(dir, name));
    }
  }
  return [...new Set(found)].sort((a,b) => {
    const na = parseInt((a.match(/-(\d+)\.[^.]+$/) || [,9999])[1], 10);
    const nb = parseInt((b.match(/-(\d+)\.[^.]+$/) || [,9999])[1], 10);
    return na - nb;
  });
}

function imgDims(file, maxWtwips, maxHtwips) {
  if (!file || !fs.existsSync(file)) {
    console.error(`\nMissing image file: ${file}\nCheck the corresponding path in config.json - the report cannot be generated without it.\n`);
    process.exit(1);
  }
  const dim = sizeOf(file);
  const pxToTwip = 15; // ~96dpi screen px -> twip approx (1px=15twip @96dpi); we instead scale by ratio below
  let wTwip = (dim.width / dim.height) * (maxHtwips || 999999);
  let hTwip = maxHtwips || 999999;
  if (!maxHtwips || wTwip > maxWtwips) {
    wTwip = maxWtwips;
    hTwip = (dim.height / dim.width) * maxWtwips;
  }
  return { width: Math.round(wTwip), height: Math.round(hTwip) };
}

// twip -> EMU for ImageRun transformation (docx wants px-equivalent via "transformation" in EMU-independent units;
// docx-js ImageRun transformation is in *pixels at 96dpi*, so convert twips->px: px = twip/1440*96
function twipToPx(t) { return Math.round((t / 1440) * 96); }

function pic(file, maxWtwips, maxHtwips, opts = {}) {
  const d = imgDims(file, maxWtwips, maxHtwips);
  return new ImageRun({
    type: path.extname(file).slice(1).toLowerCase() === "jpg" ? "jpg" : path.extname(file).slice(1).toLowerCase(),
    data: fs.readFileSync(file),
    transformation: { width: twipToPx(d.width), height: twipToPx(d.height) },
    ...opts,
  });
}

// ---------- watermark (floating, behind text, centered on every page) ----
function watermarkParagraph() {
  const wmPath = resolveAsset(cfg.assets.watermark);
  const d = sizeOf(wmPath);
  const target = 620; // twips-independent px target width on page
  const w = target, h = Math.round((d.height / d.width) * target);
  return new Paragraph({
    children: [
      new ImageRun({
        type: "png",
        data: fs.readFileSync(wmPath),
        transformation: { width: w, height: h },
        floating: {
          horizontalPosition: { relative: HorizontalPositionRelativeFrom.PAGE, align: HorizontalPositionAlign.CENTER },
          verticalPosition: { relative: VerticalPositionRelativeFrom.PAGE, align: VerticalPositionAlign.CENTER },
          behindDocument: true,
          wrap: { type: TextWrappingType.NONE },
          allowOverlap: true,
        },
      }),
    ],
  });
}

// ---------- running header (title text + AB logo + rule) -----------------
// isFirstPage=true additionally renders the MSME/QCI/MCLS logos band BELOW
// this title+rule row (used only for page 1, via the section's titlePage:
// true + headers.first - see buildDocument()). It's rendered as part of the
// HEADER, not as page-1 body content. Per the golden GOE-164 report, the
// separating rule sits below the logos band on page 1 (not below the title
// row) -- on later pages, with no logos band, the rule stays on the title
// row as the header/body divider.
function buildHeader(isFirstPage = false) {
  const abLogo = resolveAsset(cfg.assets.abLogo);
  const d = imgDims(abLogo, 1800, 420);
  const children = [watermarkParagraph()];
  const rule = { bottom: { style: BorderStyle.SINGLE, size: 6, color: "000000", space: 4 } };
  children.push(new Paragraph({
    tabStops: [{ type: "right", position: CONTENT_W }],
    children: [
      new TextRun({ text: `Intermediate Stage Completion Report (${cfg.goeCode})`, size: 18, font: FONT }),
      new TextRun({ text: "\t" }),
      pic(abLogo, 1800, 420),
    ],
  }));
  if (isFirstPage) {
    const logosBand = resolveAsset(cfg.assets.coverLogosBand);
    children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 120 }, border: rule, children: [pic(logosBand, CONTENT_W, 1500)] }));
  }
  return new Header({ children });
}

// ---------- footer (full-width gradient bar + left/right text) -----------
function buildFooter() {
  const bar = resolveAsset(cfg.assets.footerBar);
  const d = sizeOf(bar);
  const barW = CONTENT_W;
  const barH = Math.round((d.height / d.width) * barW);
  return new Footer({
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 60 },
        children: [pic(bar, barW, null)],
      }),
      new Paragraph({
        tabStops: [{ type: "right", position: CONTENT_W }],
        children: [
          new TextRun({ text: "Intermediate Stage Completion Report", size: 18, font: FONT }),
          new TextRun({ text: "\t" }),
          new TextRun({ text: "Page ", size: 18, font: FONT }),
          new TextRun({ children: [PageNumber.CURRENT], bold: true, size: 18, font: FONT }),
          new TextRun({ text: " of ", size: 18, font: FONT }),
          new TextRun({ children: [PageNumber.TOTAL_PAGES], bold: true, size: 18, font: FONT }),
        ],
      }),
    ],
  });
}

// ---------- small helpers --------------------------------------------------
// spacing/size/font/color values below are matched exactly to GOE-164's
// actual styles.xml / document.xml (verified byte-for-byte against the
// golden .docx, not estimated):
//   Heading1 style: before=360 (18pt), after=80  (4pt)
//   Heading2 style: before=160 (8pt),  after=80  (4pt)
//   Heading3 style: before=160 (8pt),  after=80  (4pt)
//   Body-text paragraphs: line=276/auto (1.15x), NOT single/exact 1.0x -
//     this is the line spacing used on 21 of the 25 body paragraphs sampled
//     from the golden file; the previous line:240/auto here was exact
//     single-spacing, which reads visibly tighter line-to-line than GOE-164.
// Using 240/auto (exact single spacing) and smaller before/after values
// than the golden styles is what produced the "content sits too close under
// the header" look - not the header's own position, which is untouched by
// any of this.
function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    keepNext: true,
    spacing: { before: 360, after: 80, line: 276, lineRule: "auto" },
    children: [new TextRun({ text, color: BLUE, bold: false, size: 40, font: FONT_LIGHT })],
  });
}
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    keepNext: true,
    spacing: { before: 160, after: 80, line: 276, lineRule: "auto" },
    children: [new TextRun({ text, bold: false, size: 28, font: FONT_LIGHT, color: "000000" })],
  });
}
function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    keepNext: true,
    spacing: { before: 160, after: 80, line: 276, lineRule: "auto" },
    children: [new TextRun({ text, bold: true, size: 24, font: FONT, color: "000000" })],
  });
}
// Same heading style as h3(), but with a non-bold "(Kindly refer ...)" note
// appended in the same paragraph/line, in normal body-text weight and size -
// this matches GOE-164 exactly for "B. Lean Implementation Overview" and
// "D. Top 3 Kaizens in detail:", which both carry this note directly on the
// heading line instead of any further written-out detail underneath.
function h3WithNote(text, note, opts = {}) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    keepNext: true,
    pageBreakBefore: opts.pageBreakBefore === true,
    spacing: { before: 160, after: 80, line: 276, lineRule: "auto" },
    children: [
      new TextRun({ text, bold: true, size: 24, font: FONT, color: "000000" }),
      new TextRun({ text: ` (${note})`, bold: false, size: 22, font: FONT, color: "000000" }),
    ],
  });
}
function body(text, opts = {}) {
  return new Paragraph({
    spacing: { before: 160, after: 0, line: 276, lineRule: "auto" },
    children: [new TextRun({ text, size: 22, font: FONT, ...opts })],
  });
}
function bullet(children, opts = {}) {
  return new Paragraph({
    bullet: { level: 0 },
    spacing: { before: 180, after: 0, line: 276, lineRule: "auto" },
    keepLines: true,
    keepNext: opts.keepNext === true,
    children,
  });
}
function labelValue(label, value, opts = {}) {
  return bullet([
    new TextRun({ text: label + ": ", bold: true, size: 22, font: FONT }),
    new TextRun({ text: value || "", size: 22, font: FONT }),
  ], opts);
}
// Label with NO value after the colon - matches GOE-164's "B. Lean
// Implementation Overview" bullets exactly (Product Family included:,
// Lean Tools Applied:, etc. are all left blank there, since that whole
// section is covered by the unit's PPT instead - see h3WithNote() above).
function labelOnly(label, opts = {}) {
  return bullet([new TextRun({ text: label + ":", bold: true, size: 22, font: FONT })], opts);
}
function imageParagraph(file, maxWtwips, opts = {}) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    keepNext: opts.keepNext === true,
    keepLines: true,
    spacing: { before: 160, after: 0, line: 240, lineRule: "auto" },
    children: [pic(file, maxWtwips, opts.maxH || null)],
  });
}
function italicNote(text, opts = {}) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    keepNext: opts.keepNext === true,
    keepLines: true,
    spacing: { before: 120, after: 0, line: 240, lineRule: "auto" },
    children: [new TextRun({ text, italics: true, size: 20, font: FONT })],
  });
}
function wordCount(s) { return (s || "").trim().split(/\s+/).filter(Boolean).length; }
function capWords(s, n) {
  const words = (s || "").trim().split(/\s+/);
  if (words.length <= n) return s;
  return words.slice(0, n).join(" ") + "…";
}

// ---------- Cover page -----------------------------------------------------
// Text sizing/weight/color here is matched exactly to GOE-164's document.xml:
// the ENTIRE cover page (every line except the disclaimer) is one uniform
// sz=32 (16pt), and nothing on the cover is bold - not the GOE code, not the
// scheme name, not the prepared-by name. The three section labels ("Under
// Scheme", "Submitted To", "Prepared By") are color 404040 (gray); the value
// lines beneath them are plain black. Previously this used a mix of
// 24/26/34pt with bold on several lines, which is why the cover didn't match
// GOE-164 - that's fixed below.
const COVER_LABEL_COLOR = "404040";
function coverPage() {
  const children = [];
  // before:860 (up from 600) - measured directly against GOE-164: the gap
  // between the header's separator rule and this title line runs ~260
  // twips (~13pt) shorter here than in the golden file at the same 600
  // value. This only adds space ABOVE the title (i.e. below the header) -
  // it does not move the header itself, and there's ample free space
  // lower on the cover page (the "Prepared By" block ends well short of
  // the bottom-anchored disclaimer frame) so this can't push anything
  // onto page 2.
  children.push(new Paragraph({ spacing: { before: 860, after: 200 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "MCLS Lean Intermediate Stage Completion Report Of", size: 32, font: FONT, color: COVER_LABEL_COLOR })] }));
  children.push(new Paragraph({ spacing: { after: 500 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: cfg.goeCode, size: 32, font: FONT })] }));
  children.push(new Paragraph({ spacing: { before: 400, after: 100 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "Under Scheme", size: 32, font: FONT, color: COVER_LABEL_COLOR })] }));
  children.push(new Paragraph({ spacing: { after: 500 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "MSME Competitive (Lean) Scheme - 2022", size: 32, font: FONT })] }));
  children.push(new Paragraph({ spacing: { before: 400, after: 100 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "Submitted To", size: 32, font: FONT, color: COVER_LABEL_COLOR })] }));
  children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 },
    children: [new TextRun({ text: "MSME (Ministry of Micro, Small & Medium Enterprises)", size: 32, font: FONT })] }));
  children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 },
    children: [new TextRun({ text: "Government of India", size: 32, font: FONT })] }));
  children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 },
    children: [new TextRun({ text: "&", size: 32, font: FONT })] }));
  children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 500 },
    children: [new TextRun({ text: "Quality Council of India (QCI)", size: 32, font: FONT })] }));
  children.push(new Paragraph({ spacing: { before: 400, after: 100 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "Prepared By", size: 32, font: FONT, color: COVER_LABEL_COLOR })] }));
  children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 },
    children: [new TextRun({ text: cfg.preparedByName, size: 32, font: FONT })] }));
  children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 1800 },
    children: [new TextRun({ text: cfg.preparedByRole, size: 32, font: FONT })] }));

  // Anchored to the bottom of page 1 via a Word "frame" (absolute position
  // relative to the page margins) rather than a large spacer paragraph.
  // The previous approach (an empty paragraph with a big fixed "after"
  // spacing) pushed the disclaimer down by a constant amount regardless of
  // how much content came before it - so on any run where the cover page
  // content was a little taller than usual (e.g. longer GOE code/consultant
  // name wrapping to 2 lines), the cumulative spacing overflowed page 1 and
  // the disclaimer landed at the TOP of page 2 instead. A frame anchored to
  // the page's bottom margin always renders in the same place regardless of
  // what's above it, so it can't overflow onto the next page.
  children.push(new Paragraph({
    border: { top: { style: BorderStyle.SINGLE, size: 4, color: "999999", space: 6 } },
    frame: {
      type: "alignment",
      alignment: { x: HorizontalPositionAlign.CENTER, y: VerticalPositionAlign.BOTTOM },
      anchor: { horizontal: FrameAnchorType.MARGIN, vertical: FrameAnchorType.MARGIN },
      width: CONTENT_W,
      height: 1300,
      wrap: FrameWrap.NONE,
    },
    children: [new TextRun({
      text: `Disclaimer: This report has been prepared exclusively for the information of the Ministry of MSME, Government of India (GoI), and the Quality Council of India (QCI) under the MSME (Competitive) Lean Scheme. The contents of this report are confidential and must not be shared, reproduced, or distributed in any form without prior written consent from M/s AB Associates and the respective unit owners.`,
      size: 18, font: FONT, italics: false,
    })],
  }));
  return children;
}


// ---------- page-break-with-gap helper --------------------------------------
// GOE-164's header sits noticeably farther above the body content than a
// bare page break alone produces here. The obvious fix - adding
// `spacing.before` to whatever paragraph lands first on the new page - does
// NOT work: Word/LibreOffice both unconditionally ignore a paragraph's
// `spacing.before` when that paragraph is the first one on a page (whether
// the break arrived via an explicit PageBreak run or a `pageBreakBefore`
// property on the paragraph itself). This was verified directly: setting
// before anywhere from 0 to 800 twips on such a paragraph produced an
// IDENTICAL rendered position every time.
// The fix that works is a tiny near-invisible spacer paragraph carrying real
// line-height (which is never suppressed the way paragraph spacing is),
// marked with `pageBreakBefore` so IT starts the new page.
//
// This must NOT be done as a separate paragraph containing only a manual
// `<w:br type="page"/>` run followed by a second spacer paragraph - that
// two-paragraph form caused a real bug: if the content right before the
// break already fills the current page almost exactly to its bottom edge,
// that near-empty break-carrier paragraph's own line box (bigger than it
// looks, since it uses the document's default 11pt style) can itself fail
// to fit in the sliver of space remaining, spilling onto a page of its own
// - producing a genuinely blank page before the real content resumes one
// page later than intended. This was reproduced directly from a real
// generated report: the page right after the GOE-level cumulative-outcome
// table came out completely blank, pushing "2. Unit-Wise Summary" (and
// every unit's KPI table after it) one full page later than it should be.
// Using `pageBreakBefore` on the spacer paragraph itself avoids this
// entirely, since it is then a single paragraph/property, not two separate
// flow objects where the first can overflow independently of the second.
function pageBreak() {
  return [
    new Paragraph({
      pageBreakBefore: true,
      spacing: { before: 0, after: 0, line: 240, lineRule: "auto" },
      children: [new TextRun({ text: " ", size: 14, font: FONT, color: "FFFFFF" })],
    }),
  ];
}

// ---------- Acknowledgement -------------------------------------------------
// Font sizes here are matched exactly to GOE-164's document.xml: the
// "Acknowledgement" heading is sz=28 (14pt) - NOT 30/15pt, which is what the
// generic title-page heading size elsewhere in this file uses - and every
// body paragraph on this page is sz=24 (12pt, via GOE-164's Heading5 style),
// one size up from the sz=22 (11pt) body text used everywhere else in the
// document. ackBody() exists solely to carry that 12pt override; regular
// body() (11pt) is intentionally not reused here.
function ackBody(text) {
  return new Paragraph({
    spacing: { before: 160, after: 0, line: 276, lineRule: "auto" },
    children: [new TextRun({ text, size: 24, font: FONT })],
  });
}
function acknowledgementPage() {
  const children = [];
  children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, alignment: AlignmentType.CENTER, keepNext: true,
    spacing: { after: 80 },
    children: [new TextRun({ text: "Acknowledgement", bold: true, size: 28, font: FONT, color: "000000" })] }));
  children.push(ackBody(cfg.acknowledgement.intro));
  if (cfg.acknowledgement.msmeThanks) {
    children.push(ackBody(cfg.acknowledgement.msmeThanks.pre));
    for (const line of cfg.acknowledgement.msmeThanks.lines) children.push(ackBody(line));
  }
  if (cfg.acknowledgement.qciThanks) {
    children.push(ackBody(cfg.acknowledgement.qciThanks.pre));
    for (const line of cfg.acknowledgement.qciThanks.lines) children.push(ackBody(line));
  }
  if (cfg.acknowledgement.auditorThanks) children.push(ackBody(cfg.acknowledgement.auditorThanks));
  children.push(ackBody(cfg.acknowledgement.unitThanksPre));
  for (const u of cfg.units) {
    children.push(ackBody(`${u.ownerName}, ${u.ownerRole}`));
  }
  children.push(ackBody(cfg.acknowledgement.closingLine));
  children.push(ackBody("Regards,"));
  children.push(ackBody(cfg.acknowledgement.signOff));
  return children;
}


// ---------- Index / TOC ------------------------------------------------------
// "Index" heading matched exactly to GOE-164: sz=32 (16pt), Calibri Light,
// NOT bold (previously this used 30/15pt bold Calibri, which is the
// Acknowledgement-heading style, not the Index-heading style - they're
// different in the golden file).
function indexPage() {
  return [
    new Paragraph({ heading: HeadingLevel.HEADING_1, alignment: AlignmentType.CENTER, keepNext: true,
      spacing: { after: 80 },
      children: [new TextRun({ text: "Index", bold: false, size: 32, font: FONT_LIGHT, color: "000000" })] }),
    new TableOfContents("Index", { hyperlink: true, headingStyleRange: "1-3" }),
  ];
}

// ---------- GOE Introduction -------------------------------------------------
function introSection() {
  const children = [];
  children.push(h1("1. Introduction of GOE"));
  children.push(h2("1.1 Basic Information"));
  children.push(body(`Name of the Consultant Organisation: ${cfg.consultantOrg}`, { bold: false }));
  children.push(body(`Name of the Consultants Involved: ${cfg.consultantNames}`));
  children.push(body(`Period of Implementation: ${cfg.periodStart} to ${cfg.periodEnd}`));
  children.push(h2("1.2 GOE Overview"));
  children.push(labelValue("GOE Name", cfg.goeCode));
  children.push(labelValue("Location", cfg.location));
  children.push(labelValue("Total Units in GOE", String(cfg.totalUnits)));
  children.push(labelValue("Units Participated in Lean Implementation", String(cfg.unitsParticipated)));
  const usp = capWords(cfg.uspOfGoe, 80);
  children.push(bullet([
    new TextRun({ text: "Unique Features/USP of GOE: ", bold: true, size: 22, font: FONT }),
    new TextRun({ text: usp, size: 22, font: FONT }),
  ]));
  // "1.3 GOE-level Outcomes" flows naturally right after the USP bullet,
  // exactly like GOE-164 itself: checked its real document.xml directly,
  // and there is no forced page break here at all - "2. Unit-Wise Summary"
  // later in that file only happens to start a fresh page because of where
  // natural pagination landed for that specific report's content length
  // (marked by Word's own <w:lastRenderedPageBreak/> cache, not a real
  // break in the document structure). Forcing a break here, as a previous
  // fix did, contradicts that and doesn't match the golden report.
  children.push(h2("1.3 GOE-level Outcomes"));
  children.push(new Paragraph({ spacing: { after: 80 },
    children: [new TextRun({ text: "Overall Outcome (for all units, cumulative)", size: 22, font: FONT })] }));

  // The cumulative outcome source is a multi-page PDF. Every rendered page
  // is inserted in numeric order; never silently fall back to a single page.
  const uniqueCumulativeImages = discoverCumulativeImages();
  if (uniqueCumulativeImages.length) {
    uniqueCumulativeImages.forEach((img) => {
      // maxH must stay generous here: it exists only to stop an unusually
      // tall multi-row table from overflowing the page, not to compete with
      // width. A tight cap (previously 6200) can bind before CONTENT_W does
      // for tables with fewer rows, leaving the image narrower than the
      // page's text width -- which is the "size alignment" mismatch against
      // the golden GOE-164 report (its cumulative table spans full width).
      // keepNext:true binds this image only to its own caption
      // (cfg.cumulativeOutcomeNote, pushed right after this loop) - not to
      // the heading or label above, which deliberately have no such
      // binding here, matching GOE-164's real document.xml (only its
      // Heading2 style carries an inherent keepNext; nothing else in this
      // section does). This narrow pairing was chosen after seeing both
      // extremes fail on the real report that reproduced this: binding
      // the whole heading+label+image+caption chain together (an earlier
      // attempt) relocated the entire ~4-inch block whenever longer intro
      // text left too little room, wasting most of two pages; no binding
      // at all left the one-line caption stranded alone at the top of an
      // otherwise blank page, since the image can run right up against
      // the bottom margin on its own and leave no room for it. Binding
      // only the image to its own caption means just those two move
      // together on the rare occasion they don't fit, while the heading
      // and label above stay on whatever page they were already on.
      children.push(imageParagraph(img, CONTENT_W, { keepNext: true, maxH: 9000 }));
    });
    if (cfg.cumulativeOutcomeNote) children.push(italicNote(cfg.cumulativeOutcomeNote));
  } else {
    throw new Error('CUMULATIVE OUTCOME TABLE IS MISSING. The report will not be generated. Put the GOE-wide Cumulative Outcome table source inside raw_data and re-run.');
  }
  return children;
}

// ---------- one unit's "2.x" section -----------------------------------------
function unitSection(u, idx) {
  const children = [];
  // GOE-164 keeps "2. Unit-Wise Summary" and unit 2.1 on the SAME page -
  // the page break happens before the "2." heading itself (see build()),
  // not before the first unit. Every unit after the first still starts
  // on its own fresh page, matching the golden report (2.2 starts a new
  // page, 2.3 starts a new page, etc.)
  if (idx > 0) children.push(...pageBreak());
  children.push(h2(`2.${idx + 1} ${u.name.toUpperCase()}`));

  children.push(h3("A. Unit Profile"));
  children.push(labelValue("Name of Unit", u.name));
  children.push(labelValue("Location & State", u.address));
  children.push(labelValue("Product(s)", u.products));
  children.push(bullet([new TextRun({ text: "Turnover (In Lakh Rs.):", bold: true, size: 22, font: FONT })]));
  for (const t of u.turnover) {
    children.push(new Paragraph({ indent: { left: 720 }, bullet: { level: 1 },
      spacing: { before: 160, after: 0, line: 240, lineRule: "auto" },
      children: [new TextRun({ text: `FY ${t.fy} = Rs. ${t.value}`, size: 22, font: FONT })] }));
  }
  children.push(labelValue("Lean Coordinator", `${u.leanCoordinatorName} (Contact Number: ${u.phone}, Email: ${u.email})`));

  // GOE-164 leaves every bullet under "B" blank (label + colon, nothing
  // after it) and instead points the reader to the unit's PPT via a note
  // right on the heading line - it does NOT write out productFamily /
  // leanTools / goodPractices / beforeAfter as prose. Matching that
  // exactly is also what keeps this section short enough that A + B + C's
  // KPI table reliably land on the same page (see the maxH note below) -
  // an earlier version of this file wrote real sentences into these
  // bullets, which was extra height that (combined with the KPI image)
  // regularly failed to fit on one page and pushed C to the next page.
  children.push(h3WithNote("B. Lean Implementation Overview", "Kindly refer the attached unit PPT for this section"));
  children.push(labelOnly("Product Family included"));
  children.push(labelOnly("Lean Tools Applied"));
  children.push(labelOnly("Key Good Practices/Processes implemented and established"));
  children.push(labelOnly("Before & After Photos (Tool/Project-wise)"));

  children.push(h3("C. Key Performance Indicators (KPIs) Summary"));
  // Keep the KPI table with A + B, as in the golden report.
  if (u.kpiImage) children.push(imageParagraph(resolveAsset(u.kpiImage), CONTENT_W, { maxH: 6200 }));

  // Pagination rule: the Lean Savings table comes FIRST. Section D's heading
  // is intentionally BELOW the Lean Savings table, followed immediately by
  // the first Kaizen sheet. The savings table, source note, Section D heading
  // and first Kaizen are chained so they remain on the same page when they fit.
  if (u.savingsImage) {
    children.push(imageParagraph(resolveAsset(u.savingsImage), CONTENT_W, { keepNext: true, maxH: 5200 }));
    children.push(italicNote("The above figures are as per IT system", { keepNext: true }));
  } else {
    children.push(italicNote("The above figures are as per IT system", { keepNext: true }));
  }

  children.push(h3WithNote("D. Top 3 Kaizens in detail:", "Kindly refer to the attached unit PPT where all kaizens are explained in detail", {
    keepNext: true,
  }));

  const kaizens = (u.kaizenImages || []).slice(0, 3);
  kaizens.forEach((k, i) => {
    const maxH = (i === 0 && u.savingsImage) ? 6000 : 6900;
    children.push(imageParagraph(resolveAsset(k), CONTENT_W, { keepNext: false, maxH }));
  });

  // pageBreakBefore on the heading paragraph itself suffers the exact same
  // spacing.before suppression described in pageBreak() above, so the break
  // and the gap-spacer are inserted as separate preceding paragraphs instead.
  children.push(...pageBreak());
  children.push(new Paragraph({ heading: HeadingLevel.HEADING_3, keepNext: true, spacing: { before: 180, after: 0, line: 240, lineRule: "auto" }, children: [new TextRun({ text: "E. Testimonial from the MSME Owners:", bold: true, size: 24, font: FONT, color: "000000" })] }));
  if (u.ownerImage) {
    children.push(new Paragraph({ keepNext: true, spacing: { after: 100 },
      children: [pic(resolveAsset(u.ownerImage), 1800, 2200)] }));
  }
  children.push(body(u.ownerName, { bold: true }));
  children.push(body(u.ownerRole));
  for (const para of u.testimonial) children.push(body(para));

  children.push(h3("F. Consultant's closing observations & recommendations:"));
  children.push(new Paragraph({ keepNext: true, spacing: { after: 80 },
    children: [new TextRun({ text: "Observations:", bold: true, size: 22, font: FONT })] }));
  for (const para of u.observations) children.push(body(para));
  children.push(new Paragraph({ keepNext: true, spacing: { after: 80 },
    children: [new TextRun({ text: "Recommendations:", bold: true, size: 22, font: FONT })] }));
  for (const para of u.recommendationsIntro) children.push(body(para));
  for (const rec of u.recommendations) {
    children.push(bullet([
      new TextRun({ text: rec.title + ": ", bold: true, size: 22, font: FONT }),
      new TextRun({ text: rec.body, size: 22, font: FONT }),
    ]));
  }
  for (const para of u.recommendationsClosing) children.push(body(para));
  return children;
}

// ---------- assemble ---------------------------------------------------------
// buildDocument() returns a brand-new Document object graph (fresh Header/
// Footer/Paragraph/ImageRun instances every time it's called). Keeping this
// separate from packing/writing lets packAndVerify() below retry with a
// completely fresh object graph if a build comes out with a branding image
// silently missing (see packAndVerify() for why that check exists).
function buildDocument() {
  const body1 = coverPage();
  const body2 = [...acknowledgementPage()];
  const body3 = [...indexPage()];
  // "2. Unit-Wise Summary" always starts a fresh page in GOE-164 (unit 2.1
  // then follows immediately underneath it on that same page - see the
  // idx>0 check in unitSection()).
  const body4 = [...introSection(), ...pageBreak(), h1("2. Unit-Wise Summary")];
  cfg.units.forEach((u, i) => { body4.push(...unitSection(u, i)); });
  body4.push(new Paragraph({ spacing: { before: 300 },
    children: [new TextRun({ text: "End of Report", italics: true, size: 22, font: FONT, color: "000000" })], alignment: AlignmentType.CENTER }));

  // IMPORTANT: this must stay a SINGLE section, matching the GOE-164 golden
  // report's actual structure (verified directly from its document.xml -
  // it has exactly one <w:headerReference>/<w:footerReference> pair for the
  // whole file, cover page included). Previously this built TWO sections
  // (cover page as its own section, everything else as a second section),
  // which meant buildHeader()/buildFooter() were each called twice -
  // producing two independently-constructed Header objects with floating
  // images. The `docx` library silently drops the second floating image
  // (the watermark) from the FIRST of those two Header objects, which is
  // exactly the header governing the cover page - hence "logos/watermark
  // missing only on page 1" while every later page (rendered from the
  // second, intact Header object) was fine. Building one section - so
  // buildHeader()/buildFooter() are each called exactly once - avoids that
  // failure mode entirely and matches the golden file byte-for-byte in
  // structure. Page breaks (already used elsewhere in this file) handle the
  // cover -> acknowledgement -> index -> body transitions that the old code
  // got "for free" from the section break.
  return new Document({
    features: { updateFields: true },
    styles: {
      default: { document: { run: { font: FONT, size: 22 } } },
    },
    sections: [
      {
        properties: {
          page: { size: { width: PAGE_W, height: PAGE_H }, margin: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN, header: HEADER_DIST, footer: FOOTER_DIST } },
          titlePage: true,
        },
        headers: { default: buildHeader(false), first: buildHeader(true) },
        footers: { default: buildFooter(), first: buildFooter() },
        children: [
          ...body1,
          ...pageBreak(),
          ...body2,
          ...pageBreak(),
          ...body3,
          ...pageBreak(),
          ...body4,
        ],
      },
    ],
  });
}

// ---------- pack, verify, retry -----------------------------------------
// Two separate branding images (the header watermark, and - even after that
// fix - the cover-page logos band) have each been observed missing from a
// "successfully" generated .docx, with no error thrown: the `docx` library
// occasionally drops one image's relationship/media entry during Packer.
// toBuffer() on documents with this many embedded images, silently. That's
// exactly the class of defect this tool exists to prevent (see the
// validateConfig() comment above) - so instead of trusting that a completed
// build is a correct build, every attempt is packed, unzipped again in
// memory, and checked byte-for-byte against the four required branding
// assets on disk. If any is missing, the ENTIRE document is rebuilt from
// scratch (fresh Header/Footer/Paragraph/ImageRun objects - not a re-pack of
// the same object graph, since the flake has so far only ever affected
// early-constructed images, so a fresh object graph is more likely to embed
// them cleanly) and retried, rather than ever writing a file with a
// silently-missing logo.
const JSZip = require("jszip");
const MAX_PACK_ATTEMPTS = 4;

// Word's generated TOC is a field, so its entries inherit the normal
// paragraph spacing unless the TOC styles are explicitly defined. A small
// "after" gap on every TOC line is enough to push the last entry onto a
// fourth page in this report, after which the deliberate page break creates
// an otherwise blank page. Normalize TOC1/TOC2/TOC3 to zero after-spacing so
// the complete Index stays on exactly one page, matching GOE-164.
async function patchGoldenLayout(buf) {
  const zip = await JSZip.loadAsync(buf);
  const stylesFile = zip.file("word/styles.xml");
  if (!stylesFile) return buf;
  let styles = await stylesFile.async("string");
  for (const styleId of ["TOC1", "TOC2", "TOC3"]) {
    const re = new RegExp(`<w:style\\s+[^>]*w:styleId="${styleId}"[^>]*>[\\s\\S]*?<\\/w:style>`);
    const match = styles.match(re);
    if (!match) continue;
    let block = match[0];
    block = block.replace(/<w:spacing[^>]*\/>/g, "");
    block = block.replace(/<w:pPr>/, '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>');
    styles = styles.replace(match[0], block);
  }
  zip.file("word/styles.xml", styles);

  // docx.js (the `docx` npm package) intermittently drops the `keepNext`
  // paragraph property specifically from "C. Key Performance Indicators"
  // and "D. Top 3 Kaizens in detail:" - and only those two headings, out of
  // six per unit, on every single unit, 100% reproducibly - while every
  // other Heading1/2/3 paragraph in the same document keeps it correctly.
  // A minimal reproduction with the same heading sequence and a
  // TableOfContents field did NOT reproduce this, so it's tied to some
  // specific interaction in the real document's fuller content rather than
  // heading order alone; rather than chase the exact internal cause
  // further, every Heading1/2/3-styled paragraph in this app is always
  // supposed to carry keepNext (all of h1()/h2()/h3()/h3WithNote() and the
  // Acknowledgement/Index headings pass keepNext:true unconditionally), so
  // guaranteeing it here for any that lost it is safe and correct - not a
  // narrower fix that could accidentally miss a similar case elsewhere.
  // Confirmed directly: without this, "C. Key Performance Indicators"
  // orphaned alone at the bottom of a page while its KPI table image alone
  // moved to the next page, wasting roughly half a page each time.
  const documentXmlFile = zip.file("word/document.xml");
  if (documentXmlFile) {
    let documentXml = await documentXmlFile.async("string");
    documentXml = documentXml.replace(
      /<w:pPr><w:pStyle w:val="Heading[123]"\/>(?!<w:keepNext\/>)/g,
      (m) => m.replace('"/>', '"/><w:keepNext/>')
    );
    zip.file("word/document.xml", documentXml);
  }

  return await patchUniqueDrawingIds(zip);
}

// The `docx` npm library hardcodes every embedded picture's <wp:docPr id="1">
// (and <pic:cNvPr id="0">) to the SAME literal value instead of incrementing
// it per image - see https://github.com/dolanmiu/docx/issues/2719. Per the
// OOXML spec these ids must be unique within each document part. Word
// tolerates the duplication inconsistently: depending on version/patch, it
// can silently render only one of the duplicate-id pictures and leave every
// other one as a blank paragraph-shaped gap - with no error, no red-X
// placeholder, nothing wrong visible in the XML to a byte/structure check,
// and no consistent reproduction outside real Word. This is exactly the
// failure this report hit: the GOE cumulative-outcome picture was correctly
// generated, correctly embedded, and simply lost the "which duplicate wins"
// coin flip in Word's renderer. Renumbering every docPr/cNvPr id to be
// unique removes the ambiguity Word is resolving inconsistently in the
// first place, rather than trying to predict which duplicate Word picks.
async function patchUniqueDrawingIds(zip) {
  // IMPORTANT: header/footer parts render on the SAME page as the body
  // content, so a docPr id that's merely unique *within* document.xml (or
  // *within* header1.xml) is not enough - if header1.xml also has a
  // docPr id="1" and document.xml has a picture with id="1" on that same
  // page, Word sees two shapes with the same id on one page and hits the
  // identical "which duplicate wins" ambiguity this function exists to
  // remove. The counters below are therefore shared across every part
  // (declared outside the loop, never reset), so every docPr/cNvPr id in
  // the whole packaged document - body, both headers, both footers - is
  // globally unique. This is what actually eliminates the ambiguity;
  // resetting per-file (the previous behavior) only guaranteed uniqueness
  // within each file individually and left cross-part collisions intact.
  const partsToPatch = [
    "word/document.xml",
    "word/header1.xml",
    "word/header2.xml",
    "word/footer1.xml",
    "word/footer2.xml",
  ];
  let docPrCounter = 0;
  let picCounter = 0;
  for (const partName of partsToPatch) {
    const file = zip.file(partName);
    if (!file) continue;
    let xml = await file.async("string");
    xml = xml.replace(/<wp:docPr id="\d+"/g, () => `<wp:docPr id="${++docPrCounter}"`);
    xml = xml.replace(/<pic:cNvPr id="\d+"/g, () => `<pic:cNvPr id="${++picCounter}"`);
    zip.file(partName, xml);
  }
  return await zip.generateAsync({ type: "nodebuffer" });
}

// Checks that the cumulative-outcome image(s) aren't just present somewhere
// in word/media (byte-equality can pass even if the picture never actually
// got wired into the document body), but are actually placed as real
// <w:drawing> runs between the "Overall Outcome (for all units, cumulative)"
// paragraph and the following "2. Unit-Wise Summary" heading. This is the
// exact failure signature seen in a delivered report.docx once: a paragraph
// with the correct centered/keepLines/spacing shell that imageParagraph()
// produces, but with zero runs inside it - i.e. every other check (byte
// presence in media/, no thrown error) passed, yet the picture itself was
// never inserted into that paragraph. A missing <w:drawing> count here is
// treated the same as a missing asset: rebuild and retry.
function countCumulativeDrawingsInDocumentXml(documentXml, expectedCount) {
  const startMarker = "Overall Outcome (for all units, cumulative)";
  const endMarker = "2. Unit-Wise Summary";
  const startIdx = documentXml.indexOf(startMarker);
  const endIdx = documentXml.indexOf(endMarker, startIdx);
  if (startIdx === -1 || endIdx === -1 || endIdx <= startIdx) {
    return { ok: false, found: 0, reason: "could not locate the 1.3 GOE-level Outcomes section in document.xml at all" };
  }
  const section = documentXml.slice(startIdx, endIdx);
  const drawingCount = (section.match(/<w:drawing>/g) || []).length;
  return {
    ok: drawingCount >= expectedCount,
    found: drawingCount,
    reason: drawingCount < expectedCount
      ? `found ${drawingCount} <w:drawing> element(s) under "1.3 GOE-level Outcomes" but expected ${expectedCount} (one per cumulative table page) - the picture paragraph(s) came out empty`
      : null,
  };
}

// Confirms patchUniqueDrawingIds() actually achieved what it claims: every
// wp:docPr id, and separately every pic:cNvPr id, unique across ALL of
// document.xml + both headers + both footers combined - not just unique
// within each part. Cross-part collisions are exactly what let the
// cumulative table (or any picture) still go blank in real Word even after
// per-file renumbering. Returns the offending id(s) if uniqueness was not
// achieved, so a regression here fails loudly instead of shipping a report
// that only sometimes renders correctly.
async function verifyGlobalUniqueDrawingIds(zip) {
  const partsToCheck = [
    "word/document.xml",
    "word/header1.xml",
    "word/header2.xml",
    "word/footer1.xml",
    "word/footer2.xml",
  ];
  const seenDocPr = new Map();
  const seenCNvPr = new Map();
  const dupes = [];
  for (const partName of partsToCheck) {
    const file = zip.file(partName);
    if (!file) continue;
    const xml = await file.async("string");
    for (const m of xml.matchAll(/<wp:docPr id="(\d+)"/g)) {
      const id = m[1];
      if (seenDocPr.has(id)) dupes.push(`wp:docPr id="${id}" in both ${seenDocPr.get(id)} and ${partName}`);
      else seenDocPr.set(id, partName);
    }
    for (const m of xml.matchAll(/<pic:cNvPr id="(\d+)"/g)) {
      const id = m[1];
      if (seenCNvPr.has(id)) dupes.push(`pic:cNvPr id="${id}" in both ${seenCNvPr.get(id)} and ${partName}`);
      else seenCNvPr.set(id, partName);
    }
  }
  return { ok: dupes.length === 0, dupes };
}

async function packAndVerify() {
  const requiredAssets = [
    { label: "watermark", path: resolveAsset(cfg.assets.watermark) },
    { label: "AB Associates logo", path: resolveAsset(cfg.assets.abLogo) },
    { label: "cover logos band", path: resolveAsset(cfg.assets.coverLogosBand) },
    { label: "footer bar", path: resolveAsset(cfg.assets.footerBar) },
  ];
  const cumulativeImagePaths = discoverCumulativeImages();
  const cumulativeRequired = cumulativeImagePaths.map((p, i) => ({ label: `cumulative table page ${i + 1}`, path: p }));
  const allRequired = [...requiredAssets, ...cumulativeRequired].map((a) => ({ ...a, bytes: fs.readFileSync(a.path) }));

  for (let attempt = 1; attempt <= MAX_PACK_ATTEMPTS; attempt++) {
    const doc = buildDocument();
    let buf = await Packer.toBuffer(doc);
    buf = await patchGoldenLayout(buf);
    const zip = await JSZip.loadAsync(buf);
    const mediaNames = Object.keys(zip.files).filter((n) => n.startsWith("word/media/"));
    const mediaBuffers = await Promise.all(mediaNames.map((n) => zip.files[n].async("nodebuffer")));

    const missing = allRequired.filter(
      (a) => !mediaBuffers.some((mb) => mb.length === a.bytes.length && mb.equals(a.bytes))
    );

    const documentXml = await zip.file("word/document.xml").async("string");
    const placement = countCumulativeDrawingsInDocumentXml(documentXml, cumulativeImagePaths.length);
    const idUniqueness = await verifyGlobalUniqueDrawingIds(zip);

    if (missing.length === 0 && placement.ok && idUniqueness.ok) {
      fs.mkdirSync(path.dirname(OUT_PATH), { recursive: true });
      fs.writeFileSync(OUT_PATH, buf);
      console.log(
        `Written: ${OUT_PATH}` +
          (attempt > 1 ? ` (attempt ${attempt}/${MAX_PACK_ATTEMPTS} - an earlier attempt this run had a branding image silently dropped by the docx packer; this attempt verified clean)` : "")
      );
      return;
    }

    const reasons = [
      ...missing.map((m) => `${m.label} missing from word/media`),
      ...(placement.ok ? [] : [placement.reason]),
      ...(idUniqueness.ok ? [] : [`duplicate picture id(s) across document/header/footer parts: ${idUniqueness.dupes.join("; ")}`]),
    ];
    console.error(
      `Attempt ${attempt}/${MAX_PACK_ATTEMPTS}: ${reasons.length} problem(s) found in the packed .docx (${reasons.join("; ")}). ` +
        (attempt < MAX_PACK_ATTEMPTS ? "Rebuilding from scratch and retrying..." : "")
    );
  }

  console.error(
    `\nFATAL: after ${MAX_PACK_ATTEMPTS} attempts, this report still came out with one or more required assets missing or ` +
      `not actually placed in the document (including possibly the GOE cumulative outcome table). ` +
      `No file was written - a report in this state must not be treated as final. Please re-run the generator.\n` +
      `\nIf this keeps happening, it is very unlikely to be this script: it means the source cumulative-table image(s) ` +
      `themselves are the problem (e.g. corrupt/zero-byte PNG). Re-run scripts/render_cumulative_source.py on the ` +
      `original source file and confirm the resulting PNG(s) open correctly before generating again.\n`
  );
  process.exit(1);
}

packAndVerify();
