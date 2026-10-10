#!/usr/bin/env python3
"""
prepare_assets.py
------------------
Converts the raw files a lean-consultant already collects per unit
(KPI table PDF, savings table PDF, cumulative-outcome PDF, kaizen-sheet
PDF(s), owner photo) into trimmed, high-res PNGs ready to be embedded
in the Word report by generate_report.js.

Usage:
    python3 prepare_assets.py <input_dir> <output_dir> [--dpi 220]

It does NOT guess which kaizens are the "top 3" — it renders every
page of every kaizen PDF it finds and lets the config.json for the
report simply reference which image files to use (kaizen_01.png,
kaizen_02.png, ...), in the order the consultant chooses.
"""
import argparse
import re
import subprocess
import sys
from pathlib import Path
from PIL import Image, ImageChops


def natural_key(s):
    s = str(s)
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", s)]


def run(cmd):
    subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def trim(img: Image.Image, pad=6, bg=(255, 255, 255)) -> Image.Image:
    """Trim uniform-color margins around content, leaving a small pad."""
    rgb = img.convert("RGB")
    bg_img = Image.new("RGB", rgb.size, bg)
    diff = ImageChops.difference(rgb, bg_img)
    bbox = diff.getbbox()
    if not bbox:
        return img
    l, t, r, b = bbox
    l = max(0, l - pad)
    t = max(0, t - pad)
    r = min(img.width, r + pad)
    b = min(img.height, b + pad)
    return img.crop((l, t, r, b))


def pdf_page_text(pdf_path: Path, page: int) -> str:
    try:
        out = subprocess.run(
            ["pdftotext", "-f", str(page), "-l", str(page), str(pdf_path), "-"],
            capture_output=True, text=True,
            encoding="utf-8", errors="replace",
            check=True,
        )
        return out.stdout or ""
    except Exception:
        return ""


# Markers that appear on a genuine kaizen sheet but not on title/training/
# attendance slides. A page must hit at least KAIZEN_MIN_HITS of these to
# be treated as a kaizen sheet worth putting in the report.
KAIZEN_MARKERS = [
    "kaizen no", "kaizen theme", "counter measure", "countermeasure",
    "root cause", "before counter", "after counter", "5w1h", "waste attacked",
    "benefits", "results after implementation", "target & target date",
    "problem description", "results / benefit",
]
KAIZEN_MIN_HITS = 2

NON_KAIZEN_MARKERS = [
    "training attendance record", "attendance record", "topic – kaizen training",
    "topic - kaizen training", "list of participants",
]


def is_kaizen_sheet_page(text: str) -> bool:
    low = (text or "").lower()
    if any(m in low for m in NON_KAIZEN_MARKERS):
        return False
    hits = sum(1 for m in KAIZEN_MARKERS if m in low)
    return hits >= KAIZEN_MIN_HITS


def extract_savings_value(text: str):
    """Pull the largest plausible annual/monthly savings figure mentioned on
    a kaizen sheet so kaizens can be ranked by monetary impact, per the
    master prompt's 'Top 3 by highest savings' rule. Indian kaizen sheets
    write currency both ways ("Rs. 12,480", "28800Rs.", "12,500/-"), so we
    match number-before-currency and currency-before-number alike. Returns
    the largest number found, or 0 if nothing looks like a currency figure.
    """
    candidates = []
    patterns = [
        r"(?:₹|Rs\.?|INR)\s*([\d][\d,]*(?:\.\d+)?)",          # Rs. 12,480 / ₹1,040
        r"([\d][\d,]*(?:\.\d+)?)\s*(?:₹|Rs\.?|INR)\b",         # 28800Rs. / 6500 Rs
        r"([\d][\d,]*(?:\.\d+)?)\s*/-",                        # 12,500/-
    ]
    for pat in patterns:
        for m in re.finditer(pat, text, re.IGNORECASE):
            raw = m.group(1).replace(",", "")
            try:
                val = float(raw)
                if val >= 100:  # ignore tiny numbers (kaizen counts, dates, etc.)
                    candidates.append(val)
            except ValueError:
                pass
    return max(candidates) if candidates else 0.0


def pdf_to_pngs(pdf_path: Path, out_stub: Path, dpi: int, filter_kaizen_pages=False, rank_by_savings=False):
    """Render pages of a PDF to trimmed PNGs: <out_stub>-1.png, -2.png, ...
    If filter_kaizen_pages is True, only pages whose text looks like an
    actual kaizen sheet (not a title/training/attendance slide) are kept.
    If rank_by_savings is True, the kept pages are additionally re-ordered
    highest-savings-first (using the value found by extract_savings_value)
    so that config.json's kaizenImages[0:3] is genuinely the top 3, not
    just the first 3 in file order.
    """
    tmp_prefix = out_stub.parent / (out_stub.name + "_tmp")
    n_pages = int(subprocess.run(["pdfinfo", str(pdf_path)], capture_output=True, text=True)
                   .stdout.split("Pages:")[1].split()[0])

    page_texts = {p: pdf_page_text(pdf_path, p) for p in range(1, n_pages + 1)}
    keep_pages = list(range(1, n_pages + 1))
    if filter_kaizen_pages:
        keep_pages = [p for p in keep_pages if is_kaizen_sheet_page(page_texts[p])]
        if not keep_pages:  # fall back: better to include everything than emit nothing
            keep_pages = list(range(1, n_pages + 1))

    if rank_by_savings:
        keep_pages = sorted(keep_pages, key=lambda p: extract_savings_value(page_texts[p]), reverse=True)

    run(["pdftoppm", "-png", "-r", str(dpi), str(pdf_path), str(tmp_prefix)])
    produced = sorted(out_stub.parent.glob(tmp_prefix.name + "-*.png"), key=natural_key)
    by_page = {i: p for i, p in enumerate(produced, start=1)}

    results = []
    savings_found = []
    counter = 1
    order = keep_pages if rank_by_savings else [p for p in range(1, n_pages + 1) if p in keep_pages]
    for page_no in order:
        src = by_page.get(page_no)
        if src is None:
            continue
        im = Image.open(src)
        im = trim(im)
        final = out_stub.parent / f"{out_stub.name}-{counter}.png"
        im.save(final)
        results.append(final)
        savings_found.append(extract_savings_value(page_texts[page_no]))
        counter += 1
    # clean up any leftover unkept renders
    for p in produced:
        if p.exists():
            p.unlink()
    return results, savings_found


def photo_to_png(img_path: Path, out_path: Path, max_w=900):
    im = Image.open(img_path).convert("RGB")
    if im.width > max_w:
        ratio = max_w / im.width
        im = im.resize((max_w, int(im.height * ratio)), Image.LANCZOS)
    out_path = out_path.with_suffix(".png")
    im.save(out_path)
    return out_path


def process_unit_folder(unit_dir: Path, out_dir: Path, dpi: int):
    out_dir.mkdir(parents=True, exist_ok=True)
    manifest = {"kpi": None, "savings": None, "cumulative": None, "owner": None, "kaizens": [], "kaizen_savings": []}

    files = list(unit_dir.rglob("*"))
    for f in files:
        if not f.is_file():
            continue
        name = f.name.lower()
        if f.suffix.lower() == ".pdf":
            if "kpi" in name:
                res, _ = pdf_to_pngs(f, out_dir / "kpi", dpi)
                manifest["kpi"] = str(res[0]) if res else None
            elif "saving" in name:
                res, _ = pdf_to_pngs(f, out_dir / "savings", dpi)
                manifest["savings"] = str(res[0]) if res else None
            elif "cumulative" in name:
                res, _ = pdf_to_pngs(f, out_dir / "cumulative", dpi)
                manifest["cumulative"] = str(res[0]) if res else None
            elif "kaizen" in name:
                stub = out_dir / f"kaizen_{f.stem}".replace(" ", "_")
                res, savings = pdf_to_pngs(f, stub, dpi, filter_kaizen_pages=True, rank_by_savings=True)
                manifest["kaizens"].extend(str(r) for r in res)
                manifest["kaizen_savings"].extend(savings)
        elif f.suffix.lower() in (".jpg", ".jpeg", ".png"):
            if "owner" in name:
                res = photo_to_png(f, out_dir / "owner")
                manifest["owner"] = str(res)

    # A unit's kaizens often arrive as several separate single-kaizen PDFs
    # (KAIZEN-1.pdf, KAIZEN-2.pdf, ...) rather than one combined PDF - each
    # file was internally ranked above, but that only orders pages *within*
    # a file. Do one final cross-file sort here so kaizens[0:3] is genuinely
    # the top 3 by savings across the whole unit, not just per-file.
    if manifest["kaizens"]:
        paired = sorted(zip(manifest["kaizens"], manifest["kaizen_savings"]), key=lambda kv: kv[1], reverse=True)
        manifest["kaizens"] = [k for k, _ in paired]
        manifest["kaizen_savings"] = [s for _, s in paired]

    return manifest


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("input_dir")
    ap.add_argument("output_dir")
    ap.add_argument("--dpi", type=int, default=220)
    args = ap.parse_args()

    in_dir = Path(args.input_dir)
    out_dir = Path(args.output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    # Each immediate sub-folder of input_dir = one unit. Folders may follow
    # the "1_Unit Name" / "1. Unit Name" convention from the master prompt -
    # strip the leading number for the output folder name, but process in
    # that numeric order (falls back to alphabetical if no number present).
    import json, re as _re

    def unit_sort_key(p):
        m = _re.match(r"^\s*(\d+)[\._\-]\s*", p.name)
        return (0, int(m.group(1))) if m else (1, p.name.lower())

    def clean_unit_name(p):
        return _re.sub(r"^\s*\d+[\._\-]\s*", "", p.name).strip()

    all_manifests = {}
    for unit_dir in sorted((p for p in in_dir.iterdir() if p.is_dir()), key=unit_sort_key):
        clean_name = clean_unit_name(unit_dir)
        unit_out = out_dir / clean_name.replace(" ", "_")
        print(f"Processing unit: {unit_dir.name}  ->  {clean_name}")
        manifest = process_unit_folder(unit_dir, unit_out, args.dpi)
        all_manifests[clean_name] = manifest

    with open(out_dir / "assets_manifest.json", "w") as f:
        json.dump(all_manifests, f, indent=2)
    print(f"\nDone. Manifest written to {out_dir / 'assets_manifest.json'}")


if __name__ == "__main__":
    main()
