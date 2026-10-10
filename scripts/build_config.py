#!/usr/bin/env python3
"""
build_config.py
-----------------
Combines the output of parse_units_txt.py (unit profile facts) with the
output of prepare_assets.py (image paths, kaizens ranked by savings) into
a config.json that's ready to feed straight into generate_report.js.

Per the master prompt's priority order, the .txt file is authoritative for
profile facts and prepare_assets.py's manifest is authoritative for which
image goes where. The only things this script CANNOT fill in - because
they require real writing, not data lookup - are:
  - testimonial body paragraphs
  - consultant's observations
  - the GOE-wide "Unique Features/USP" paragraph
  - period of implementation / auditor names (not in the .txt file format
    seen so far - pass them with --period-start/--period-end/--auditors,
    or edit the output afterward)
Those are left as clearly marked [EDIT: ...] placeholders, exactly like a
blank field on a paper form - fill them in, or hand the config plus the
unit .txt/PPT to Claude and ask it to draft them.

Usage:
    python3 build_config.py units_parsed.json assets_manifest.json out/config.json \
        --period-start "10 August 2025" --period-end "16 June 2026" \
        --auditors "Mr. A,Mr. B,Mr. C" --location Nashik
"""
import argparse
import json
import os
import re
import difflib
from pathlib import Path


# Very small keyword -> recommendation-tool map, used only to pick a
# plausible *starting* set of 3 recommendations relevant to the unit's own
# process text, per the master prompt's "avoid generic recommendations"
# instruction. This is a heuristic nudge, not a substitute for a consultant
# actually reviewing the unit - it just avoids the report defaulting to
# the same 3 tools for every unit regardless of what they make.
TOOL_KEYWORDS = [
    (["weld", "solder", "braz"], "Poka Yoke", "Error-proofing the welding/joining process will reduce rework and improve first-time-right quality."),
    (["punch", "bend", "press", "die", "stamp"], "SMED", "Faster die/tool changeovers will reduce setup time and support smaller batch sizes."),
    (["assembl", "wiring", "panel", "testing"], "Standard Work", "Standardising the assembly/wiring sequence reduces variation and dependency on individual operators."),
    (["paint", "coat", "plating", "surface"], "Visual Management", "Visual standards for surface-finish acceptance criteria will reduce rejection at final inspection."),
    (["machin", "turning", "milling", "drilling", "grinding"], "TPM", "Preventive/autonomous maintenance on machining centres will reduce breakdown-related downtime."),
    (["mold", "injection", "extrusion"], "Process Capability", "Process-capability studies on critical dimensions will tighten first-time-right yield."),
    (["pack", "corrugat", "carton", "dispatch"], "Line Balancing", "Balancing the corrugation-to-dispatch line will smooth flow and reduce WIP buildup."),
]
DEFAULT_TOOLS = [
    ("Advanced 5S and Visual Management", "More sophisticated visual controls and layered audits will further reduce searching time and minor delays."),
    ("7 QC Tools", "Data-based problem solving will help identify and reduce recurring defects systematically."),
    ("Kaizen", "Continuing structured kaizen activity will build employee involvement and shop-floor ownership."),
]


def pick_recommendations(process_text):
    process_text = (process_text or "").lower()
    picked = []
    for keywords, title, body in TOOL_KEYWORDS:
        if any(k in process_text for k in keywords):
            picked.append({"title": title, "body": body})
        if len(picked) == 3:
            break
    for title, body in DEFAULT_TOOLS:
        if len(picked) >= 3:
            break
        if not any(p["title"] == title for p in picked):
            picked.append({"title": title, "body": body})
    return picked[:3]


def auto_intro_clause(products, process):
    """Returns a lowercase clause like 'engaged in the manufacture of X,
    with core processes including Y' (no leading capital, no trailing
    period) so callers can drop it into different sentence shapes."""
    bits = []
    if products:
        bits.append(f"engaged in the manufacture of {products.strip().rstrip('.')}")
    if process:
        bits.append(f"with core processes including {process.strip().rstrip('.')}")
    return ", ".join(bits)


def _norm(s):
    return re.sub(r"[^a-z0-9]+", "", (s or "").lower())


def _clean_period_value(value, fallback="Not provided in Units Text Information"):
    return value.strip() if isinstance(value, str) and value.strip() else fallback


def auto_testimonial(name, products, process, consultant_name):
    products = (products or "our products").strip().rstrip('.')
    process = (process or "our manufacturing processes").strip().rstrip('.')
    return [
        f"We, {name}, sincerely thank the Ministry of Micro, Small & Medium Enterprises (MSME), Government of India, "
        f"the Quality Council of India (QCI), and the Lean Manufacturing Consultant, {consultant_name} and his team "
        f"from AB Associates, for their valuable guidance and support during the implementation of Lean Manufacturing Practices under the MCLS scheme.",
        f"Our unit is engaged in the manufacture of {products}, with core processes including {process}. The Lean journey has provided a structured approach to strengthen workplace practices, standardisation and continuous improvement around these operations. The focus is to sustain disciplined practices and use systematic problem-solving to improve operational performance.",
        "We remain committed to sustaining these improvements and progressing toward the Advanced Stage of Lean Implementation for long-term operational excellence."
    ]


def auto_observations(name, products, process):
    products = (products or "the unit's products").strip().rstrip('.')
    process = (process or "the stated manufacturing process").strip().rstrip('.')
    return [
        f"{name} is engaged in the manufacture of {products}, with the key manufacturing process identified as {process}. The unit profile indicates a process-oriented manufacturing setup where systematic control and standardisation can support consistent performance.",
        f"Based on the available unit information, the Lean implementation is positioned around practical shop-floor improvement, including 5S, Kaizen and Daily Work Management. For the reported product and process family, continued focus on standard work, visual controls, workplace organisation and data-based problem solving will help sustain the improvements and support further Lean maturity."
    ]


def auto_usp(units):
    products = []
    processes = []
    for u in units:
        if u.get("products") and u["products"] not in products:
            products.append(u["products"].strip().rstrip('.'))
        if u.get("process") and u["process"] not in processes:
            processes.append(u["process"].strip().rstrip('.'))
    product_text = "; ".join(products[:4])
    process_text = "; ".join(processes[:4])
    text = "The GOE brings together diverse MSME manufacturing units covering " + product_text
    if process_text:
        text += ", with processes spanning " + process_text
    text += ". A common focus is workplace organisation, standard work, visual management and continuous improvement across varied manufacturing environments."
    # Keep the source-derived summary comfortably below the report's 80-word cap.
    return " ".join(text.split()[:75])

def build(units_parsed, manifest, args, out_json_path):
    consultant = units_parsed.get("consultant", {})
    consultant_name = args.consultant_name or consultant.get("name") or "[EDIT: consultant name]"

    out_dir = os.path.dirname(os.path.abspath(out_json_path)) or "."
    brand_dir = args.brand_assets_dir or os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "assets")

    def rel(path_str):
        """Rewrite a path (as produced by prepare_assets.py, relative to
        wherever that script was run from) so it resolves correctly from
        wherever THIS config.json ends up living - generate_report.js
        resolves every path relative to the config file's own directory."""
        if not path_str or path_str.startswith("[EDIT"):
            return path_str
        abs_path = os.path.abspath(path_str)
        return os.path.relpath(abs_path, out_dir).replace(os.sep, "/")

    manifest_norm = {_norm(k): v for k, v in manifest.items()}

    units_out = []
    skipped = []
    for u in units_parsed["units"]:
        # manifest keys come from prepare_assets.py's clean_unit_name(), which
        # strips the same "1_" / "1." prefix - match on the cleaned name,
        # normalizing away punctuation/spacing differences ("M/S Amerco Box
        # Pack Pvt. Ltd." vs "MS Amerco Box Pack Pvt Ltd", "Neha Enterprises
        # (Ambad)" vs "Neha Enterprises Ambad", etc).
        m = manifest.get(u["name"]) or manifest_norm.get(_norm(u["name"]))
        if m is None:
            # last resort: fuzzy match for near-miss spelling between the
            # .txt file and the folder name (e.g. "Ozar" in the txt file vs
            # a folder actually named "Ozer") - only accept a close, unique
            # match so we don't silently mix up two different units.
            close = difflib.get_close_matches(_norm(u["name"]), manifest_norm.keys(), n=1, cutoff=0.82)
            if close:
                m = manifest_norm[close[0]]
        if m is None:
            if args.only_uploaded:
                skipped.append(u["name"])
                continue
            m = {}

        turnover = [{"fy": t["fy"], "value": t["value"]} for t in u.get("turnover", [])]
        kaizens = [rel(k) for k in m.get("kaizens", [])[:3]]
        clause = auto_intro_clause(u.get("products"), u.get("process"))
        our_unit_sentence = f"Our unit is {clause}." if clause else ""

        units_out.append({
            "name": u["name"],
            "address": u.get("address") or "[EDIT: address]",
            "products": u.get("products") or "[EDIT: products]",
            "turnover": turnover,
            "leanCoordinatorName": u.get("leanCoordinatorName") or "[EDIT: lean coordinator]",
            "phone": u.get("phone") or "Not provided in Units Text Information",
            "email": u.get("email") or "Not provided in Units Text Information",
            "overview": {
                # GOE-164 (the golden reference) leaves every "B. Lean
                # Implementation Overview" bullet blank - label + colon,
                # nothing after it - and instead points the reader to the
                # unit's PPT via a note on the heading line itself (see
                # h3WithNote() in engine/generate_report.js). generate_report.js
                # no longer renders these values at all, so there is nothing
                # to auto-fill or hand-edit here anymore; they're kept as
                # empty strings only so older config.json files that still
                # set them keep working without erroring.
                "productFamily": "",
                "leanTools": "",
                "goodPractices": "",
                "beforeAfter": "",
            },
            "kpiImage": rel(m.get("kpi")),
            "savingsImage": rel(m.get("savings")),
            "kaizenImages": kaizens,
            "ownerImage": rel(m.get("owner")),
            "ownerName": u.get("ownerName") or "[EDIT: owner name]",
            "ownerRole": u.get("ownerRole") or f"Director, {u['name']}",
            "testimonial": auto_testimonial(u["name"], u.get("products"), u.get("process"), consultant_name),
            "observations": auto_observations(u["name"], u.get("products"), u.get("process")),
            "recommendationsIntro": [f"{u['name']} has laid a solid foundation in lean manufacturing and now stands at a promising inflection point to deepen its transformation."],
            "recommendations": pick_recommendations(u.get("process")),
            "recommendationsClosing": ["Overall, the unit exhibits strong potential to adopt advanced lean tools and systems, and is well-positioned to achieve sustained excellence and competitive advantage."],
        })

    auditors = [a.strip() for a in args.auditors.split(",")] if args.auditors else ["QCI-appointed auditors"]

    config = {
        "goeCode": units_parsed.get("goeCode") or "[EDIT: GOE code]",
        "consultantOrg": "AB Associates",
        "consultantNames": consultant_name,
        "periodStart": _clean_period_value(args.period_start),
        "periodEnd": _clean_period_value(args.period_end),
        "location": args.location or "Nashik",
        "totalUnits": args.total_units or len(units_parsed["units"]),
        "unitsParticipated": len(units_parsed["units"]),
        "preparedByName": f"{consultant_name}, AB Associates",
        "preparedByRole": "Lean Manufacturing Consultant",
        "uspOfGoe": auto_usp(units_parsed["units"]),
        "assets": {
            "watermark": rel(os.path.join(brand_dir, "watermark.png")),
            "abLogo": rel(os.path.join(brand_dir, "ab_logo.png")),
            "coverLogosBand": rel(os.path.join(brand_dir, "cover_logos_band.png")),
            "footerBar": rel(os.path.join(brand_dir, "footer_bar.png")),
        },
        # Preserve all rendered pages of the cumulative outcome PDF. The
        # generator also auto-discovers sibling -2/-3/... PNGs, so keeping
        # the legacy singular field remains backward-compatible.
        # Always discover the COMPLETE cumulative table image sequence.
        # The first-page argument is only a pointer; do not let it collapse
        # the multi-page source into a single image in the config.
        "cumulativeOutcomeImage": rel(args.cumulative_image) if args.cumulative_image else None,
        "cumulativeOutcomeImages": [
            rel(str(p)) for p in sorted(
                list(Path("prepared_assets").glob("goe_cumulative-*.png"))
                or (list(Path(args.cumulative_image).parent.glob(
                    f"{Path(args.cumulative_image).stem.rsplit('-', 1)[0]}-*{Path(args.cumulative_image).suffix}"
                )) if args.cumulative_image and Path(args.cumulative_image).exists() else []),
                key=lambda p: int(re.search(r"-(\d+)\.[^.]+$", p.name).group(1))
                if re.search(r"-(\d+)\.[^.]+$", p.name) else 9999
            )
        ],
        "cumulativeOutcomeNote": "The above figures are as per IT system",
        "acknowledgement": {
            "intro": f"It has been a true pleasure to be a part of MSME Competitive (Lean) Scheme ({units_parsed.get('goeCode', '')}) as a Lean Manufacturing Consultant. Overall, it has been a mutually value adding program and we believe that we too have learnt some things under this program.",
            "msmeThanks": {"pre": "We express our heartfelt gratitude towards,", "lines": ["MSME DI team for their support and coordination throughout the program."]},
            "qciThanks": {"pre": "Our sincere thanks to,", "lines": ["QCI Team Members for their constant support & coordination throughout the program."]},
            "auditorThanks": "Our thanks to auditors from QCI, " + ", ".join(auditors) + ", for their valuable guidance and points for improvement.",
            "unitThanksPre": "This program wouldn't have been successful without the cooperation, involvement and support from the unit owners & their teams. We thank,",
            "closingLine": "We also thank Lean Coordinators of all the units and Lean team members for their active involvement & participation to make this program successful.",
            "signOff": f"{consultant_name} & Team at AB Associates",
        },
        "units": units_out,
    }
    if skipped:
        print("Skipped (no matching uploaded folder): " + ", ".join(skipped))
    return config


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("units_parsed_json")
    ap.add_argument("assets_manifest_json")
    ap.add_argument("out_json")
    ap.add_argument("--consultant-name")
    ap.add_argument("--period-start")
    ap.add_argument("--period-end")
    ap.add_argument("--location")
    ap.add_argument("--total-units", type=int)
    ap.add_argument("--auditors", help="comma-separated list of 3 auditor names")
    ap.add_argument("--cumulative-image")
    ap.add_argument("--only-uploaded", action="store_true",
                     help="Skip units listed in the .txt file that have no matching folder in the assets manifest "
                          "(use this when testing with a partial batch, e.g. 4 of 7 units).")
    ap.add_argument("--brand-assets-dir", default=None,
                     help="Folder containing watermark.png/ab_logo.png/cover_logos_band.png/footer_bar.png "
                          "(defaults to the 'assets' folder next to this script).")
    args = ap.parse_args()

    units_parsed = json.load(open(args.units_parsed_json))
    manifest = json.load(open(args.assets_manifest_json))
    config = build(units_parsed, manifest, args, args.out_json)
    json.dump(config, open(args.out_json, "w"), indent=2, ensure_ascii=False)
    print(f"Wrote {args.out_json} - {len(config['units'])} units. Text-derived profile and narrative fields were filled automatically.")
