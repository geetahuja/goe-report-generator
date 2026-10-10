#!/usr/bin/env python3
"""
parse_units_txt.py
-------------------
Parses the "Units Text Information.txt" file that comes with every GOE
data drop into structured JSON: consultant name/phone/email, GOE code,
and one record per unit (name, address, products, process, CEO/owner,
lean coordinator + contact, turnover for the last 3 FYs) - in the exact
numeric order the units are listed in the file.

This is the single biggest manual-entry killer in the whole pipeline:
per the master prompt, the .txt file is the primary source for almost
every textual field in the report, so getting it parsed automatically
means the config only needs testimonial/observations text added by hand.

Usage:
    python3 parse_units_txt.py "Units Text Information.txt" out.json
"""
import re
import sys
import json


def parse(text: str) -> dict:
    text = text.replace("\r\n", "\n")

    consultant = {"name": None, "phone": None, "email": None}
    m = re.search(r"Lean Consultant Name\s*:\s*(.+?)\s+Phone\s*:\s*(\S+)\s+Email\s*:\s*(\S+)", text)
    if m:
        consultant = {"name": m.group(1).strip(), "phone": m.group(2).strip(), "email": m.group(3).strip()}

    goe_code = None
    m = re.search(r"GOE CODE-\(([^)]+)\)", text)
    if m:
        goe_code = m.group(1).strip()

    # Split into per-unit blocks on lines that start with "<number>. <Unit Name>"
    unit_header_re = re.compile(r"^\s*(\d+)\.\s+(.+?)\s*$", re.MULTILINE)
    headers = list(unit_header_re.finditer(text))
    units = []
    for i, h in enumerate(headers):
        start = h.end()
        end = headers[i + 1].start() if i + 1 < len(headers) else len(text)
        block = text[start:end]
        num = int(h.group(1))
        name = h.group(2).strip()

        def field(label, block=block):
            mm = re.search(rf"{label}\s*:\s*(.+)", block)
            return mm.group(1).strip() if mm else None

        address = field("Address")
        products = field("Main Products")
        process = field("Key Manufacturing Process")
        year_est = field("Year of Establishment")
        # CEO/owner may also have Phone/Email on the same line.
        ceo_m = re.search(r"CEO\s*:\s*(.*?)(?=\s+Phone\s*:|\s+Email\s*:|$)", block, re.MULTILINE)
        ceo = ceo_m.group(1).strip() if ceo_m else None

        # Lean Coordinator may have Phone/Email on the same line or may be
        # followed by them on the next line. Keep the parser tolerant of
        # both layouts and never swallow the contact fields into the name.
        coord_m = re.search(
            r"Lean Coordinator\s*:\s*(.*?)(?:\s+Phone\s*:\s*(\S+))?(?:\s+Email\s*:\s*(\S+))?\s*$",
            block, re.MULTILINE
        )
        coord_name = coord_m.group(1).strip() if coord_m else None
        coord_phone = coord_m.group(2).strip() if coord_m and coord_m.group(2) else None
        coord_email = coord_m.group(3).strip() if coord_m and coord_m.group(3) else None

        turnover = []
        for tm in re.finditer(r"FY\s*([\d]{4}-[\d]{2})\s*=\s*Rs\.?\s*([\d.,]+)", block):
            turnover.append({"fy": tm.group(1), "value": tm.group(2)})

        units.append({
            "order": num,
            "name": name,
            "address": address,
            "products": products,
            "process": process,
            "yearEstablished": year_est,
            "ownerName": ceo,
            "ownerRole": f"Director, {name}",
            "leanCoordinatorName": coord_name,
            "phone": coord_phone,
            "email": coord_email,
            "turnover": turnover,
        })

    units.sort(key=lambda u: u["order"])
    return {"consultant": consultant, "goeCode": goe_code, "units": units}


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python3 parse_units_txt.py <Units Text Information.txt> [out.json]")
        sys.exit(1)
    src = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) > 2 else None
    with open(src, "r", encoding="utf-8", errors="replace") as f:
        data = parse(f.read())
    js = json.dumps(data, indent=2, ensure_ascii=False)
    if out:
        with open(out, "w", encoding="utf-8") as f:
            f.write(js)
        print(f"Parsed {len(data['units'])} units -> {out}")
    else:
        print(js)
