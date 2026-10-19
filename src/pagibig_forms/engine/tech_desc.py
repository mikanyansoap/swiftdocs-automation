"""
Technical description formatting, alignment, and short-form extraction.
Enforces Courier font, 2-inch left margin (144 pt), columnar layout, and short-form rules.
"""
from __future__ import annotations
import re
from typing import Dict, List, Tuple
from .number_to_words import area_to_words_and_figures


def parse_tech_desc_lines(raw_desc: str) -> List[Tuple[str, str]]:
    """
    Parses lines of technical description into (label, value).
    If no colon is present, label is empty and value is the line.
    """
    lines = []
    for line in raw_desc.splitlines():
        trimmed = line.strip()
        if not trimmed:
            continue
        if ":" in trimmed:
            parts = trimmed.split(":", 1)
            lbl = parts[0].strip()
            val = parts[1].strip()
            lines.append((lbl, val))
        else:
            lines.append(("", trimmed))
    return lines


def format_columnar_tech_desc(raw_desc: str, max_chars_per_line: int = 70) -> str:
    """
    If written as 'Label : value' lines, lays it out in columns:
    labels, aligned colons, values wrapping under value column.
    Keeps 'Block No. : 7    Lot No. : 14' on one row.
    """
    if not raw_desc:
        return ""

    raw_lines = [l.strip() for l in raw_desc.splitlines() if l.strip()]
    parsed = []
    max_label_len = 0

    # Special handling to keep Block No. and Lot No. on one row if consecutive
    merged_lines = []
    i = 0
    while i < len(raw_lines):
        curr = raw_lines[i]
        # Check if curr is Block No and next is Lot No
        if i + 1 < len(raw_lines):
            nxt = raw_lines[i + 1]
            if re.match(r"(?i)^block\s*(?:no\.?)?\s*:", curr) and re.match(r"(?i)^lot\s*(?:no\.?)?\s*:", nxt):
                merged_lines.append(f"{curr}    {nxt}")
                i += 2
                continue
        merged_lines.append(curr)
        i += 1

    formatted_rows = []
    for line in merged_lines:
        if ":" in line and not line.count(":") > 1:
            lbl, val = line.split(":", 1)
            lbl = lbl.strip()
            val = val.strip()
            max_label_len = max(max_label_len, len(lbl))
            parsed.append(("pair", lbl, val))
        else:
            parsed.append(("raw", line, ""))

    col_align = min(24, max(12, max_label_len + 1))
    val_width = max(30, max_chars_per_line - col_align - 2)

    for item in parsed:
        kind = item[0]
        if kind == "pair":
            lbl, val = item[1], item[2]
            header = f"{lbl.ljust(col_align)}: "
            indent = " " * len(header)
            # Wrap value if needed
            words = val.split()
            cur_line = ""
            wrapped = []
            for w in words:
                if len(cur_line) + len(w) + 1 <= val_width:
                    cur_line = f"{cur_line} {w}".strip()
                else:
                    if cur_line:
                        wrapped.append(cur_line)
                    cur_line = w
            if cur_line:
                wrapped.append(cur_line)

            if not wrapped:
                formatted_rows.append(header.rstrip())
            else:
                formatted_rows.append(header + wrapped[0])
                for extra in wrapped[1:]:
                    formatted_rows.append(indent + extra)
        else:
            formatted_rows.append(item[1])

    return "\n".join(formatted_rows)


def extract_short_tech_desc(
    full_desc: str,
    block_no: str,
    lot_no: str,
    lot_area: float = 36.0,
    max_chars: int = 65
) -> str:
    """
    Extracts the short form for Deed of Absolute Sale (HLF-525) and both LMAs (HLF-162 / HLF-163):
      - Survey Plan No.
      - Block No. + Lot No. on one row
      - Portion of Lot
      - LRC Record No.
      - Area: from the description, or built from Lot Area as 'THIRTY SIX (36) SQUARE METERS more or less.'
    """
    lines = full_desc.splitlines() if full_desc else []
    extracted = {
        "survey_plan": "",
        "portion_of_lot": "",
        "lrc_record": "",
        "area_text": "",
    }

    for line in lines:
        l = line.strip()
        if re.search(r"(?i)survey\s*plan\s*(?:no\.?)?\s*:", l):
            extracted["survey_plan"] = l.split(":", 1)[1].strip()
        elif re.search(r"(?i)portion\s*of\s*lot\s*:", l):
            extracted["portion_of_lot"] = l.split(":", 1)[1].strip()
        elif re.search(r"(?i)lrc\s*record\s*(?:no\.?)?\s*:", l):
            extracted["lrc_record"] = l.split(":", 1)[1].strip()
        elif re.search(r"(?i)area\s*:", l):
            extracted["area_text"] = l.split(":", 1)[1].strip()

    # Anything missing from the description stays blank: never print another lot's data
    survey_plan = extracted["survey_plan"]
    portion_of_lot = extracted["portion_of_lot"]
    if not portion_of_lot and block_no and lot_no:
        portion_of_lot = f"BLOCK {block_no}, LOT {lot_no}"
    lrc_record = extracted["lrc_record"]

    area_str = extracted["area_text"]
    if not area_str and lot_area and lot_area > 0:
        area_str = area_to_words_and_figures(lot_area)

    short_lines = [
        f"Survey Plan No. : {survey_plan}",
        f"Block No. : {block_no}    Lot No. : {lot_no}",
        f"Portion of Lot  : {portion_of_lot}",
        f"LRC Record No.  : {lrc_record}",
        f"Area            : {area_str}",
    ]
    return "\n".join(line.rstrip() for line in short_lines)
