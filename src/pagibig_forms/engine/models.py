"""
Data models and normalization logic for Pag-IBIG Forms.
Supports columns A-O fixed by position, extra dynamic columns, and optional fields.
"""
from __future__ import annotations
import datetime
import re
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple

FIXED_COLUMNS = [
    "Principal Borrower",     # A (0)
    "Status",                 # B (1)
    "Co Borrower",            # C (2)
    "Address",                # D (3)
    "Block & Lot",            # E (4)
    "Title Number",           # F (5)
    "Technical Description",  # G (6)
    "Interest rate",          # H (7)
    "Consideration",          # I (8)
    "Loanable amount",        # J (9)
    "Technical description (continued)", # K (10)
    "Notary of CTS",          # L (11)
    "TIN principal",          # M (12)
    "TIN co-borrower",        # N (13)
    "TIN spouse",             # O (14)
]

STANDARD_EXTRA_COLUMNS = [
    "Date of Notarization",
    "Loan Term (years)",
    "Re-pricing (years)",
    "MRI/SRI",
    "Non-Life",
    "Principal & Interest",
    "Attorney-in-Fact",
    "Co Attorney-in-Fact",
    "Attorney-in-Fact Relationship",
    "House Type",
    "Lot Area (sqm)",
    "Co Borrower status",
    "Co Borrower address",
    "Email",
]

# Compound surname particles that must stay with the surname (matched as whole words)
COMPOUND_PARTICLES = (
    ("VDA.", "DE"), ("VDA", "DE"),
    ("DE", "LA"), ("DE", "LOS"), ("DE", "LAS"),
    ("DELA",), ("DELOS",), ("DELAS",), ("DEL",), ("DE",),
    ("SAN",), ("SANTA",), ("SANTO",),
)
NAME_SUFFIXES = ("JR.", "JR", "SR.", "SR", "III", "II", "IV", "V")


def extract_surname(full_name: str) -> str:
    """Extract surname keeping compound Filipino surnames together.

    Only the word right before the last word is checked for a particle, so a
    middle name such as 'DEL CARMEN' is not taken as part of the surname.
    """
    if not full_name:
        return ""
    words = full_name.strip().split()
    suffix = ""
    if len(words) >= 3 and words[-1].upper().rstrip(",") in NAME_SUFFIXES:
        suffix = words.pop()
        words[-1] = words[-1].rstrip(",")
    if len(words) < 2:
        return " ".join(words + ([suffix] if suffix else []))

    upper = [w.upper() for w in words]

    def with_particles(idx: int) -> int:
        # Walk back over particles directly in front of word idx (e.g. VDA. DE LA CRUZ)
        for particles in COMPOUND_PARTICLES:
            n = len(particles)
            if idx - n >= 1 and tuple(upper[idx - n:idx]) == particles:
                return idx - n
        return idx

    start = with_particles(len(words) - 1)
    # 'PATERNAL Y MATERNAL' surnames (e.g. DELOS SANTOS Y SEVERINO) stay together
    if start - 2 >= 1 and upper[start - 1] == "Y":
        start = with_particles(start - 2)
    surname = " ".join(words[start:])
    return f"{surname} {suffix}" if suffix else surname


def parse_marital_status(status_str: str) -> Tuple[str, str]:
    """
    Parses Status string.
    Returns (status_label, spouse_name).
    e.g. 'Single' -> ('Single', '')
    'Married to MARIA DELA CRUZ' -> ('Married', 'MARIA DELA CRUZ')
    """
    if not status_str:
        return "Single", ""
    cleaned = status_str.strip()
    m = re.match(r"(?i)married\s+to\s+(.+)", cleaned)
    if m:
        return "Married", m.group(1).strip()
    if re.match(r"(?i)married", cleaned):
        return "Married", ""
    return cleaned, ""


def parse_block_lot(bl_str: str) -> Tuple[str, str]:
    """
    Parses 'BLOCK 7 LOT 14', 'BLK 7, LT 14' or 'B7 L14' into (block_no, lot_no).
    Keywords must start a word, so the 'L' inside 'BLOCK' is never read as 'Lot'.
    """
    if not bl_str:
        return "", ""
    block_no = ""
    lot_no = ""
    bm = re.search(r"(?i)\b(?:block|blk|b)\.?\s*(?:no\.?\s*)?([0-9][0-9a-zA-Z-]*)", bl_str)
    if bm:
        block_no = bm.group(1)
    lm = re.search(r"(?i)\b(?:lot|lt|l)\.?\s*(?:no\.?\s*)?([0-9][0-9a-zA-Z-]*)", bl_str)
    if lm:
        lot_no = lm.group(1)
    return block_no, lot_no


def parse_float(val: Any, default: float = 0.0) -> float:
    if val is None or val == "":
        return default
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).replace(",", "").replace("₱", "").replace("P", "").strip()
    if s.endswith("%"):
        try:
            return float(s[:-1].strip()) / 100.0
        except ValueError:
            return default
    try:
        return float(s)
    except ValueError:
        return default


DATE_FORMATS = ("%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%B %d, %Y", "%b %d, %Y", "%d %B %Y")


def parse_date(val: Any) -> Optional[datetime.date]:
    """Reads an Excel date cell or a typed date (mm/dd/yyyy, yyyy-mm-dd, 'October 16, 2026')."""
    if val is None or val == "":
        return None
    if isinstance(val, datetime.datetime):
        return val.date()
    if isinstance(val, datetime.date):
        return val
    s = str(val).strip()
    for fmt in DATE_FORMATS:
        try:
            return datetime.datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def date_to_mdy(val: Any) -> str:
    """Date as mm/dd/yyyy for editing; unreadable text is kept as typed."""
    d = parse_date(val)
    if d:
        return d.strftime("%m/%d/%Y")
    return "" if val is None else str(val).strip()


def date_to_long(val: Any) -> str:
    """Date as 'OCTOBER 16, 2026' for the forms; unreadable text is kept as typed."""
    d = parse_date(val)
    if d:
        return f"{d.strftime('%B').upper()} {d.day}, {d.year}"
    return "" if val is None else str(val).strip().upper()


def parse_int(val: Any, default: int = 0) -> int:
    f = parse_float(val, float(default))
    return int(round(f))


@dataclass
class BuyerRecord:
    row_idx: int = 0  # 1-based row in Excel
    loaded_name: str = ""  # principal name as last read from / saved to Excel (to detect renames)
    principal_borrower: str = ""
    status: str = "Single"
    co_borrower: str = "N/A"
    address: str = ""
    block_lot: str = ""
    title_no: str = ""
    tech_desc: str = ""
    interest_rate: float = 0.0625  # default 6.25%
    consideration: float = 0.0
    loanable_amount: float = 0.0
    tech_desc_cont: str = ""
    notary_cts: str = ""
    tin_principal: str = ""
    tin_co_borrower: str = ""
    tin_spouse: str = ""

    # Extra columns
    date_notarization: str = ""  # mm/dd/yyyy
    loan_term_years: int = 30
    repricing_years: int = 1
    mri_sri: float = 0.0
    non_life: float = 0.0
    principal_interest: float = 0.0
    attorney_in_fact: str = ""
    co_attorney_in_fact: str = ""
    attorney_in_fact_rel: str = ""
    house_type: str = "TOWNHOUSE"
    lot_area: float = 36.0
    co_borrower_status: str = ""
    co_borrower_address: str = ""
    email: str = ""

    # Optional extra fields
    mid: str = ""
    birthday: str = ""
    sex: str = ""
    maiden_name: str = ""
    employer: str = ""
    employer_address: str = ""
    employer_tel: str = ""
    employer_id: str = ""
    contact_no: str = ""

    co_mid: str = ""
    co_birthday: str = ""
    co_sex: str = ""
    co_employer: str = ""
    co_contact_no: str = ""

    extra_fields: Dict[str, Any] = field(default_factory=dict)

    @property
    def surname(self) -> str:
        return extract_surname(self.principal_borrower)

    @property
    def co_surname(self) -> str:
        if not self.has_co_borrower:
            return ""
        return extract_surname(self.co_borrower)

    @property
    def has_co_borrower(self) -> bool:
        c = self.co_borrower.strip().upper()
        return bool(c and c != "N/A" and c != "NONE")

    @property
    def has_attorney_in_fact(self) -> bool:
        a = self.attorney_in_fact.strip().upper()
        return bool(a and a != "N/A" and a != "NONE")

    @property
    def marital_status(self) -> str:
        status, _ = parse_marital_status(self.status)
        return status

    @property
    def spouse_name(self) -> str:
        _, spouse = parse_marital_status(self.status)
        return spouse

    @property
    def co_marital_status(self) -> str:
        if not self.co_borrower_status.strip():
            return ""
        status, _ = parse_marital_status(self.co_borrower_status)
        return status

    @property
    def co_spouse_name(self) -> str:
        _, spouse = parse_marital_status(self.co_borrower_status)
        return spouse

    @property
    def block_no(self) -> str:
        b, _ = parse_block_lot(self.block_lot)
        return b

    @property
    def lot_no(self) -> str:
        _, l = parse_block_lot(self.block_lot)
        return l

    @property
    def lot_label(self) -> str:
        """'BLOCK 7 LOT 14', or the Block & Lot text as typed if it cannot be read."""
        if self.block_no and self.lot_no:
            return f"BLOCK {self.block_no} LOT {self.lot_no}"
        return self.block_lot.strip().upper()

    @property
    def equity(self) -> float:
        return max(0.0, self.consideration - self.loanable_amount)

    @property
    def is_affordable(self) -> bool:
        """Loanable amount below ₱950,000 -> Affordable; ₱950,000 and up -> End-User."""
        return self.loanable_amount < 950000.0

    @property
    def stage_program_name(self) -> str:
        return "Affordable Housing Program" if self.is_affordable else "End-User Home Financing Program"

    @property
    def full_tech_desc(self) -> str:
        p1 = (self.tech_desc or "").strip()
        p2 = (self.tech_desc_cont or "").strip()
        if p1 and p2:
            return f"{p1}\n{p2}"
        return p1 or p2

    @property
    def resolved_co_address(self) -> str:
        if self.co_borrower_address.strip():
            return self.co_borrower_address.strip()
        return self.address.strip()
