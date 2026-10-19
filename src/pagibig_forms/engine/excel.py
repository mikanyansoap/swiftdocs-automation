"""
Excel reader and atomic writer for BookData.xlsx.
Complies with Excel integrity rules:
- Reload newest bytes before save
- Verify target row matches principal name or find/append
- Atomic replace using temporary file in same directory
- File lock detection: 'Close the Excel file and try again'
- Never delete rows; convert numbers and dates to real Excel types
"""
from __future__ import annotations
import os
import shutil
import tempfile
import time
from typing import Any, Dict, List, Optional, Tuple

from .models import (
    FIXED_COLUMNS,
    STANDARD_EXTRA_COLUMNS,
    BuyerRecord,
    parse_float,
    parse_int,
    parse_date,
    date_to_mdy
)

# Fields stored as real Excel dates
DATE_FIELDS = ("date_notarization", "birthday")


def _openpyxl():
    """openpyxl, loaded on first use: importing it takes about 2 seconds, which would
    otherwise delay the app window from appearing."""
    import openpyxl
    import openpyxl.styles
    import openpyxl.utils
    return openpyxl

# Header matching variants (case-insensitive)
HEADER_VARIANTS = {
    "date_notarization": ["date of notarization", "cts date", "date notarized", "notarization date"],
    "loan_term_years": ["loan term (years)", "loan term", "term (years)", "term"],
    "repricing_years": ["re-pricing (years)", "repricing (years)", "repricing", "re-pricing", "fixing period"],
    "mri_sri": ["mri/sri", "mri", "sri", "mri_sri"],
    "non_life": ["non-life", "nonlife", "fire insurance", "non life"],
    "principal_interest": ["principal & interest", "p&i", "principal and interest", "monthly amortization"],
    "attorney_in_fact": ["attorney-in-fact", "attorney in fact", "aif", "spa"],
    "co_attorney_in_fact": ["co attorney-in-fact", "co attorney in fact", "co aif", "co-aif"],
    "attorney_in_fact_rel": ["attorney-in-fact relationship", "aif relationship", "relationship"],
    "house_type": ["house type", "model", "unit type"],
    "lot_area": ["lot area (sqm)", "lot area", "lot size"],
    "co_borrower_status": ["co borrower status", "co-borrower status", "co status"],
    "co_borrower_address": ["co borrower address", "co-borrower address", "co address"],
    "email": ["email", "e-mail", "email address"],
    "mid": ["mid", "pag-ibig mid", "pagibig mid", "hdmf no"],
    "birthday": ["birthday", "birth date", "dob"],
    "sex": ["sex", "gender"],
    "maiden_name": ["maiden name"],
    "employer": ["employer", "company"],
    "employer_address": ["employer address", "company address"],
    "employer_tel": ["employer tel", "employer telephone", "company tel"],
    "employer_id": ["employer id"],
    "contact_no": ["contact no", "contact number", "mobile", "mobile no", "phone"],
}


class ExcelManager:
    def __init__(self, excel_path: str):
        self.excel_path = excel_path

    def load_buyers(self) -> List[BuyerRecord]:
        """Loads all buyer records from BookData.xlsx."""
        if not os.path.exists(self.excel_path):
            return []

        # Read-only mode streams the rows instead of building the whole workbook in memory,
        # which keeps loading fast as BookData.xlsx grows.
        wb = _openpyxl().load_workbook(self.excel_path, data_only=True, read_only=True)
        try:
            ws = wb.active
            if not ws:
                return []
            rows = ws.iter_rows(values_only=True)
            header = next(rows, ())
            header_map = self._map_header_values(header)
            buyers = [record for row_idx, values in enumerate(rows, start=2)
                      if (record := self._record_from_row(row_idx, values, header_map)) is not None]
        finally:
            wb.close()
        return buyers

    @staticmethod
    def _record_from_row(row_idx: int, values: tuple, header_map: Dict[str, int]) -> Optional[BuyerRecord]:
        def cell(col: int):  # 1-based column, like ws.cell
            return values[col - 1] if col - 1 < len(values) else None

        principal_val = cell(1)
        if not principal_val or not str(principal_val).strip():
            return None

        record = BuyerRecord(
            row_idx=row_idx,
            loaded_name=str(principal_val).strip(),
            principal_borrower=str(principal_val).strip(),
            status=str(cell(2) or "Single").strip(),
            co_borrower=str(cell(3) or "N/A").strip(),
            address=str(cell(4) or "").strip(),
            block_lot=str(cell(5) or "").strip(),
            title_no=str(cell(6) or "").strip(),
            tech_desc=str(cell(7) or "").strip(),
            interest_rate=parse_float(cell(8), 0.0625),
            consideration=parse_float(cell(9), 0.0),
            loanable_amount=parse_float(cell(10), 0.0),
            tech_desc_cont=str(cell(11) or "").strip(),
            notary_cts=str(cell(12) or "").strip(),
            tin_principal=str(cell(13) or "").strip(),
            tin_co_borrower=str(cell(14) or "").strip(),
            tin_spouse=str(cell(15) or "").strip(),
        )

        # Map dynamic extra columns
        for field_name, col_idx in header_map.items():
            val = cell(col_idx)
            if val is not None:
                if field_name in ("loan_term_years", "repricing_years"):
                    setattr(record, field_name, parse_int(val, 30 if "term" in field_name else 1))
                elif field_name in ("mri_sri", "non_life", "principal_interest", "lot_area"):
                    setattr(record, field_name, parse_float(val, 0.0))
                elif field_name in DATE_FIELDS:
                    setattr(record, field_name, date_to_mdy(val))
                elif hasattr(record, field_name):
                    setattr(record, field_name, str(val).strip())
                else:
                    record.extra_fields[field_name] = val
        return record

    def save_buyer(self, buyer: BuyerRecord) -> int:
        """
        Saves a buyer record to Excel:
        - Reloads the newest file bytes before every save.
        - Verifies target row still holds the same principal name; otherwise finds buyer by name, or appends.
        - Writes to a temp file in the same folder, then does an atomic replace.
        - If locked, raises PermissionError ('Close the Excel file and try again').
        Returns the saved row index.
        """
        if not os.path.exists(self.excel_path):
            self.create_empty_workbook(self.excel_path)

        # Check file lock before attempting
        try:
            with open(self.excel_path, "a+b"):
                pass
        except (IOError, PermissionError):
            raise PermissionError("Close the Excel file and try again")

        wb = _openpyxl().load_workbook(self.excel_path)
        ws = wb.active
        header_map = self._ensure_headers(ws)

        target_row = buyer.row_idx
        # Verify the target row still holds this buyer. Match on the name as it was
        # loaded too, so renaming a buyer updates their row instead of adding a copy.
        names = {buyer.principal_borrower.strip().upper()}
        if buyer.loaded_name.strip():
            names.add(buyer.loaded_name.strip().upper())
        if target_row > 1 and target_row <= ws.max_row:
            existing_name = ws.cell(row=target_row, column=1).value
            if existing_name and str(existing_name).strip().upper() not in names:
                # Rows moved (sorted/inserted): search by name
                target_row = self._find_buyer_row(ws, names)

        if target_row <= 1 or target_row > ws.max_row:
            # Append as new row
            target_row = ws.max_row + 1

        # Write Columns A-O
        ws.cell(row=target_row, column=1, value=buyer.principal_borrower.strip())
        ws.cell(row=target_row, column=2, value=buyer.status.strip())
        ws.cell(row=target_row, column=3, value=buyer.co_borrower.strip())
        ws.cell(row=target_row, column=4, value=buyer.address.strip())
        ws.cell(row=target_row, column=5, value=buyer.block_lot.strip())
        ws.cell(row=target_row, column=6, value=buyer.title_no.strip())
        ws.cell(row=target_row, column=7, value=buyer.tech_desc.strip())
        ws.cell(row=target_row, column=8, value=float(buyer.interest_rate))
        ws.cell(row=target_row, column=9, value=float(buyer.consideration))
        ws.cell(row=target_row, column=10, value=float(buyer.loanable_amount))
        ws.cell(row=target_row, column=11, value=buyer.tech_desc_cont.strip())
        ws.cell(row=target_row, column=12, value=buyer.notary_cts.strip())
        ws.cell(row=target_row, column=13, value=buyer.tin_principal.strip())
        ws.cell(row=target_row, column=14, value=buyer.tin_co_borrower.strip())
        ws.cell(row=target_row, column=15, value=buyer.tin_spouse.strip())

        # Write Extra Columns
        for field_name, col_idx in header_map.items():
            val = getattr(buyer, field_name, None)
            if val is not None:
                if field_name in DATE_FIELDS:
                    d = parse_date(val)
                    cell = ws.cell(row=target_row, column=col_idx, value=d if d else (str(val).strip() or None))
                    if d:
                        cell.number_format = "mm/dd/yyyy"
                elif isinstance(val, float):
                    ws.cell(row=target_row, column=col_idx, value=float(val))
                elif isinstance(val, int):
                    ws.cell(row=target_row, column=col_idx, value=int(val))
                else:
                    ws.cell(row=target_row, column=col_idx, value=str(val).strip())

        # Save to temp file in same folder and perform atomic replace
        folder = os.path.dirname(os.path.abspath(self.excel_path))
        temp_file = os.path.join(folder, f".temp_{int(time.time()*1000)}.xlsx")

        try:
            wb.save(temp_file)
            wb.close()
            # Atomic replace
            os.replace(temp_file, self.excel_path)
        except BaseException as e:
            if os.path.exists(temp_file):
                try:
                    os.remove(temp_file)
                except Exception:
                    pass
            if isinstance(e, (IOError, PermissionError)):
                raise PermissionError("Close the Excel file and try again") from e
            raise

        buyer.row_idx = target_row
        buyer.loaded_name = buyer.principal_borrower.strip()
        return target_row

    def _find_buyer_row(self, ws, names) -> int:
        for r in range(2, ws.max_row + 1):
            val = ws.cell(row=r, column=1).value
            if val and str(val).strip().upper() in names:
                return r
        return ws.max_row + 1

    def _map_headers(self, ws) -> Dict[str, int]:
        return self._map_header_values([ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)])

    @staticmethod
    def _map_header_values(header: tuple) -> Dict[str, int]:
        """Maps extra columns (P onwards) to buyer fields by their header text."""
        mapping = {}
        for col_idx in range(16, len(header) + 1):
            header_val = header[col_idx - 1]
            if not header_val:
                continue
            header_str = str(header_val).strip().lower()
            for key, variants in HEADER_VARIANTS.items():
                if header_str in variants:
                    mapping[key] = col_idx
                    break
        return mapping

    def _ensure_headers(self, ws) -> Dict[str, int]:
        """Ensures all standard extra headers exist; appends missing headers."""
        mapping = self._map_headers(ws)
        col_cursor = max(15, ws.max_column)

        for std_header in STANDARD_EXTRA_COLUMNS:
            std_clean = std_header.lower()
            key_match = None
            for key, variants in HEADER_VARIANTS.items():
                if std_clean in variants:
                    key_match = key
                    break
            if key_match and key_match not in mapping:
                col_cursor += 1
                cell = ws.cell(row=1, column=col_cursor, value=std_header)
                cell.font = _openpyxl().styles.Font(bold=True)
                mapping[key_match] = col_cursor

        return mapping

    @staticmethod
    def _new_workbook_with_headers():
        wb = _openpyxl().Workbook()
        ws = wb.active
        ws.title = "Buyers"

        # Headers A - O
        headers = list(FIXED_COLUMNS) + list(STANDARD_EXTRA_COLUMNS)
        header_fill = _openpyxl().styles.PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
        header_font = _openpyxl().styles.Font(color="FFFFFF", bold=True, name="Calibri", size=11)

        for col_idx, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_idx, value=h)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = _openpyxl().styles.Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws.column_dimensions[_openpyxl().utils.get_column_letter(col_idx)].width = max(12, len(h) + 2)
        return wb, ws

    @staticmethod
    def create_empty_workbook(file_path: str):
        """Creates a new BookData.xlsx with the column headers only (no buyers)."""
        wb, _ws = ExcelManager._new_workbook_with_headers()
        os.makedirs(os.path.dirname(os.path.abspath(file_path)), exist_ok=True)
        wb.save(file_path)
        wb.close()

    @staticmethod
    def create_sample_workbook(file_path: str):
        """Creates a BookData.xlsx with fictional sample buyers, for tests only."""
        wb, ws = ExcelManager._new_workbook_with_headers()

        # Sample rows for verification
        sample_rows = [
            # Row 1: Single Buyer (Affordable loan < 950k)
            [
                "JUAN DELA CRUZ", "Single", "N/A", "123 SAMPAGUITA ST., MORONG, RIZAL",
                "BLOCK 7 LOT 14", "T-892134",
                "Survey Plan No. : PCS-04-030177\nPortion of Lot : BLOCK 7, LOT 14\nLRC Record No. : N-17185",
                0.0625, 1150000, 850000,
                "Area : THIRTY SIX (36) SQUARE METERS more or less.",
                "Doc 10, Page 2, Book 11, Series of 2026",
                "123-456-789-000", "N/A", "N/A",
                "10/16/2026", 30, 1, 350.0, 220.0, 5230.50,
                "", "", "", "TOWNHOUSE", 36.0, "", "", "juan.delacruz@example.com"
            ],
            # Row 2: Married Buyer with Married Co-Borrower (End-User loan >= 950k)
            [
                "ROBERTO SANTOS", "Married to MARIA SANTOS", "EDUARDO DELOS REYES",
                "456 RIZAL AVE., MAYBANCAL, MORONG, RIZAL", "BLOCK 12 LOT 5", "T-994321",
                "Survey Plan No. : PCS-04-030177\nPortion of Lot : BLOCK 12, LOT 5\nLRC Record No. : N-17185",
                0.0625, 1600000, 1250000,
                "Area : FORTY (40) SQUARE METERS more or less.",
                "Doc 45, Page 9, Book 12, Series of 2026",
                "234-567-890-000", "345-678-901-000", "456-789-012-000",
                "10/16/2026", 30, 3, 420.0, 280.0, 7700.25,
                "", "", "", "TWO-STOREY TOWNHOUSE", 40.0, "Married to CARMEN DELOS REYES",
                "456 RIZAL AVE., MAYBANCAL, MORONG, RIZAL", "roberto.santos@example.com"
            ],
            # Row 3: Buyer with Attorney-in-Fact
            [
                "ANNA MARIE SAN JOSE", "Single", "N/A", "789 QUEZON BLVD., MORONG, RIZAL",
                "BLOCK 3 LOT 9", "T-772156",
                "Survey Plan No. : PCS-04-030177\nPortion of Lot : BLOCK 3, LOT 9\nLRC Record No. : N-17185",
                0.0625, 1200000, 949999,
                "Area : THIRTY SIX (36) SQUARE METERS more or less.",
                "Doc 12, Page 3, Book 11, Series of 2026",
                "567-890-123-000", "N/A", "N/A",
                "10/16/2026", 25, 1, 380.0, 240.0, 5840.10,
                "CARLITO SAN JOSE", "", "BROTHER", "TOWNHOUSE", 36.0, "", "", "anna.sanjose@example.com"
            ],
            # Row 4: Boundary test (Loanable amount 950,000 exact -> End-User) with long technical description
            [
                "CRISTINA DE CASTRO", "Married to FERDINAND DE CASTRO", "N/A",
                "UNIT 501 TOWER B, RESIDENCES AT MORONG, BRGY. MAYBANCAL, MORONG, RIZAL",
                "BLOCK 15 LOT 22", "T-883344",
                "Survey Plan No. : PCS-04-030177\nBlock No. : 15    Lot No. : 22\nPortion of Lot : BLOCK 15, LOT 22\nLRC Record No. : N-17185\nBoundaries : NE Lot 21, SE Road Lot 4, SW Lot 23, NW Lot 10\nBeginning at a point marked 1 on plan being N 45 deg 30 min E 2500.50 m from BLLM No. 1 Morong",
                0.0625, 1400000, 950000,
                "Area : FORTY TWO (42) SQUARE METERS more or less.",
                "Doc 88, Page 17, Book 14, Series of 2026",
                "678-901-234-000", "N/A", "789-012-345-000",
                "10/16/2026", 30, 3, 400.0, 260.0, 5850.75,
                "", "", "", "TOWNHOUSE", 42.0, "", "", "cristina.decastro@example.com"
            ]
        ]

        for r_idx, row_data in enumerate(sample_rows, 2):
            for c_idx, val in enumerate(row_data, 1):
                ws.cell(row=r_idx, column=c_idx, value=val)

        # Set column widths nicely
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = _openpyxl().utils.get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = min(35, max(12, max_len + 2))

        os.makedirs(os.path.dirname(os.path.abspath(file_path)), exist_ok=True)
        wb.save(file_path)
        wb.close()
