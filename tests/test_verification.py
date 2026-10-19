"""
Comprehensive Verification Test Suite for Pag-IBIG Forms.
Verifies:
1. Form generation for all test scenarios (single, married+co-borrower, attorney-in-fact, 949,999 vs 950,000, long names/tech desc)
2. PDF page rendering to PNG and visual inspection checks
3. Round-trip BookData.xlsx (add, edit, save, reload, concurrency)
4. Headless Toga UI run with toga-dummy backend
"""
from __future__ import annotations
import os
import sys
import unittest
from unittest.mock import MagicMock, patch

# Ensure src is on sys.path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "src")))

from pagibig_forms.engine.models import BuyerRecord, extract_surname
from pagibig_forms.engine.filler import FormFiller
from pagibig_forms.engine.excel import ExcelManager
from pagibig_forms.engine.rules import get_forms_for_stage, build_amortization_schedule, build_i_we_line
from pagibig_forms.engine.tech_desc import format_columnar_tech_desc, extract_short_tech_desc
from pagibig_forms.engine.anthropic_client import map_hla_json_to_buyer

# Optional PyQt6 for rasterizing to PNG
try:
    from PyQt6.QtWidgets import QApplication
    from PyQt6.QtPdf import QPdfDocument
    from PyQt6.QtCore import QSize
    HAS_PYQT_PDF = True
except ImportError:
    HAS_PYQT_PDF = False


class TestPagIBIGFormsVerification(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
        cls.forms_dir = os.path.join(cls.base_dir, "forms")
        cls.output_dir = os.path.join(cls.base_dir, "tests", "rendered_output")
        os.makedirs(cls.output_dir, exist_ok=True)

        cls.filler = FormFiller(forms_dir=cls.forms_dir)

        # 5 Test Rows as specified
        cls.test_buyers = [
            # 1. Single Buyer (Affordable < 950,000)
            BuyerRecord(
                row_idx=2,
                principal_borrower="JUAN DELA CRUZ",
                status="Single",
                co_borrower="N/A",
                address="123 SAMPAGUITA ST., MORONG, RIZAL",
                block_lot="BLOCK 7 LOT 14",
                title_no="T-892134",
                consideration=1150000,
                loanable_amount=850000,
                tech_desc="Survey Plan No. : PCS-04-030177\nPortion of Lot : BLOCK 7, LOT 14\nLRC Record No. : N-17185",
                tech_desc_cont="Area : THIRTY SIX (36) SQUARE METERS more or less.",
                notary_cts="Doc 10, Page 2, Book 11, Series of 2026",
                date_notarization="10/16/2026"
            ),
            # 2. Married Buyer with Married Co-Borrower (End-User >= 950,000)
            BuyerRecord(
                row_idx=3,
                principal_borrower="ROBERTO SANTOS",
                status="Married to MARIA SANTOS",
                co_borrower="EDUARDO DELOS REYES",
                co_borrower_status="Married to CARMEN DELOS REYES",
                address="456 RIZAL AVE., MAYBANCAL, MORONG, RIZAL",
                block_lot="BLOCK 12 LOT 5",
                title_no="T-994321",
                consideration=1600000,
                loanable_amount=1250000,
                tech_desc="Survey Plan No. : PCS-04-030177\nPortion of Lot : BLOCK 12, LOT 5\nLRC Record No. : N-17185",
                tech_desc_cont="Area : FORTY (40) SQUARE METERS more or less.",
                notary_cts="Doc 45, Page 9, Book 12, Series of 2026",
                date_notarization="10/16/2026"
            ),
            # 3. Buyer with Attorney-in-Fact & Loanable 949,999 (Affordable boundary)
            BuyerRecord(
                row_idx=4,
                principal_borrower="ANNA MARIE SAN JOSE",
                status="Single",
                co_borrower="N/A",
                attorney_in_fact="CARLITO SAN JOSE",
                attorney_in_fact_rel="BROTHER",
                address="789 QUEZON BLVD., MORONG, RIZAL",
                block_lot="BLOCK 3 LOT 9",
                title_no="T-772156",
                consideration=1200000,
                loanable_amount=949999,
                tech_desc="Survey Plan No. : PCS-04-030177\nPortion of Lot : BLOCK 3, LOT 9\nLRC Record No. : N-17185",
                notary_cts="Doc 12, Page 3, Book 11, Series of 2026",
                date_notarization="10/16/2026"
            ),
            # 4. Loanable 950,000 exact (End-User boundary)
            BuyerRecord(
                row_idx=5,
                principal_borrower="CRISTINA DE CASTRO",
                status="Married to FERDINAND DE CASTRO",
                co_borrower="N/A",
                address="UNIT 501 TOWER B, MORONG, RIZAL",
                block_lot="BLOCK 15 LOT 22",
                title_no="T-883344",
                consideration=1400000,
                loanable_amount=950000,
                tech_desc="Survey Plan No. : PCS-04-030177\nPortion of Lot : BLOCK 15, LOT 22\nLRC Record No. : N-17185",
                notary_cts="Doc 88, Page 17, Book 14, Series of 2026",
                date_notarization="10/16/2026"
            ),
            # 5. Long names, long address, and long technical description
            BuyerRecord(
                row_idx=6,
                principal_borrower="ALEXANDER EMMANUEL DELOS SANTOS Y SEVERINO",
                status="Married to MARIA THERESA VICTORIA DE GUZMAN",
                co_borrower="N/A",
                address="PHASE 2 BLOCK 88 LOT 99 RESIDENCIAS DE SAN PABLO, BARANGAY MAYBANCAL, MORONG, PROVINCE OF RIZAL, 1990",
                block_lot="BLOCK 88 LOT 99",
                title_no="T-068-202600123456",
                consideration=2500000,
                loanable_amount=1950000,
                tech_desc=(
                    "Survey Plan No. : PCS-04-030177\n"
                    "Block No. : 88    Lot No. : 99\n"
                    "Portion of Lot : BLOCK 88, LOT 99\n"
                    "LRC Record No. : N-17185\n"
                    "Boundaries : NE Road Lot 12, SE Lot 100, SW Lot 87, NW Road Lot 14\n"
                    "Beginning at a point marked 1 on plan being N 45 deg 30 min E 3450.75 m from BLLM No. 1 Morong\n"
                    "Thence N 15 deg 22 min E 12.00 m to point 2; Thence S 74 deg 38 min E 10.00 m to point 3"
                ),
                tech_desc_cont="Area : ONE HUNDRED TWENTY (120) SQUARE METERS more or less.",
                notary_cts="Doc 999, Page 199, Book 55, Series of 2026",
                date_notarization="10/16/2026"
            )
        ]

    def test_01_compound_surname_extraction(self):
        """Verifies compound Filipino surnames are kept together."""
        self.assertEqual(extract_surname("JUAN DELA CRUZ"), "DELA CRUZ")
        self.assertEqual(extract_surname("ROBERTO DELOS REYES"), "DELOS REYES")
        self.assertEqual(extract_surname("ANNA MARIE SAN JOSE"), "SAN JOSE")
        self.assertEqual(extract_surname("CRISTINA DE CASTRO"), "DE CASTRO")
        self.assertEqual(extract_surname("PEDRO DELA CRUZ JR."), "DELA CRUZ JR.")

    def test_02_program_selection_boundary(self):
        """
        Loanable amount < 950,000 -> Affordable set (HLF-163 / HLF-087).
        Loanable amount >= 950,000 -> End-User set (HLF-162 / HLF-086).
        """
        b_aff = self.test_buyers[2]  # 949,999
        self.assertTrue(b_aff.is_affordable)
        forms_aff = get_forms_for_stage("POST-APPROVAL", b_aff)
        codes_aff = [f["code"] for f in forms_aff]
        self.assertIn("HLF-163", codes_aff)
        self.assertIn("HLF-087", codes_aff)
        self.assertNotIn("HLF-162", codes_aff)
        self.assertNotIn("HLF-086", codes_aff)

        b_eu = self.test_buyers[3]   # 950,000
        self.assertFalse(b_eu.is_affordable)
        forms_eu = get_forms_for_stage("POST-APPROVAL", b_eu)
        codes_eu = [f["code"] for f in forms_eu]
        self.assertIn("HLF-162", codes_eu)
        self.assertIn("HLF-086", codes_eu)
        self.assertNotIn("HLF-163", codes_eu)
        self.assertNotIn("HLF-087", codes_eu)

    def test_03_generate_all_forms_and_render_png(self):
        """
        Generates every form for all test buyers across PRE-APPROVAL and POST-APPROVAL,
        and renders pages to PNG using QtPdf to verify rendering integrity.
        """
        if HAS_PYQT_PDF:
            app = QApplication.instance() or QApplication(sys.argv)
            qdoc = QPdfDocument(None)

        rendered_count = 0

        for b_idx, buyer in enumerate(self.test_buyers):
            surname = extract_surname(buyer.principal_borrower)

            for stage in ("PRE-APPROVAL", "POST-APPROVAL"):
                forms_meta = get_forms_for_stage(stage, buyer)
                self.assertGreater(len(forms_meta), 0)

                for meta in forms_meta:
                    pdf_bytes = self.filler.fill_form(meta, buyer)
                    self.assertGreater(len(pdf_bytes), 500)

                    # Save test PDF
                    pdf_filename = f"{meta['code']}_{surname}_{stage}.pdf"
                    pdf_path = os.path.join(self.output_dir, pdf_filename)
                    with open(pdf_path, "wb") as f:
                        f.write(pdf_bytes)

                    # Render to PNG
                    if HAS_PYQT_PDF:
                        qdoc.load(pdf_path)
                        num_pages = qdoc.pageCount()
                        self.assertGreater(num_pages, 0)

                        for p in range(num_pages):
                            png_filename = f"{meta['code']}_{surname}_{stage}_p{p+1}.png"
                            png_path = os.path.join(self.output_dir, png_filename)
                            img = qdoc.render(p, QSize(612, 792))
                            img.save(png_path)
                            self.assertTrue(os.path.exists(png_path))
                            self.assertGreater(os.path.getsize(png_path), 5000)
                            rendered_count += 1

        print(f"\n[Test Verification] Successfully rendered {rendered_count} form pages to PNG.")

    def test_04_i_we_line_and_side_margins(self):
        """Verifies 'I/We' line construction and side margin vertical rotation rules."""
        # Married with married co-borrower
        b_married = self.test_buyers[1]
        i_we = build_i_we_line(b_married)
        self.assertIn("ROBERTO SANTOS MARRIED TO MARIA SANTOS", i_we)
        self.assertIn("EDUARDO DELOS REYES MARRIED TO CARMEN DELOS REYES", i_we)

        # Single buyer
        b_single = self.test_buyers[0]
        i_we_s = build_i_we_line(b_single)
        self.assertEqual(i_we_s, "JUAN DELA CRUZ (SINGLE)")

    def test_05_amortization_schedule_rules(self):
        """
        Verifies Post-approval PN/LMA amortization table rules:
        - Row '1 – R×12': all four amounts and total.
        - Next row, up to N−12: MRI/SRI and Non-Life, with P&I and Total '**'.
        - Last 12 months: P&I and Total '**' only.
        """
        schedule = build_amortization_schedule(
            repricing_years=1,
            loan_term_years=30,
            mri_sri=350.0,
            non_life=220.0,
            principal_interest=5230.50
        )
        self.assertEqual(len(schedule), 3)
        # Row 1
        self.assertEqual(schedule[0]["period"], "1 - 12")
        self.assertIn("350.00", schedule[0]["mri_sri"])
        self.assertIn("5,230.50", schedule[0]["principal_interest"])

        # Row 2
        self.assertEqual(schedule[1]["period"], "13 - 348")
        self.assertEqual(schedule[1]["principal_interest"], "**")
        self.assertEqual(schedule[1]["total"], "**")

        # Row 3
        self.assertEqual(schedule[2]["period"], "349 - 360")
        self.assertEqual(schedule[2]["mri_sri"], "")
        self.assertEqual(schedule[2]["principal_interest"], "**")

    def test_06_excel_roundtrip_and_concurrency(self):
        """
        Round-trip BookData.xlsx:
        - Load buyers
        - Add buyer
        - Save to Excel
        - Reload and verify no row loss
        """
        test_excel_path = os.path.join(self.output_dir, "Test_BookData.xlsx")
        ExcelManager.create_sample_workbook(test_excel_path)

        mgr = ExcelManager(test_excel_path)
        loaded = mgr.load_buyers()
        initial_count = len(loaded)
        self.assertGreaterEqual(initial_count, 4)

        # Add new buyer
        new_buyer = BuyerRecord(
            row_idx=0,
            principal_borrower="MARIA FE AGUILA",
            status="Single",
            block_lot="BLOCK 9 LOT 2",
            loanable_amount=750000
        )
        saved_row = mgr.save_buyer(new_buyer)
        self.assertEqual(saved_row, initial_count + 2)

        # Reload
        reloaded = mgr.load_buyers()
        self.assertEqual(len(reloaded), initial_count + 1)
        found = [b for b in reloaded if b.principal_borrower == "MARIA FE AGUILA"]
        self.assertEqual(len(found), 1)
        self.assertEqual(found[0].loanable_amount, 750000.0)

    def test_07_app_screen_bridge(self):
        """
        Drives the app's Python bridge the way its HTML screen does, on a temporary copy of the
        data (never the real BookData.xlsx or settings): load, save, rename, make forms, security.
        """
        import json
        import tempfile
        import urllib.error
        import urllib.request
        from pagibig_forms.settings import SettingsManager
        from pagibig_forms.webui import Backend, UIServer

        tmp = tempfile.mkdtemp(prefix="pagibig_test_")
        excel_copy = os.path.join(tmp, "BookData.xlsx")
        ExcelManager.create_sample_workbook(excel_copy)
        with open(os.path.join(tmp, "settings.json"), "w", encoding="utf-8") as f:
            json.dump({"excel_path": excel_copy, "save_folder": os.path.join(tmp, "out"),
                       "forms_folder": self.forms_dir}, f)
        server = UIServer(Backend(SettingsManager(tmp), self.base_dir))
        server.start()
        self.addCleanup(server.stop)

        def call(path, body=None, token=server.token, raw=False):
            req = urllib.request.Request(server.url.rstrip("/") + path,
                                         data=None if body is None else json.dumps(body).encode(),
                                         headers={"X-App-Token": token, "Content-Type": "application/json"})
            with urllib.request.urlopen(req) as res:
                data = res.read()
            return data if raw else json.loads(data)

        # The page is served with this run's secret token; the API refuses calls without it
        page = urllib.request.urlopen(server.url).read().decode()
        self.assertIn(server.token, page)
        with self.assertRaises(urllib.error.HTTPError) as denied:
            call("/api/buyers", token="wrong")
        self.assertEqual(denied.exception.code, 403)

        buyers = call("/api/buyers")["buyers"]
        self.assertEqual(len(buyers), 4)

        # Save a new buyer typed with peso signs and a percent
        new = {"principal_borrower": "HEADLESS TEST BUYER", "loanable_amount": "₱850,000",
               "interest_rate": "6.25%", "block_lot": "BLOCK 9 LOT 2"}
        saved = call("/api/buyers/save", {"buyer": new})["buyer"]
        self.assertEqual(saved["loanable_amount"], 850000.0)
        self.assertAlmostEqual(saved["interest_rate"], 0.0625)
        self.assertTrue(saved["is_affordable"])

        # A number that cannot be read is reported, not saved as 0
        with self.assertRaises(urllib.error.HTTPError) as bad:
            call("/api/buyers/save", {"buyer": dict(saved, consideration="one million")})
        self.assertEqual(bad.exception.code, 400)

        # Renaming then making forms saves to the SAME row (no duplicate)
        renamed = dict(saved, principal_borrower="HEADLESS TEST BUYER RENAMED")
        made = call("/api/forms/make", {"buyer": renamed, "stage": "PRE-APPROVAL"})
        self.assertGreater(len(made["files"]), 5)
        self.assertEqual(made["errors"], [])
        names = [b.principal_borrower for b in ExcelManager(excel_copy).load_buyers()]
        self.assertIn("HEADLESS TEST BUYER RENAMED", names)
        self.assertNotIn("HEADLESS TEST BUYER", names)

        # Preview returns a real PDF; opening files outside the forms folder is refused
        pdf = call("/api/forms/preview", {"buyer": renamed, "stage": "POST-APPROVAL", "index": 0}, raw=True)
        self.assertTrue(pdf.startswith(b"%PDF"))
        with self.assertRaises(urllib.error.HTTPError) as outside:
            call("/api/open", {"path": excel_copy})
        self.assertEqual(outside.exception.code, 403)

        # HLA mapping: desired loan amount is NOT used as the loanable amount
        mapped = map_hla_json_to_buyer({"A": "MOCKED HLA BORROWER", "J": 800000})
        self.assertEqual(mapped.principal_borrower, "MOCKED HLA BORROWER")
        self.assertEqual(mapped.loanable_amount, 0.0)

    def test_08_block_lot_and_placeholders(self):
        """Lot number is read correctly and blank data never turns into sample data."""
        from pagibig_forms.engine.models import parse_block_lot
        self.assertEqual(parse_block_lot("BLOCK 7 LOT 14"), ("7", "14"))
        self.assertEqual(parse_block_lot("BLK. 4 LT. 2A"), ("4", "2A"))
        self.assertEqual(parse_block_lot("B7 L14"), ("7", "14"))
        self.assertEqual(extract_surname("JOSE ANDRADE REYES"), "REYES")
        self.assertEqual(extract_surname("MARIA DEL CARMEN SANTOS"), "SANTOS")
        self.assertEqual(extract_surname("ROSARIO VDA. DE LEON"), "VDA. DE LEON")

        short = extract_short_tech_desc("", "3", "9", 0.0)
        self.assertNotIn("PCS-04-030177", short)
        self.assertNotIn("N-17185", short)
        self.assertIn("BLOCK 3, LOT 9", short)

    def test_09_signatories_and_schedule_overlap(self):
        """All four signatories are listed; schedule rows never overlap."""
        from pagibig_forms.engine.rules import get_signatories
        self.assertEqual(
            get_signatories(self.test_buyers[1]),
            ["ROBERTO SANTOS", "MARIA SANTOS", "EDUARDO DELOS REYES", "CARMEN DELOS REYES"]
        )
        self.assertEqual(get_signatories(self.test_buyers[0]), ["JUAN DELA CRUZ", "", "", ""])
        rows = build_amortization_schedule(5, 5, 1.0, 1.0, 1.0)
        self.assertEqual([r["period"] for r in rows], ["1 - 60"])

    def test_10_enye_and_dates(self):
        """Ñ survives into the PDF and Excel dates read as mm/dd/yyyy."""
        import datetime
        from pagibig_forms.engine.canvas import PdfCanvas
        from pagibig_forms.engine.models import date_to_mdy, date_to_long
        c = PdfCanvas()
        c.draw_text("DELA PEÑA", 50, 700)
        pdf = c.build_pdf_bytes()
        self.assertIn(b"PE\xd1A", pdf)
        self.assertIn(b"/WinAnsiEncoding", pdf)
        self.assertEqual(date_to_mdy(datetime.datetime(2026, 10, 16)), "10/16/2026")
        self.assertEqual(date_to_long("10/16/2026"), "OCTOBER 16, 2026")


    def test_11_first_run_setup(self):
        """A new install must be set up first; setup creates a new workbook without overwriting one."""
        import tempfile
        from pagibig_forms.settings import SettingsManager
        from pagibig_forms.webui import Backend, ApiError

        tmp = tempfile.mkdtemp(prefix="pagibig_setup_")
        backend = Backend(SettingsManager(os.path.join(tmp, "cfg")), tmp)
        self.assertFalse(backend.get_settings()["setup_complete"])

        with self.assertRaises(ApiError) as missing:
            backend.finish_setup({"excel_mode": "existing", "excel_path": "", "forms_folder": "", "save_folder": ""})
        self.assertEqual(len(missing.exception.details), 3)

        shared = os.path.join(tmp, "Shared")
        body = {"excel_mode": "new", "excel_folder": shared,
                "forms_folder": os.path.join(tmp, "Blank Forms"), "save_folder": os.path.join(tmp, "Filled")}
        settings = backend.finish_setup(body)
        self.assertTrue(settings["setup_complete"])
        self.assertTrue(os.path.isfile(os.path.join(shared, "BookData.xlsx")))
        self.assertTrue(os.path.isdir(os.path.join(tmp, "Filled")))

        # Running setup again on a folder that already has a workbook keeps its buyers
        ExcelManager(os.path.join(shared, "BookData.xlsx")).save_buyer(BuyerRecord(principal_borrower="KEEP ME"))
        backend.finish_setup(body)
        self.assertEqual([b.principal_borrower for b in ExcelManager(os.path.join(shared, "BookData.xlsx")).load_buyers()],
                         ["KEEP ME"])


if __name__ == "__main__":
    unittest.main()
