"""
Template Provider for Pag-IBIG Forms.
Loads existing templates from the `forms/` directory or synthesizes clean, official-standard
Pag-IBIG blank forms when blank files are not yet provided.
"""
from __future__ import annotations
import io
import os
from typing import Dict, List, Optional
import pypdf
from .canvas import PdfCanvas

# (center x, baseline y of the printed name, label) for the draft layout's signature block
SIGNATURE_SLOTS = [
    (150.0, 203.0, "PRINCIPAL BORROWER"),
    (450.0, 203.0, "SPOUSE OF PRINCIPAL BORROWER"),
    (150.0, 158.0, "CO-BORROWER"),
    (450.0, 158.0, "SPOUSE OF CO-BORROWER"),
]

# (label, baseline y) of the draft layout's financial rows; values start at x=175
FINANCIAL_ROWS = [
    ("Total Consideration", 530.0),
    ("Loanable Amount", 515.0),
    ("Equity Amount", 500.0),
    ("Interest Rate", 485.0),
]
FIN_Y = {label: y for label, y in FINANCIAL_ROWS}

# (forms folder, folder mtime, code, template name, title) -> template path or None
_TEMPLATE_CACHE: Dict[tuple, Optional[str]] = {}

class TemplateProvider:
    def __init__(self, forms_dir: str):
        self.forms_dir = forms_dir

    def find_template_path(self, form_code: str, template_name: str = "", form_title: str = "") -> Optional[str]:
        """
        Finds the official blank PDF in forms_dir: the exact template file name first,
        then a file whose name contains the form code or title. Returns None if missing.
        Results are cached until a file is added to or removed from the folder.
        """
        if not os.path.isdir(self.forms_dir):
            return None
        key = (os.path.realpath(self.forms_dir), os.path.getmtime(self.forms_dir), form_code, template_name, form_title)
        if key not in _TEMPLATE_CACHE:
            if len(_TEMPLATE_CACHE) > 500:
                _TEMPLATE_CACHE.clear()
            _TEMPLATE_CACHE[key] = self._search_template(form_code, template_name, form_title)
        return _TEMPLATE_CACHE[key]

    def _search_template(self, form_code: str, template_name: str, form_title: str) -> Optional[str]:
        pdfs = sorted(f for f in os.listdir(self.forms_dir) if f.lower().endswith(".pdf"))
        candidates = []
        if template_name:
            candidates += [f for f in pdfs if f.lower() == template_name.lower()]
        code_clean = form_code.replace("-", "").lower()
        candidates += [f for f in pdfs if code_clean in f.replace("-", "").replace("_", "").lower()]
        if form_title:
            candidates += [f for f in pdfs if form_title.lower() in f.lower()]
        for fname in candidates:
            file_path = os.path.join(self.forms_dir, fname)
            try:
                if len(pypdf.PdfReader(file_path).pages) > 0:
                    return file_path
            except Exception:
                continue
        return None

    def get_template_reader(
        self,
        form_code: str,
        fallback_pages: int = 1,
        form_title: str = "",
        template_name: str = ""
    ) -> pypdf.PdfReader:
        """
        Loads the official template PDF from forms_dir if present; otherwise builds a
        clearly marked DRAFT layout so it can never be mistaken for the official form.
        """
        path = self.find_template_path(form_code, template_name, form_title)
        if path:
            return pypdf.PdfReader(path)
        return self._synthesize_blank_form(form_code, fallback_pages, form_title)

    def _synthesize_blank_form(self, form_code: str, page_count: int, title: str) -> pypdf.PdfReader:
        """Creates a pristine multi-page Pag-IBIG blank document with official styling."""
        writer = pypdf.PdfWriter()

        for page_idx in range(page_count):
            canvas = PdfCanvas(width=612.0, height=792.0)  # Standard Letter

            # Header
            canvas.draw_text("PAG-IBIG FUND", 306.0, 755.0, font_name="Helvetica-Bold", font_size=11.0, align="center")
            canvas.draw_text("(HOME DEVELOPMENT MUTUAL FUND)", 306.0, 743.0, font_name="Helvetica", font_size=8.0, align="center")
            display_title = title.upper() if title else f"FORM {form_code}"
            canvas.draw_text(display_title, 306.0, 727.0, font_name="Helvetica-Bold", font_size=12.0, align="center")
            canvas.draw_text(f"FORM HQP-{form_code} (Page {page_idx + 1} of {page_count})", 562.0, 760.0, font_name="Helvetica", font_size=7.0, align="right")

            # Border
            canvas.draw_line(40.0, 715.0, 572.0, 715.0, width=1.0)
            canvas.draw_line(40.0, 45.0, 572.0, 45.0, width=0.5)

            # Footer notice
            canvas.draw_text("SUMMERFIELD RESOURCES CORP. - BLOOMFIELD TERRACES HOUSING LOAN", 40.0, 32.0, font_name="Helvetica", font_size=6.5)
            canvas.draw_text("DRAFT LAYOUT - OFFICIAL PAG-IBIG FORM NOT FOUND, DO NOT SUBMIT", 572.0, 32.0, font_name="Helvetica-Bold", font_size=6.5, align="right")

            # Page-specific blank lines and boxes
            if page_idx == 0:
                # Borrower & Property Details Section
                canvas.draw_text("I. PARTIES AND PROPERTY DESCRIPTION", 45.0, 695.0, font_name="Helvetica-Bold", font_size=9.0)
                canvas.draw_text("Principal Borrower:", 45.0, 675.0, font_name="Helvetica", font_size=8.0)
                canvas.draw_line(135.0, 673.0, 360.0, 673.0, width=0.5)

                canvas.draw_text("Marital Status:", 375.0, 675.0, font_name="Helvetica", font_size=8.0)
                canvas.draw_line(440.0, 673.0, 565.0, 673.0, width=0.5)

                canvas.draw_text("Spouse Name:", 45.0, 655.0, font_name="Helvetica", font_size=8.0)
                canvas.draw_line(135.0, 653.0, 360.0, 653.0, width=0.5)

                canvas.draw_text("TIN Principal:", 375.0, 655.0, font_name="Helvetica", font_size=8.0)
                canvas.draw_line(440.0, 653.0, 565.0, 653.0, width=0.5)

                canvas.draw_text("Co-Borrower:", 45.0, 635.0, font_name="Helvetica", font_size=8.0)
                canvas.draw_line(135.0, 633.0, 360.0, 633.0, width=0.5)

                canvas.draw_text("TIN Co-Borrower:", 375.0, 635.0, font_name="Helvetica", font_size=8.0)
                canvas.draw_line(440.0, 633.0, 565.0, 633.0, width=0.5)

                canvas.draw_text("Present Address:", 45.0, 615.0, font_name="Helvetica", font_size=8.0)
                canvas.draw_line(135.0, 613.0, 565.0, 613.0, width=0.5)

                canvas.draw_text("Property / Project:", 45.0, 595.0, font_name="Helvetica", font_size=8.0)
                canvas.draw_line(135.0, 593.0, 360.0, 593.0, width=0.5)

                canvas.draw_text("Block & Lot:", 375.0, 595.0, font_name="Helvetica", font_size=8.0)
                canvas.draw_line(440.0, 593.0, 565.0, 593.0, width=0.5)

                canvas.draw_text("Title Number (TCT):", 45.0, 575.0, font_name="Helvetica", font_size=8.0)
                canvas.draw_line(135.0, 573.0, 360.0, 573.0, width=0.5)

                # Loan terms & consideration
                canvas.draw_text("II. FINANCIAL TERMS AND CONSIDERATION", 45.0, 545.0, font_name="Helvetica-Bold", font_size=9.0)
                # Full-width rows: amounts are written in words and figures
                for label, y in FINANCIAL_ROWS:
                    canvas.draw_text(f"{label}:", 45.0, y, font_name="Helvetica", font_size=8.0)
                    canvas.draw_line(170.0, y - 2.0, 565.0, y - 2.0, width=0.5)

                # Technical description area
                canvas.draw_text("III. TECHNICAL DESCRIPTION OF PROPERTY", 45.0, 470.0, font_name="Helvetica-Bold", font_size=9.0)
                canvas.draw_rect(45.0, 260.0, 520.0, 198.0, fill=False, stroke=True)

            elif page_idx == 1:
                # Amortization Table / Terms
                canvas.draw_text("IV. SCHEDULE OF AMORTIZATION AND MONTHLY INSTALLMENTS", 45.0, 695.0, font_name="Helvetica-Bold", font_size=9.0)
                # Table headers
                canvas.draw_rect(45.0, 560.0, 520.0, 120.0, fill=False, stroke=True)
                canvas.draw_line(45.0, 655.0, 565.0, 655.0, width=0.5)
                canvas.draw_line(135.0, 560.0, 135.0, 680.0, width=0.5)
                canvas.draw_line(235.0, 560.0, 235.0, 680.0, width=0.5)
                canvas.draw_line(335.0, 560.0, 335.0, 680.0, width=0.5)
                canvas.draw_line(445.0, 560.0, 445.0, 680.0, width=0.5)

                canvas.draw_text("Period (Months)", 90.0, 663.0, font_name="Helvetica-Bold", font_size=7.5, align="center")
                canvas.draw_text("MRI / SRI", 185.0, 663.0, font_name="Helvetica-Bold", font_size=7.5, align="center")
                canvas.draw_text("Non-Life", 285.0, 663.0, font_name="Helvetica-Bold", font_size=7.5, align="center")
                canvas.draw_text("Principal & Interest", 390.0, 663.0, font_name="Helvetica-Bold", font_size=7.5, align="center")
                canvas.draw_text("Total Monthly", 505.0, 663.0, font_name="Helvetica-Bold", font_size=7.5, align="center")

            # Signature block on last page
            if page_idx == page_count - 1:
                canvas.draw_text("CONFORME / SIGNATURE OF PARTIES:", 45.0, 240.0, font_name="Helvetica-Bold", font_size=8.5)
                # Four signatory slots: principal & spouse, co-borrower & spouse
                for (sx, sy, label) in SIGNATURE_SLOTS:
                    canvas.draw_line(sx - 100.0, sy - 3.0, sx + 100.0, sy - 3.0, width=0.5)
                    canvas.draw_text(label, sx, sy - 13.0, font_name="Helvetica", font_size=7.0, align="center")

                canvas.draw_text("SIGNED IN THE PRESENCE OF:", 45.0, 120.0, font_name="Helvetica-Bold", font_size=8.0)
                canvas.draw_line(50.0, 80.0, 250.0, 80.0, width=0.5)
                canvas.draw_text("WITNESS", 150.0, 68.0, font_name="Helvetica", font_size=7.5, align="center")

                canvas.draw_line(350.0, 80.0, 550.0, 80.0, width=0.5)
                canvas.draw_text("WITNESS", 450.0, 68.0, font_name="Helvetica", font_size=7.5, align="center")

            writer.add_page(canvas.to_page_object())

        buf = io.BytesIO()
        writer.write(buf)
        buf.seek(0)
        return pypdf.PdfReader(buf)
