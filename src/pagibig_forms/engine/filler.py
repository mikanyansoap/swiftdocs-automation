"""
Pag-IBIG Forms Fill Engine.
Applies exact fill rules, calculations, technical description layouts, side margins,
and amortization schedules onto form templates via pure-Python PDF overlay.
Blank Excel cells leave the form blank: no sample or placeholder values are ever printed.
"""
from __future__ import annotations
import io
from typing import Any, Dict, List, Optional
import pypdf

from .models import BuyerRecord, date_to_long
from .number_to_words import (
    amount_words_and_figures,
    interest_rate_to_words_and_figures,
    term_to_words_and_figures
)
from .tech_desc import format_columnar_tech_desc, extract_short_tech_desc
from .rules import (
    build_i_we_line,
    build_side_margins,
    build_amortization_schedule,
    get_signatories,
    get_witnesses,
)
from .canvas import PdfCanvas, measure_text_width
from .template_provider import TemplateProvider, SIGNATURE_SLOTS, FIN_Y

# Default page count for the draft layout used when the official PDF is missing
PAGE_COUNTS = {
    "HLF-161": 4, "HLF-524": 2, "HLF-785": 1, "HLF-063": 1,
    "CIBI": 1, "HLF-058": 1, "HLF-913": 1, "HLF-525": 2,
    "HLF-163": 4, "HLF-162": 4, "HLF-087": 2, "HLF-086": 2,
    "HLA-ACCEPT": 1, "HLF-083": 1, "HLF-371": 1
}

# Signature pages (0-indexed); side-margin names are left off these pages
SIGNED_PAGES = {
    "HLF-161": [3],   # Page 4
    "HLF-524": [1],   # Page 2
    "HLF-525": [1],   # Page 2
    "HLF-163": [3],   # Page 4
    "HLF-162": [3],   # Page 4
}

TECH_DESC_TOP = 445.0
TECH_DESC_BOTTOM = 268.0  # bottom of the technical description box


class FormFiller:
    def __init__(self, forms_dir: str, settings: Optional[Dict[str, Any]] = None):
        self.forms_dir = forms_dir
        self.settings = settings or {}
        self.template_provider = TemplateProvider(forms_dir)

    def has_official_template(self, form_meta: Dict[str, Any]) -> bool:
        """True if the official blank PDF for this form is in the forms folder."""
        return self.template_provider.find_template_path(
            form_meta.get("code", ""), form_meta.get("template_name", ""), form_meta.get("name", "")
        ) is not None

    def fill_form(
        self,
        form_meta: Dict[str, Any],
        buyer: BuyerRecord
    ) -> bytes:
        """
        Fills a single form and returns the final PDF bytes.
        """
        code = form_meta.get("code", "")
        name = form_meta.get("name", "")
        borrower_type = form_meta.get("borrower_type", "principal")

        template_reader = self.template_provider.get_template_reader(
            form_code=code,
            fallback_pages=PAGE_COUNTS.get(code, 1),
            form_title=name,
            template_name=form_meta.get("template_name", "")
        )

        total_pages = len(template_reader.pages)
        writer = pypdf.PdfWriter()
        signed_pages = SIGNED_PAGES.get(code, [total_pages - 1])

        # Prepare common values
        witness_1, witness_2 = get_witnesses(
            is_affordable=buyer.is_affordable,
            is_lma_or_pn=(code in ("HLF-163", "HLF-162", "HLF-087", "HLF-086")),
            settings=self.settings
        )

        # Developer settings
        dev_name = self.settings.get("developer_name", "")
        dev_rep = self.settings.get("developer_rep", "")
        bvs_mun = self.settings.get("bvs_municipality", "")
        proj_addr = self.settings.get("project_address", "")

        # Iterate pages and overlay
        for page_idx in range(total_pages):
            base_page = template_reader.pages[page_idx]
            page_box = base_page.mediabox
            width = float(page_box.width)
            height = float(page_box.height)

            canvas = PdfCanvas(width=width, height=height)

            # 1. Side margins: printed vertically, spaced for signatures, no lines.
            # Left off any page where that person signs.
            if code in ("HLF-161", "HLF-524", "HLF-525", "HLF-163", "HLF-162"):
                side_names = build_side_margins(buyer, page_idx, signed_pages)
                if side_names:
                    # Up the left edge, each name followed by room to sign; long names never overlap
                    y_start, y_end, gap = 120.0, height - 100.0, 60.0
                    size = 7.5
                    widths = [measure_text_width(n, "Helvetica-Bold", size) for n in side_names]
                    needed = sum(widths) + gap * (len(side_names) - 1)
                    if needed > y_end - y_start:
                        scale = (y_end - y_start) / needed
                        size *= scale
                        gap *= scale
                        widths = [w * scale for w in widths]
                    cur_y = y_start
                    for s_name, w in zip(side_names, widths):
                        # 22 points from left edge
                        canvas.draw_vertical_text(s_name, 22.0, cur_y, font_name="Helvetica-Bold", font_size=size)
                        cur_y += w + gap

            # 2. Form-specific fields
            if code == "HLF-161":  # Contract-to-Sell
                self._fill_cts(canvas, buyer, page_idx, proj_addr)
            elif code == "HLF-524":  # Deed of Assignment
                self._fill_doa(canvas, buyer, page_idx)
            elif code == "HLF-785":  # HLA Acknowledgement Receipt
                self._fill_hla_receipt(canvas, buyer)
            elif code == "HLF-063":  # MSVS
                self._fill_msvs(canvas, buyer, borrower_type)
            elif code == "CIBI":     # CIBI / Data Privacy
                self._fill_cibi(canvas, buyer, borrower_type)
            elif code == "HLF-058":  # BVS
                self._fill_bvs(canvas, buyer, borrower_type, dev_name, bvs_mun, proj_addr)
            elif code == "HLF-913":  # ARIS
                self._fill_aris(canvas, buyer)
            elif code == "HLF-525":  # Deed of Absolute Sale
                self._fill_das(canvas, buyer, page_idx, dev_name, dev_rep, witness_1, witness_2)
            elif code in ("HLF-163", "HLF-162"):  # LMA
                self._fill_lma(canvas, buyer, page_idx, witness_1, witness_2)
            elif code in ("HLF-087", "HLF-086"):  # Promissory Note
                self._fill_pn(canvas, buyer, page_idx, witness_1, witness_2)
            elif code in ("HLA-ACCEPT", "HLF-083", "HLF-371"):
                self._fill_acceptance_forms(canvas, buyer, bvs_mun, proj_addr)

            # Add to writer first, then merge overlay to adhere to newest pypdf standard
            writer.add_page(base_page)
            writer.pages[-1].merge_page(canvas.to_page_object())

        out_buf = io.BytesIO()
        writer.write(out_buf)
        return out_buf.getvalue()

    # --- Shared helpers ---

    @staticmethod
    def _draw_signatories(canvas: PdfCanvas, buyer: BuyerRecord):
        """Principal, spouse, co-borrower and co-borrower's spouse each in their own signature slot."""
        for (x, y, _label), name in zip(SIGNATURE_SLOTS, get_signatories(buyer)):
            canvas.draw_text(name, x, y, font_name="Helvetica-Bold", font_size=8.5, max_width=195.0, align="center")

    @staticmethod
    def _draw_single_signatory(canvas: PdfCanvas, name: str):
        x, y, _label = SIGNATURE_SLOTS[0]
        canvas.draw_text(name, x, y, font_name="Helvetica-Bold", font_size=8.5, max_width=195.0, align="center")

    @staticmethod
    def _draw_witnesses(canvas: PdfCanvas, w1: str, w2: str):
        canvas.draw_text(w1, 150.0, 83.0, font_name="Helvetica", font_size=8.0, align="center")
        canvas.draw_text(w2, 450.0, 83.0, font_name="Helvetica", font_size=8.0, align="center")

    @staticmethod
    def _draw_financial(canvas: PdfCanvas, label: str, text: str, font_name: str = "Helvetica"):
        """Writes an amount (words and figures) on its full-width row of section II."""
        canvas.draw_text(text, 175.0, FIN_Y[label], font_name=font_name, font_size=8.0, max_width=390.0)

    @staticmethod
    def _draw_parties(canvas: PdfCanvas, buyer: BuyerRecord):
        """Principal, status, spouse and co-borrower rows at the top of page 1."""
        canvas.draw_text(buyer.principal_borrower.upper(), 140.0, 675.0, font_name="Helvetica-Bold", font_size=8.5, max_width=210.0)
        canvas.draw_text(buyer.marital_status.upper(), 445.0, 675.0, font_name="Helvetica", font_size=8.0, max_width=115.0)
        if buyer.marital_status.upper() == "MARRIED" and buyer.spouse_name:
            canvas.draw_text(buyer.spouse_name.upper(), 140.0, 655.0, font_name="Helvetica-Bold", font_size=8.5, max_width=210.0)
        if buyer.has_co_borrower:
            co_line = buyer.co_borrower.upper()
            if buyer.co_marital_status.upper() == "MARRIED" and buyer.co_spouse_name:
                co_line += f" MARRIED TO {buyer.co_spouse_name.upper()}"
            canvas.draw_text(co_line, 140.0, 635.0, font_name="Helvetica-Bold", font_size=8.5, max_width=210.0)

    @staticmethod
    def _draw_represented_by(canvas: PdfCanvas, buyer: BuyerRecord):
        if buyer.has_attorney_in_fact:
            aif = buyer.attorney_in_fact.upper()
            if buyer.co_attorney_in_fact.strip():
                aif += f" AND {buyer.co_attorney_in_fact.strip().upper()}"
            canvas.draw_text(f"REPRESENTED BY: {aif}", 375.0, 575.0, font_name="Helvetica-Bold", font_size=8.0, max_width=190.0)

    # --- Specific Form Fill Routines ---

    def _fill_cts(self, canvas: PdfCanvas, buyer: BuyerRecord, page_idx: int, proj_addr: str):
        if page_idx == 0:
            # Parties & Property
            self._draw_parties(canvas, buyer)
            canvas.draw_text(buyer.tin_principal, 445.0, 655.0, font_name="Helvetica", font_size=8.0, max_width=115.0)
            if buyer.has_co_borrower:
                canvas.draw_text(buyer.tin_co_borrower, 445.0, 635.0, font_name="Helvetica", font_size=8.0, max_width=115.0)

            canvas.draw_text(buyer.address.upper(), 140.0, 615.0, font_name="Helvetica", font_size=8.0, max_width=415.0)
            canvas.draw_text(proj_addr.upper(), 140.0, 595.0, font_name="Helvetica", font_size=8.0, max_width=210.0)
            canvas.draw_text(buyer.lot_label, 445.0, 595.0, font_name="Helvetica-Bold", font_size=8.5, max_width=115.0)
            canvas.draw_text(buyer.title_no, 140.0, 575.0, font_name="Helvetica", font_size=8.0, max_width=210.0)
            self._draw_represented_by(canvas, buyer)

            # Consideration, Loanable Amount, Equity
            if buyer.consideration > 0:
                self._draw_financial(canvas, "Total Consideration", amount_words_and_figures(buyer.consideration), font_name="Helvetica")
            if buyer.loanable_amount > 0:
                self._draw_financial(canvas, "Loanable Amount", amount_words_and_figures(buyer.loanable_amount), font_name="Helvetica")
            if buyer.consideration > 0:
                self._draw_financial(canvas, "Equity Amount", amount_words_and_figures(buyer.equity), font_name="Helvetica")
            if buyer.interest_rate > 0:
                rate_str = interest_rate_to_words_and_figures(buyer.interest_rate)
                self._draw_financial(canvas, "Interest Rate", rate_str, font_name="Helvetica")

            # Full Technical Description: starts 2 inches (144 pt) from left page edge, Courier, max 10 pt
            full_desc = format_columnar_tech_desc(buyer.full_tech_desc)
            canvas.draw_multiline_text(full_desc, 144.0, TECH_DESC_TOP, font_name="Courier", font_size=8.5,
                                       line_height=10.5, max_width=420.0, bottom_y=TECH_DESC_BOTTOM)

        elif page_idx == 1:
            # CTS page 2 rules:
            # - "projected monthly installments" = MRI/SRI + Non-Life + P&I, in words and figures.
            # - Section 1.2 Interest: fill ONLY End-User or Affordable paragraph, chosen by loan amount.
            #   In Affordable paragraph, also fill "every ___ ( ) years" from Re-pricing.
            # - If those Excel cells are blank, leave form blanks empty.
            if buyer.mri_sri > 0 or buyer.non_life > 0 or buyer.principal_interest > 0:
                tot_installment = buyer.mri_sri + buyer.non_life + buyer.principal_interest
                inst_w_f = amount_words_and_figures(tot_installment)
                canvas.draw_text(f"Projected Monthly Installment: {inst_w_f}", 45.0, 540.0, font_name="Helvetica-Bold", font_size=8.5, max_width=515.0)

            if buyer.is_affordable and buyer.repricing_years > 0:
                rep_words = term_to_words_and_figures(buyer.repricing_years, "years")
                canvas.draw_text(f"[AFFORDABLE HOUSING PROGRAM] Interest re-pricing: every {rep_words}", 45.0, 520.0, font_name="Helvetica-Bold", font_size=8.0)

        elif page_idx == 3:
            # Signatures & Notary
            self._draw_signatories(canvas, buyer)
            if buyer.notary_cts:
                canvas.draw_text(f"NOTARY: {buyer.notary_cts}", 45.0, 52.0, font_name="Helvetica", font_size=8.0, max_width=520.0)

    def _fill_doa(self, canvas: PdfCanvas, buyer: BuyerRecord, page_idx: int):
        if page_idx == 0:
            self._draw_parties(canvas, buyer)
            canvas.draw_text(buyer.address.upper(), 140.0, 615.0, font_name="Helvetica", font_size=8.0, max_width=415.0)
            canvas.draw_text(buyer.lot_label, 445.0, 595.0, font_name="Helvetica-Bold", font_size=8.5, max_width=115.0)
            canvas.draw_text(buyer.title_no, 140.0, 575.0, font_name="Helvetica", font_size=8.0, max_width=210.0)

            # "Contract-to-Sell dated ___" from Date of Notarization
            cts_date = date_to_long(buyer.date_notarization)
            if cts_date:
                canvas.draw_text(f"Contract-to-Sell dated: {cts_date}", 45.0, 560.0, font_name="Helvetica-Bold", font_size=8.5)
            self._draw_represented_by(canvas, buyer)

            # Full technical description starting 2 inches from left edge
            full_desc = format_columnar_tech_desc(buyer.full_tech_desc)
            canvas.draw_multiline_text(full_desc, 144.0, TECH_DESC_TOP, font_name="Courier", font_size=8.5,
                                       line_height=10.5, max_width=420.0, bottom_y=TECH_DESC_BOTTOM)

        elif page_idx == 1:
            self._draw_signatories(canvas, buyer)

    def _fill_hla_receipt(self, canvas: PdfCanvas, buyer: BuyerRecord):
        canvas.draw_text(buyer.principal_borrower.upper(), 150.0, 675.0, font_name="Helvetica-Bold", font_size=9.0, max_width=210.0)
        if buyer.has_co_borrower:
            canvas.draw_text(buyer.co_borrower.upper(), 150.0, 635.0, font_name="Helvetica", font_size=8.5, max_width=210.0)
        if buyer.has_attorney_in_fact:
            rel = f" ({buyer.attorney_in_fact_rel.upper()})" if buyer.attorney_in_fact_rel else ""
            canvas.draw_text(f"{buyer.attorney_in_fact.upper()}{rel}", 150.0, 615.0, font_name="Helvetica-Bold", font_size=8.5, max_width=410.0)
        canvas.draw_text(date_to_long(buyer.date_notarization), 445.0, 675.0, font_name="Helvetica", font_size=8.5, max_width=115.0)

    @staticmethod
    def _person(buyer: BuyerRecord, borrower_type: str):
        if borrower_type == "co_borrower":
            return buyer.co_borrower.upper(), buyer.co_marital_status.upper()
        return buyer.principal_borrower.upper(), buyer.marital_status.upper()

    def _fill_msvs(self, canvas: PdfCanvas, buyer: BuyerRecord, borrower_type: str):
        # MSVS/CIBI rule: names and marital status only
        name, status = self._person(buyer, borrower_type)
        canvas.draw_text(name, 140.0, 675.0, font_name="Helvetica-Bold", font_size=9.0, max_width=210.0)
        canvas.draw_text(status, 445.0, 675.0, font_name="Helvetica", font_size=8.5, max_width=115.0)

    def _fill_cibi(self, canvas: PdfCanvas, buyer: BuyerRecord, borrower_type: str):
        # MSVS/CIBI rule: names and marital status only
        name, status = self._person(buyer, borrower_type)
        canvas.draw_text(name, 140.0, 675.0, font_name="Helvetica-Bold", font_size=9.0, max_width=210.0)
        canvas.draw_text(status, 445.0, 675.0, font_name="Helvetica", font_size=8.5, max_width=115.0)
        self._draw_single_signatory(canvas, name)

    def _fill_bvs(self, canvas: PdfCanvas, buyer: BuyerRecord, borrower_type: str, dev_name: str, bvs_mun: str, proj_addr: str):
        # BVS: Developer name, Block & Lot + project address, printed names. Municipality/City from Settings.
        # Leave blank for hand writing: Yes/No answers, validation mode, dates.
        name, _status = self._person(buyer, borrower_type)
        canvas.draw_text(name, 140.0, 675.0, font_name="Helvetica-Bold", font_size=8.5, max_width=210.0)
        canvas.draw_text(dev_name.upper(), 140.0, 595.0, font_name="Helvetica-Bold", font_size=8.0, max_width=210.0)
        canvas.draw_text(buyer.lot_label, 445.0, 595.0, font_name="Helvetica-Bold", font_size=8.5, max_width=115.0)
        canvas.draw_text(f"{buyer.lot_label}, {proj_addr.upper()}".strip(", "), 140.0, 615.0, font_name="Helvetica", font_size=8.0, max_width=415.0)
        canvas.draw_text(bvs_mun.upper(), 140.0, 575.0, font_name="Helvetica", font_size=8.5, max_width=210.0)
        self._draw_single_signatory(canvas, name)

    def _fill_aris(self, canvas: PdfCanvas, buyer: BuyerRecord):
        # ARIS (HLF-913): ONLY when attorney-in-fact exists.
        # Leave blank for hand writing: birth date, TIN, addresses, contacts, Yes/No answers, references.
        canvas.draw_text(buyer.attorney_in_fact.upper(), 140.0, 675.0, font_name="Helvetica-Bold", font_size=9.0, max_width=210.0)
        canvas.draw_text(buyer.principal_borrower.upper(), 140.0, 655.0, font_name="Helvetica", font_size=8.5, max_width=210.0)
        if buyer.has_co_borrower:
            canvas.draw_text(buyer.co_borrower.upper(), 140.0, 635.0, font_name="Helvetica", font_size=8.5, max_width=210.0)

    def _fill_das(self, canvas: PdfCanvas, buyer: BuyerRecord, page_idx: int, dev_name: str, dev_rep: str, w1: str, w2: str):
        if page_idx == 0:
            self._draw_parties(canvas, buyer)
            canvas.draw_text(buyer.lot_label, 445.0, 595.0, font_name="Helvetica-Bold", font_size=8.5, max_width=115.0)
            canvas.draw_text(buyer.title_no, 140.0, 575.0, font_name="Helvetica", font_size=8.0, max_width=210.0)
            self._draw_represented_by(canvas, buyer)

            if buyer.consideration > 0:
                cons_w_f = amount_words_and_figures(buyer.consideration)
                self._draw_financial(canvas, "Total Consideration", cons_w_f, font_name="Helvetica")

            # Short technical description rule:
            # Survey Plan No., Block + Lot on one row, Portion of Lot, LRC Record No., Area
            # Shrinks as needed, NEVER adding an annex page.
            short_desc = extract_short_tech_desc(buyer.full_tech_desc, buyer.block_no, buyer.lot_no, buyer.lot_area)
            canvas.draw_multiline_text(short_desc, 144.0, TECH_DESC_TOP, font_name="Courier", font_size=8.5,
                                       line_height=11.0, max_width=410.0, bottom_y=TECH_DESC_BOTTOM)

        elif page_idx == 1:
            # Vendor (developer) above the vendees' signature block, then witnesses
            if dev_name or dev_rep:
                vendor = f"VENDOR: {dev_name.upper()}" + (f" by {dev_rep.upper()}" if dev_rep else "")
                canvas.draw_text(vendor, 565.0, 240.0, font_name="Helvetica-Bold", font_size=8.0, max_width=330.0, align="right")
            self._draw_signatories(canvas, buyer)
            self._draw_witnesses(canvas, w1, w2)

    def _fill_lma(self, canvas: PdfCanvas, buyer: BuyerRecord, page_idx: int, w1: str, w2: str):
        if page_idx == 0:
            self._draw_parties(canvas, buyer)
            canvas.draw_text(buyer.lot_label, 445.0, 595.0, font_name="Helvetica-Bold", font_size=8.5, max_width=115.0)
            canvas.draw_text(buyer.title_no, 140.0, 575.0, font_name="Helvetica", font_size=8.0, max_width=210.0)

            # Loan amount, interest rate and term
            if buyer.loanable_amount > 0:
                self._draw_financial(canvas, "Loanable Amount", amount_words_and_figures(buyer.loanable_amount), font_name="Helvetica")
            if buyer.interest_rate > 0:
                self._draw_financial(canvas, "Interest Rate", interest_rate_to_words_and_figures(buyer.interest_rate), font_name="Helvetica")
            if buyer.loan_term_years > 0:
                term = term_to_words_and_figures(buyer.loan_term_years, "years")
                canvas.draw_text(f"LOAN TERM: {term}", 45.0, 560.0, font_name="Helvetica-Bold", font_size=8.0)

            # Short technical description
            short_desc = extract_short_tech_desc(buyer.full_tech_desc, buyer.block_no, buyer.lot_no, buyer.lot_area)
            canvas.draw_multiline_text(short_desc, 144.0, TECH_DESC_TOP, font_name="Courier", font_size=8.5,
                                       line_height=11.0, max_width=410.0, bottom_y=TECH_DESC_BOTTOM)

        elif page_idx == 1:
            self._draw_schedule(canvas, buyer, top_y=635.0, row_gap=22.0)

        elif page_idx == 3:
            # Signatures & Witnesses (leave HDMF signatory name/position blank for handwriting)
            self._draw_signatories(canvas, buyer)
            self._draw_witnesses(canvas, w1, w2)

    @staticmethod
    def _draw_schedule(canvas: PdfCanvas, buyer: BuyerRecord, top_y: float, row_gap: float):
        """Amortization table; left blank when no monthly amounts are entered."""
        if not (buyer.mri_sri > 0 or buyer.non_life > 0 or buyer.principal_interest > 0):
            return
        schedule = build_amortization_schedule(
            repricing_years=buyer.repricing_years,
            loan_term_years=buyer.loan_term_years,
            mri_sri=buyer.mri_sri,
            non_life=buyer.non_life,
            principal_interest=buyer.principal_interest
        )
        y = top_y
        for row in schedule:
            canvas.draw_text(row["period"], 90.0, y, font_name="Helvetica", font_size=8.0, align="center")
            canvas.draw_text(row["mri_sri"], 185.0, y, font_name="Helvetica", font_size=8.0, align="center")
            canvas.draw_text(row["non_life"], 285.0, y, font_name="Helvetica", font_size=8.0, align="center")
            canvas.draw_text(row["principal_interest"], 390.0, y, font_name="Helvetica", font_size=8.0, align="center")
            canvas.draw_text(row["total"], 505.0, y, font_name="Helvetica-Bold", font_size=8.0, align="center")
            y -= row_gap

    def _fill_pn(self, canvas: PdfCanvas, buyer: BuyerRecord, page_idx: int, w1: str, w2: str):
        if page_idx == 0:
            self._draw_parties(canvas, buyer)

            # "dated ___" from Date of Notarization
            pn_date = date_to_long(buyer.date_notarization)
            if pn_date:
                canvas.draw_text(f"DATED: {pn_date}", 375.0, 575.0, font_name="Helvetica-Bold", font_size=8.5)

            if buyer.loanable_amount > 0:
                self._draw_financial(canvas, "Loanable Amount", amount_words_and_figures(buyer.loanable_amount), font_name="Helvetica-Bold")
            if buyer.interest_rate > 0:
                self._draw_financial(canvas, "Interest Rate", interest_rate_to_words_and_figures(buyer.interest_rate), font_name="Helvetica")

        elif page_idx == 1:
            # Amortization table on the PN's draft page 2 grid, then signatures
            self._draw_schedule(canvas, buyer, top_y=635.0, row_gap=22.0)
            self._draw_signatories(canvas, buyer)
            self._draw_witnesses(canvas, w1, w2)

    def _fill_acceptance_forms(self, canvas: PdfCanvas, buyer: BuyerRecord, bvs_mun: str, proj_addr: str):
        # Certificate of Acceptance and Buyer's Conformity:
        # the "I/We" line reads "BORROWER MARRIED TO SPOUSE & CO-BORROWER MARRIED TO SPOUSE".
        i_we = build_i_we_line(buyer)
        canvas.draw_text(f"I/We, {i_we}", 45.0, 655.0, font_name="Helvetica-Bold", font_size=8.0, max_width=515.0)
        canvas.draw_text(f"{buyer.lot_label}, {proj_addr.upper()}".strip(", "), 140.0, 615.0, font_name="Helvetica", font_size=8.5, max_width=415.0)
        canvas.draw_text(bvs_mun.upper(), 140.0, 595.0, font_name="Helvetica", font_size=8.5, max_width=210.0)
        self._draw_signatories(canvas, buyer)
