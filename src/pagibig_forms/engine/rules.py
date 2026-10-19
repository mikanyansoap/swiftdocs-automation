"""
Business logic and fill rules for Pag-IBIG Forms.
Produces exact text blocks, table values, witness assignments, and side-margin layouts.
"""
from __future__ import annotations
import math
from typing import Any, Dict, List, Optional, Tuple
from .models import BuyerRecord
from .number_to_words import (
    amount_to_words,
    format_currency_figures,
    interest_rate_to_words_and_figures,
    area_to_words_and_figures,
    years_to_months,
    term_to_words_and_figures
)
from .tech_desc import format_columnar_tech_desc, extract_short_tech_desc


def build_i_we_line(buyer: BuyerRecord) -> str:
    """
    Builds the 'I/We' line for Certificate of Acceptance and Buyer's Conformity:
    'BORROWER MARRIED TO SPOUSE & CO-BORROWER MARRIED TO SPOUSE'
    """
    parts = []
    # Principal
    p_status = buyer.marital_status.upper()
    if p_status == "MARRIED" and buyer.spouse_name:
        parts.append(f"{buyer.principal_borrower.upper()} MARRIED TO {buyer.spouse_name.upper()}")
    else:
        parts.append(f"{buyer.principal_borrower.upper()} ({p_status})")

    # Co-Borrower
    if buyer.has_co_borrower:
        c_status = buyer.co_marital_status.upper()
        if c_status == "MARRIED" and buyer.co_spouse_name:
            parts.append(f"{buyer.co_borrower.upper()} MARRIED TO {buyer.co_spouse_name.upper()}")
        elif c_status:
            parts.append(f"{buyer.co_borrower.upper()} ({c_status})")
        else:
            parts.append(buyer.co_borrower.upper())

    return " & ".join(parts)


def get_signatories(buyer: BuyerRecord) -> List[str]:
    """
    Signatory names in signature-slot order:
    [principal, principal's spouse, co-borrower, co-borrower's spouse].
    A blank string means nobody signs in that slot.
    """
    spouse = buyer.spouse_name.upper() if buyer.marital_status.upper() == "MARRIED" else ""
    co = buyer.co_borrower.upper() if buyer.has_co_borrower else ""
    co_spouse = ""
    if co and buyer.co_marital_status.upper() == "MARRIED":
        co_spouse = buyer.co_spouse_name.upper()
    return [buyer.principal_borrower.upper(), spouse, co, co_spouse]


def build_side_margins(buyer: BuyerRecord, page_num: int, signed_pages: List[int]) -> List[str]:
    """
    Side-margin names are printed vertically, spaced for signatures, with no lines.
    They are left off any page where that person signs.
    """
    if page_num in signed_pages:
        return []

    names = []
    # Principal Borrower
    names.append(buyer.principal_borrower.upper())
    # Spouse
    if buyer.marital_status.upper() == "MARRIED" and buyer.spouse_name:
        names.append(buyer.spouse_name.upper())
    # Co-Borrower
    if buyer.has_co_borrower:
        names.append(buyer.co_borrower.upper())
        if buyer.co_marital_status.upper() == "MARRIED" and buyer.co_spouse_name:
            names.append(buyer.co_spouse_name.upper())

    return names


def build_amortization_schedule(
    repricing_years: int,
    loan_term_years: int,
    mri_sri: float,
    non_life: float,
    principal_interest: float
) -> List[Dict[str, Any]]:
    """
    Post-approval PN/LMA amortization table, from re-pricing R years and term N months:
      - Row '1 – R×12': all four amounts and the total.
      - Next row, up to N−12: MRI/SRI and Non-Life, with P&I and Total '**'.
      - Last 12 months: P&I and Total '**' only.
    """
    R = repricing_years if repricing_years > 0 else 1
    N = years_to_months(loan_term_years if loan_term_years > 0 else 30)

    rows = []
    r_months = min(R * 12, N)
    total_1 = mri_sri + non_life + principal_interest

    # Row 1: 1 - R*12 (never past the end of the term)
    row1 = {
        "period": f"1 - {r_months}",
        "mri_sri": format_currency_figures(mri_sri, symbol=""),
        "non_life": format_currency_figures(non_life, symbol=""),
        "principal_interest": format_currency_figures(principal_interest, symbol=""),
        "total": format_currency_figures(total_1, symbol="")
    }
    rows.append(row1)

    # Next row: up to N - 12 (if N > r_months + 12)
    next_start = r_months + 1
    next_end = N - 12
    if next_end >= next_start:
        row2 = {
            "period": f"{next_start} - {next_end}",
            "mri_sri": format_currency_figures(mri_sri, symbol=""),
            "non_life": format_currency_figures(non_life, symbol=""),
            "principal_interest": "**",
            "total": "**"
        }
        rows.append(row2)

    # Last 12 months: (N - 11) to N, starting after any month already listed
    last_start = max(r_months + 1, N - 11)
    if last_start <= N:
        row3 = {
            "period": f"{last_start} - {N}",
            "mri_sri": "",
            "non_life": "",
            "principal_interest": "**",
            "total": "**"
        }
        rows.append(row3)

    return rows


def get_witnesses(is_affordable: bool, is_lma_or_pn: bool, settings: Dict[str, Any]) -> Tuple[str, str]:
    """
    Witnesses:
      - Affordable set and other post-approval forms: DENIELLE S. TOLENTINO and SALVACION M. AZCARRAGA.
      - End-User LMA/PN: VALENTINO P. HALFON and SALVACION M. AZCARRAGA.
    """
    aff_w = settings.get("witnesses_affordable", ["DENIELLE S. TOLENTINO", "SALVACION M. AZCARRAGA"])
    eu_w = settings.get("witnesses_end_user", ["VALENTINO P. HALFON", "SALVACION M. AZCARRAGA"])

    if not is_affordable and is_lma_or_pn:
        return (eu_w[0] if len(eu_w) > 0 else "VALENTINO P. HALFON",
                eu_w[1] if len(eu_w) > 1 else "SALVACION M. AZCARRAGA")
    return (aff_w[0] if len(aff_w) > 0 else "DENIELLE S. TOLENTINO",
            aff_w[1] if len(aff_w) > 1 else "SALVACION M. AZCARRAGA")


def get_forms_for_stage(stage: str, buyer: BuyerRecord) -> List[Dict[str, Any]]:
    """
    Returns list of form metadata dicts to generate for a given stage ('PRE-APPROVAL' or 'POST-APPROVAL').
    """
    stage_upper = stage.strip().upper()
    forms = []

    if stage_upper == "PRE-APPROVAL":
        # 1. CTS (HLF-161)
        forms.append({
            "code": "HLF-161",
            "name": "Contract-to-Sell",
            "template_name": "HLF161_CTSFinancingAffordableHousingProgram_V06.pdf"
        })
        # 2. Deed of Assignment (HLF-524)
        forms.append({
            "code": "HLF-524",
            "name": "Deed of Assignment",
            "template_name": "HLF523_DeedAssignmentCTSSPAUnilateral_V03.pdf"
        })
        # 3. HLA Acknowledgement Receipt (HLF-785)
        forms.append({
            "code": "HLF-785",
            "name": "HLA Acknowledgement Receipt",
            "template_name": "HLF785_HLAAcknowledgementReceiptWindow2_V02.pdf"
        })
        # 4. MSVS (HLF-063, one per borrower)
        forms.append({
            "code": "HLF-063",
            "name": "MSVS - Principal",
            "borrower_type": "principal",
            "template_name": "HLF063_MembershipStatusVerificationSlip_V04.pdf"
        })
        if buyer.has_co_borrower:
            forms.append({
                "code": "HLF-063",
                "name": "MSVS - Co-Borrower",
                "borrower_type": "co_borrower",
                "template_name": "HLF063_MembershipStatusVerificationSlip_V04.pdf"
            })
        # 5. CIBI/Data Privacy (one per borrower)
        forms.append({
            "code": "CIBI",
            "name": "CIBI - Principal",
            "borrower_type": "principal",
            "template_name": "Authorization To Conduct CIBI (data privacy act).pdf"
        })
        if buyer.has_co_borrower:
            forms.append({
                "code": "CIBI",
                "name": "CIBI - Co-Borrower",
                "borrower_type": "co_borrower",
                "template_name": "Authorization To Conduct CIBI (data privacy act).pdf"
            })
        # 6. Borrower's Validation Sheet (HLF-058, one per borrower)
        forms.append({
            "code": "HLF-058",
            "name": "Borrower's Validation Sheet - Principal",
            "borrower_type": "principal",
            "template_name": "HLF058_BVS- DA_V03.pdf"
        })
        if buyer.has_co_borrower:
            forms.append({
                "code": "HLF-058",
                "name": "Borrower's Validation Sheet - Co-Borrower",
                "borrower_type": "co_borrower",
                "template_name": "HLF058_BVS- DA_V03.pdf"
            })
        # 7. ARIS (HLF-913): ONLY when attorney-in-fact exists
        if buyer.has_attorney_in_fact:
            forms.append({
                "code": "HLF-913",
                "name": "ARIS",
                "template_name": "HLF913_ARIS.pdf"
            })

    elif stage_upper == "POST-APPROVAL":
        # 1. Deed of Absolute Sale (HLF-525)
        forms.append({
            "code": "HLF-525",
            "name": "Deed of Absolute Sale",
            "template_name": "HLF525_DeedAbsoluteSale_V02.pdf"
        })
        # 2. LMA & PN: loanable amount below 950,000 -> Affordable; 950,000+ -> End-User
        if buyer.is_affordable:
            forms.append({
                "code": "HLF-163",
                "name": "Loan and Mortgage Agreement",
                "program": "Affordable",
                "template_name": "HLF163_LoanMortgageAgreementAffordableHousing_V05.pdf"
            })
            forms.append({
                "code": "HLF-087",
                "name": "Promissory Note",
                "program": "Affordable",
                "template_name": "HLF087_PromissoryNoteAffordableHousingProgram_V06.pdf"
            })
        else:
            forms.append({
                "code": "HLF-162",
                "name": "Loan and Mortgage Agreement",
                "program": "End-User",
                "template_name": "HLF162_LoanMortgageAgreementHomeFinancing_V06.pdf"
            })
            forms.append({
                "code": "HLF-086",
                "name": "Promissory Note",
                "program": "End-User",
                "template_name": "HLF086_PromissoryNoteHomeFinancingProgram_V06.pdf"
            })
        # 3. Certificate of House & Lot Acceptance
        forms.append({
            "code": "HLA-ACCEPT",
            "name": "Certificate of House & Lot Acceptance",
            "template_name": "Certificate of house & lot acceptance.pdf"
        })
        # 4. Certificate of Acceptance (HLF-083)
        forms.append({
            "code": "HLF-083",
            "name": "Certificate of Acceptance",
            "template_name": "HLF083_CertificateOfAcceptance.pdf"
        })
        # 5. Buyer's Conformity (HLF-371)
        forms.append({
            "code": "HLF-371",
            "name": "Buyer's Conformity",
            "template_name": "HLF371_BuyersConformity.pdf"
        })

    return forms
