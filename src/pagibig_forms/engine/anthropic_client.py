"""
Anthropic Messages API client for Housing Loan Application (HLA) scanning.
Sends images/PDF to api.anthropic.com through the official Anthropic SDK (HTTPS).
Parses ONLY the whitelisted JSON keys with the exact prompt verbatim.
"""
from __future__ import annotations
import base64
import io
import json
import os
import re
from typing import Any, Dict, List, Optional, Tuple

import importlib.util

from .models import BuyerRecord, parse_float, parse_int
from ..settings import DEFAULT_MODEL

# The Anthropic SDK and Pillow are imported only when an HLA is scanned: loading the SDK
# takes several seconds, which would otherwise slow down every app start.
# Pillow is bundled on Windows; Android works without it (JPEG/PDF only).
HAS_PIL = importlib.util.find_spec("PIL") is not None

HLA_PROMPT_VERBATIM = (
    'These images are a Philippine Pag-IBIG Fund HOUSING LOAN APPLICATION '
    '(form HQP-HLF-068, sometimes with a co-borrower form HQP-HLF-069). '
    'Entries are often handwritten. Read them carefully and reply with ONLY one '
    'JSON object with these keys (use "" when not shown or unreadable; never guess): '
    'A (borrower FIRST MIDDLE LAST-EXT in capitals), B ("Single" or "Married to " + '
    'spouse full name; else status as ticked), C (co-borrower in capitals, else "N/A"), '
    'coStatus, D (present address on one line in capitals), E ("BLOCK n LOT n"), '
    'J (desired loan amount as a number), M, N, O (TINs exactly as written), '
    'termYears, repricing, houseType, lotArea, email, '
    'flags (array of short notes about anything unreadable, incomplete or worth checking).'
)

ALLOWED_KEYS = {
    "A", "B", "C", "coStatus", "D", "E", "J", "M", "N", "O",
    "termYears", "repricing", "houseType", "lotArea", "email", "flags"
}

# API limits: 5 MB per image. Photos are scaled so handwriting stays legible
# without sending more pixels than the model uses.
MAX_IMAGE_BYTES = 5 * 1024 * 1024
MAX_IMAGE_EDGE = 2000
IMAGE_EXTS = (".jpg", ".jpeg", ".png", ".webp")


def _load_image_as_jpeg(path: str, max_edge: Optional[int]) -> Tuple[bytes, int, int]:
    """Returns (jpeg_bytes, width, height), upright and optionally scaled down. Needs Pillow."""
    from PIL import Image, ImageOps
    with Image.open(path) as img:
        img = ImageOps.exif_transpose(img).convert("RGB")
        if max_edge and max(img.size) > max_edge:
            img.thumbnail((max_edge, max_edge))
        buf = io.BytesIO()
        img.save(buf, "JPEG", quality=85)
        return buf.getvalue(), img.width, img.height


def _image_block(path: str) -> Dict[str, Any]:
    ext = os.path.splitext(path)[1].lower()
    if HAS_PIL:
        data, _w, _h = _load_image_as_jpeg(path, MAX_IMAGE_EDGE)
        media_type = "image/jpeg"
    else:
        with open(path, "rb") as f:
            data = f.read()
        media_type = {".png": "image/png", ".webp": "image/webp"}.get(ext, "image/jpeg")
    if len(data) > MAX_IMAGE_BYTES:
        raise ValueError(f"{os.path.basename(path)} is larger than 5 MB. Retake the photo at a lower resolution.")
    return {
        "type": "image",
        "source": {"type": "base64", "media_type": media_type, "data": base64.standard_b64encode(data).decode("ascii")}
    }


class AnthropicHLAClient:
    def __init__(self, api_key: str, model: str = DEFAULT_MODEL):
        self.api_key = api_key.strip()
        self.model = model.strip() or DEFAULT_MODEL

    def process_hla_files(self, file_paths: List[str]) -> Tuple[Dict[str, Any], List[str]]:
        """
        Processes images/PDFs through the Anthropic Messages API.
        Blocking network call: run it off the UI thread.
        Returns (parsed_data_dict, flags_list).
        """
        if not self.api_key:
            raise ValueError("Anthropic API key is required. Please set it in Settings.")

        # Build content blocks (files first, then the prompt)
        content_blocks = []
        for path in file_paths:
            ext = os.path.splitext(path)[1].lower()
            if ext in IMAGE_EXTS:
                content_blocks.append(_image_block(path))
            elif ext == ".pdf":
                with open(path, "rb") as f:
                    b64_data = base64.standard_b64encode(f.read()).decode("ascii")
                content_blocks.append({
                    "type": "document",
                    "source": {"type": "base64", "media_type": "application/pdf", "data": b64_data}
                })
        if not content_blocks:
            raise ValueError("Choose HLA photos (JPG/PNG) or a PDF.")

        content_blocks.append({"type": "text", "text": HLA_PROMPT_VERBATIM})

        import anthropic
        client = anthropic.Anthropic(api_key=self.api_key, timeout=180.0, max_retries=2)
        try:
            response = client.beta.messages.create(
                model=self.model,
                max_tokens=16000,
                # If a safety check declines the request, let the API retry on a fallback model
                betas=["server-side-fallback-2026-07-01"],
                fallbacks="default",
                messages=[{"role": "user", "content": content_blocks}],
            )
        except anthropic.AuthenticationError:
            raise RuntimeError("The Anthropic API key was rejected. Check it in Settings.")
        except anthropic.NotFoundError:
            raise RuntimeError(f"Model '{self.model}' was not found. Check the model name in Settings.")
        except anthropic.RateLimitError:
            raise RuntimeError("Too many requests to the Anthropic API. Wait a minute and try again.")
        except anthropic.APIStatusError as e:
            raise RuntimeError(f"Anthropic API error ({e.status_code}): {e.message}")
        except anthropic.APIConnectionError:
            raise RuntimeError("Could not reach api.anthropic.com. Check the internet connection.")

        if response.stop_reason == "refusal":
            raise RuntimeError("The model declined to read these files. Please type the details manually.")

        raw_text = "".join(block.text for block in response.content if block.type == "text")
        return self.parse_and_whitelist_json(raw_text)

    def parse_and_whitelist_json(self, response_text: str) -> Tuple[Dict[str, Any], List[str]]:
        """
        Extracts ONLY the JSON object and whitelists allowed keys.
        """
        # Find JSON substring
        match = re.search(r"\{[\s\S]*\}", response_text)
        if not match:
            raise ValueError("No valid JSON object found in model response.")

        parsed = json.loads(match.group(0))
        cleaned = {}
        for k in ALLOWED_KEYS:
            cleaned[k] = parsed.get(k, "" if k != "flags" else [])

        flags = cleaned.get("flags", [])
        if not isinstance(flags, list):
            flags = [str(flags)] if flags else []
        flags = [str(f) for f in flags]

        # The HLA shows the amount the buyer *wants*; the approved loanable amount
        # comes from Pag-IBIG later, so it is reported here instead of filled in.
        desired = parse_float(cleaned.get("J", ""), 0.0)
        if desired > 0:
            flags.append(f"HLA desired loan amount: P {desired:,.2f}. Enter the APPROVED loanable amount yourself.")

        return cleaned, flags


def _text_or(value: Any, default: str = "") -> str:
    s = str(value if value is not None else "").strip()
    return s or default


def map_hla_json_to_buyer(hla_dict: Dict[str, Any], row_idx: int = 0) -> BuyerRecord:
    """Converts whitelisted HLA JSON dictionary to a BuyerRecord.
    The loanable amount is left at 0: it must be the approved amount, typed by staff."""
    return BuyerRecord(
        row_idx=row_idx,
        principal_borrower=_text_or(hla_dict.get("A")),
        status=_text_or(hla_dict.get("B"), "Single"),
        co_borrower=_text_or(hla_dict.get("C"), "N/A"),
        address=_text_or(hla_dict.get("D")),
        block_lot=_text_or(hla_dict.get("E")),
        loanable_amount=0.0,
        tin_principal=_text_or(hla_dict.get("M")),
        tin_co_borrower=_text_or(hla_dict.get("N")),
        tin_spouse=_text_or(hla_dict.get("O")),
        loan_term_years=parse_int(hla_dict.get("termYears", 30), 30),
        repricing_years=parse_int(hla_dict.get("repricing", 1), 1),
        house_type=_text_or(hla_dict.get("houseType"), "TOWNHOUSE"),
        lot_area=parse_float(hla_dict.get("lotArea", 36.0), 36.0),
        co_borrower_status=_text_or(hla_dict.get("coStatus")),
        email=_text_or(hla_dict.get("email"))
    )


def _jpeg_size(data: bytes) -> Optional[Tuple[int, int]]:
    """Reads (width, height) from a JPEG's SOF marker."""
    i = 2
    while i + 9 < len(data):
        if data[i] != 0xFF:
            return None
        marker = data[i + 1]
        length = int.from_bytes(data[i + 2:i + 4], "big")
        if marker in (0xC0, 0xC1, 0xC2):
            return int.from_bytes(data[i + 7:i + 9], "big"), int.from_bytes(data[i + 5:i + 7], "big")
        i += 2 + length
    return None


def _jpeg_page_pdf(jpeg: bytes, width: int, height: int) -> bytes:
    """One Letter page with the JPEG scaled to fit inside a 36 pt margin."""
    page_w, page_h, margin = 612.0, 792.0, 36.0
    scale = min((page_w - 2 * margin) / width, (page_h - 2 * margin) / height)
    w, h = width * scale, height * scale
    x, y = (page_w - w) / 2, (page_h - h) / 2
    content = f"q {w:.2f} 0 0 {h:.2f} {x:.2f} {y:.2f} cm /Im1 Do Q".encode("ascii")
    objects = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        (f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {page_w:.0f} {page_h:.0f}] "
         f"/Resources << /XObject << /Im1 5 0 R >> >> /Contents 4 0 R >>").encode("ascii"),
        f"<< /Length {len(content)} >>\nstream\n".encode("ascii") + content + b"\nendstream",
        (f"<< /Type /XObject /Subtype /Image /Width {width} /Height {height} /ColorSpace /DeviceRGB "
         f"/BitsPerComponent 8 /Filter /DCTDecode /Length {len(jpeg)} >>\nstream\n").encode("ascii") + jpeg + b"\nendstream",
    ]
    out = io.BytesIO()
    out.write(b"%PDF-1.4\n")
    offsets = []
    for n, obj in enumerate(objects, 1):
        offsets.append(out.tell())
        out.write(f"{n} 0 obj\n".encode("ascii") + obj + b"\nendobj\n")
    xref = out.tell()
    out.write(f"xref\n0 {len(objects) + 1}\n0000000000 65535 f \n".encode("ascii"))
    for off in offsets:
        out.write(f"{off:010d} 00000 n \n".encode("ascii"))
    out.write(f"trailer\n<< /Size {len(objects) + 1} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF\n".encode("ascii"))
    return out.getvalue()


def join_photos_and_pdfs_to_pdf(input_paths: List[str], output_pdf_path: str):
    """
    Joins photos and PDFs into one PDF, keeping the actual photos.
    Uses Pillow when available (any photo format); otherwise embeds JPEGs directly.
    """
    import pypdf
    from .canvas import PdfCanvas

    writer = pypdf.PdfWriter()

    for path in input_paths:
        ext = os.path.splitext(path)[1].lower()
        if ext == ".pdf":
            reader = pypdf.PdfReader(path)
            for page in reader.pages:
                writer.add_page(page)
        elif ext in IMAGE_EXTS:
            page_pdf = None
            if HAS_PIL:
                jpeg, w, h = _load_image_as_jpeg(path, 2500)
                page_pdf = _jpeg_page_pdf(jpeg, w, h)
            elif ext in (".jpg", ".jpeg"):
                with open(path, "rb") as f:
                    jpeg = f.read()
                size = _jpeg_size(jpeg)
                if size:
                    page_pdf = _jpeg_page_pdf(jpeg, *size)
            if page_pdf is None:
                canvas = PdfCanvas(width=612.0, height=792.0)
                canvas.draw_text(f"Could not embed photo: {os.path.basename(path)} (keep the original file)", 306.0, 400.0,
                                 font_name="Helvetica-Bold", font_size=11.0, align="center")
                writer.add_page(canvas.to_page_object())
            else:
                writer.add_page(pypdf.PdfReader(io.BytesIO(page_pdf)).pages[0])

    os.makedirs(os.path.dirname(os.path.abspath(output_pdf_path)), exist_ok=True)
    with open(output_pdf_path, "wb") as f:
        writer.write(f)
