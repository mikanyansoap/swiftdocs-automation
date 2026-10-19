"""
Pure-Python in-memory PDF Canvas builder.
Constructs valid standard PDF streams without any external dependencies (like reportlab).
Supports Helvetica, Helvetica-Bold, Courier, Courier-Bold, Times-Roman,
text scaling, wrapping, multi-line, vertical side margins, and lines.
"""
from __future__ import annotations
import io
import re
from typing import List, Tuple, Optional
import pypdf

# Standard font approximate character widths relative to 1000 units
CHAR_WIDTHS_HELVETICA = {
    ' ': 278, '!': 278, '"': 355, '#': 556, '$': 556, '%': 889, '&': 667, "'": 191,
    '(': 333, ')': 333, '*': 389, '+': 584, ',': 278, '-': 333, '.': 278, '/': 278,
    '0': 556, '1': 556, '2': 556, '3': 556, '4': 556, '5': 556, '6': 556, '7': 556,
    '8': 556, '9': 556, ':': 278, ';': 278, '<': 584, '=': 584, '>': 584, '?': 556,
    '@': 1015, 'A': 667, 'B': 667, 'C': 722, 'D': 722, 'E': 667, 'F': 611, 'G': 778,
    'H': 722, 'I': 278, 'J': 500, 'K': 667, 'L': 556, 'M': 833, 'N': 722, 'O': 778,
    'P': 667, 'Q': 778, 'R': 722, 'S': 667, 'T': 611, 'U': 722, 'V': 667, 'W': 944,
    'X': 667, 'Y': 667, 'Z': 611, '[': 278, '\\': 278, ']': 278, '^': 469, '_': 556,
    '`': 333, 'a': 556, 'b': 556, 'c': 500, 'd': 556, 'e': 556, 'f': 278, 'g': 556,
    'h': 556, 'i': 222, 'j': 222, 'k': 500, 'l': 222, 'm': 833, 'n': 556, 'o': 556,
    'p': 556, 'q': 556, 'r': 333, 's': 500, 't': 278, 'u': 556, 'v': 500, 'w': 722,
    'x': 500, 'y': 500, 'z': 500, '{': 334, '|': 260, '}': 334, '~': 584
}

FONT_MAP = {
    "Helvetica": "/F1",
    "Helvetica-Bold": "/F2",
    "Courier": "/F3",
    "Courier-Bold": "/F4",
    "Times-Roman": "/F5",
}


# Characters the standard PDF fonts (WinAnsiEncoding) cannot show
CHAR_REPLACEMENTS = {"₱": "P", "–": "-", "—": "-", "‘": "'", "’": "'", "“": '"', "”": '"'}


def to_pdf_text(s: str) -> str:
    """Maps text onto characters the WinAnsi-encoded standard fonts can display (keeps Ñ, É, etc.)."""
    for src, dst in CHAR_REPLACEMENTS.items():
        s = s.replace(src, dst)
    return s.encode("cp1252", errors="replace").decode("cp1252")


def escape_pdf_string(s: str) -> str:
    """Escapes special characters in PDF literal string."""
    s = to_pdf_text(s)
    return s.replace('\\', '\\\\').replace('(', '\\(').replace(')', '\\)')


def measure_text_width(text: str, font_name: str, font_size: float) -> float:
    """Estimates text width in points."""
    if not text:
        return 0.0
    if "Courier" in font_name:
        return len(text) * 0.60 * font_size
    total = 0.0
    for char in text:
        w = CHAR_WIDTHS_HELVETICA.get(char, 550)
        total += w
    if "Bold" in font_name:
        total *= 1.06  # Helvetica-Bold is slightly wider than regular
    return (total / 1000.0) * font_size


def wrap_words(text: str, font_name: str, font_size: float, max_width: float) -> List[str]:
    """Splits text into lines that each fit within max_width."""
    lines: List[str] = []
    cur = ""
    for word in text.split():
        trial = f"{cur} {word}".strip()
        if cur and measure_text_width(trial, font_name, font_size) > max_width:
            lines.append(cur)
            cur = word
        else:
            cur = trial
    if cur:
        lines.append(cur)
    return lines


class PdfCanvas:
    """
    Lightweight, pure-Python PDF canvas for generating overlay pages and template pages.
    """
    def __init__(self, width: float = 612.0, height: float = 792.0):
        self.width = width
        self.height = height
        self.stream_commands: List[str] = []

    def set_color_rgb(self, r: float, g: float, b: float, stroke: bool = False):
        cmd = f"{r:.3f} {g:.3f} {b:.3f} RG" if stroke else f"{r:.3f} {g:.3f} {b:.3f} rg"
        self.stream_commands.append(cmd)

    def draw_text(
        self,
        text: str,
        x: float,
        y: float,
        font_name: str = "Helvetica",
        font_size: float = 9.0,
        max_width: Optional[float] = None,
        min_font_size: float = 6.0,
        align: str = "left",
        wrap: bool = True
    ):
        """
        Draws text at (x, y) with auto-shrinking if max_width is exceeded.
        If even the minimum size does not fit, the text wraps onto lines below
        (unless wrap=False) instead of running past the edge of the page.
        align can be 'left', 'center', or 'right'.
        """
        if not text:
            return

        clean_text = to_pdf_text(str(text).strip())
        actual_size = font_size
        if max_width:
            w = measure_text_width(clean_text, font_name, actual_size)
            if w > max_width:
                scale = max_width / w
                actual_size = max(min_font_size, font_size * scale)
                if wrap and measure_text_width(clean_text, font_name, actual_size) > max_width:
                    wrap_size = max(min_font_size, min(font_size, 7.0))
                    lines = wrap_words(clean_text, font_name, wrap_size, max_width)
                    for i, line in enumerate(lines):
                        self.draw_text(line, x, y - i * wrap_size * 1.15, font_name=font_name,
                                       font_size=wrap_size, max_width=max_width,
                                       min_font_size=min_font_size, align=align, wrap=False)
                    return

        w = measure_text_width(clean_text, font_name, actual_size)
        x_pos = x
        if align == "center":
            x_pos = x - (w / 2.0)
        elif align == "right":
            x_pos = x - w

        f_tag = FONT_MAP.get(font_name, "/F1")
        escaped = escape_pdf_string(clean_text)
        cmd = (
            f"BT\n"
            f"{f_tag} {actual_size:.2f} Tf\n"
            f"1 0 0 1 {x_pos:.2f} {y:.2f} Tm\n"
            f"({escaped}) Tj\n"
            f"ET"
        )
        self.stream_commands.append(cmd)

    def draw_multiline_text(
        self,
        text: str,
        x: float,
        top_y: float,
        font_name: str = "Courier",
        font_size: float = 9.0,
        line_height: float = 11.0,
        max_width: Optional[float] = None,
        bottom_y: Optional[float] = None,
        min_font_size: float = 5.5
    ):
        """
        Draws multi-line text downwards from top_y.
        If bottom_y is given, font size and line spacing shrink so every line stays above it.
        """
        if not text:
            return
        lines = text.split("\n")
        if bottom_y is not None and len(lines) > 1:
            available = top_y - bottom_y
            needed = (len(lines) - 1) * line_height
            if needed > available:
                scale = max(min_font_size / font_size, available / needed)
                font_size *= scale
                line_height *= scale
        cur_y = top_y
        for line in lines:
            self.draw_text(line, x, cur_y, font_name=font_name, font_size=font_size,
                           max_width=max_width, min_font_size=min(min_font_size, font_size), wrap=False)
            cur_y -= line_height

    def draw_vertical_text(
        self,
        text: str,
        x: float,
        y: float,
        font_name: str = "Helvetica-Bold",
        font_size: float = 8.0
    ):
        """
        Draws text rotated 90 degrees vertically (reading bottom-to-top or top-to-bottom).
        Used for side-margin names.
        Transformation matrix for 90 deg counter-clockwise: [0 1 -1 0 x y]
        """
        if not text:
            return
        f_tag = FONT_MAP.get(font_name, "/F1")
        escaped = escape_pdf_string(text)
        cmd = (
            f"BT\n"
            f"{f_tag} {font_size:.2f} Tf\n"
            f"0 1 -1 0 {x:.2f} {y:.2f} Tm\n"
            f"({escaped}) Tj\n"
            f"ET"
        )
        self.stream_commands.append(cmd)

    def draw_line(self, x1: float, y1: float, x2: float, y2: float, width: float = 0.5):
        """Draws a line between (x1, y1) and (x2, y2)."""
        cmd = (
            f"{width:.2f} w\n"
            f"{x1:.2f} {y1:.2f} m\n"
            f"{x2:.2f} {y2:.2f} l\n"
            f"S"
        )
        self.stream_commands.append(cmd)

    def draw_rect(self, x: float, y: float, w: float, h: float, fill: bool = False, stroke: bool = True):
        """Draws a rectangle at (x, y)."""
        op = "B" if (fill and stroke) else ("f" if fill else "s")
        cmd = f"{x:.2f} {y:.2f} {w:.2f} {h:.2f} re\n{op}"
        self.stream_commands.append(cmd)

    def build_pdf_bytes(self) -> bytes:
        """Assembles a valid, compliant single-page PDF containing the drawn content."""
        content = "\n".join(self.stream_commands).encode("cp1252", errors="replace")
        content_len = len(content)

        objects = []
        # Obj 1: Catalog
        objects.append(b"1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n")
        # Obj 2: Pages
        objects.append(b"2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n")
        # Obj 3: Page
        page_dict = (
            f"3 0 obj\n"
            f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {self.width:.2f} {self.height:.2f}] "
            f"/Resources << /Font << "
            f"/F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >> "
            f"/F2 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >> "
            f"/F3 << /Type /Font /Subtype /Type1 /BaseFont /Courier /Encoding /WinAnsiEncoding >> "
            f"/F4 << /Type /Font /Subtype /Type1 /BaseFont /Courier-Bold /Encoding /WinAnsiEncoding >> "
            f"/F5 << /Type /Font /Subtype /Type1 /BaseFont /Times-Roman /Encoding /WinAnsiEncoding >> "
            f">> >> /Contents 4 0 R >>\nendobj\n"
        ).encode("latin-1")
        objects.append(page_dict)
        # Obj 4: Contents
        contents_dict = (
            f"4 0 obj\n<< /Length {content_len} >>\nstream\n"
        ).encode("latin-1") + content + b"\nendstream\nendobj\n"
        objects.append(contents_dict)

        header = b"%PDF-1.4\n"
        body = b""
        offsets = [0]
        cur_offset = len(header)

        for obj in objects:
            offsets.append(cur_offset)
            body += obj
            cur_offset += len(obj)

        xref_pos = len(header) + len(body)
        xref = f"xref\n0 {len(offsets)}\n0000000000 65535 f \n"
        for off in offsets[1:]:
            xref += f"{off:010d} 00000 n \n"

        trailer = (
            f"trailer\n<< /Size {len(offsets)} /Root 1 0 R >>\n"
            f"startxref\n{xref_pos}\n%%EOF\n"
        )

        return header + body + xref.encode("latin-1") + trailer.encode("latin-1")

    def to_page_object(self) -> pypdf.PageObject:
        """Converts this canvas to a pypdf.PageObject ready for merge_page()."""
        pdf_data = self.build_pdf_bytes()
        reader = pypdf.PdfReader(io.BytesIO(pdf_data))
        return reader.pages[0]
