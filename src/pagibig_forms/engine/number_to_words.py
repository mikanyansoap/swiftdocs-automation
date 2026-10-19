"""
Converts numbers, currency, interest rates, and loan terms to standard Filipino legal wording.
"""
from __future__ import annotations
import math

ONES = [
    "", "ONE", "TWO", "THREE", "FOUR", "FIVE", "SIX", "SEVEN", "EIGHT", "NINE",
    "TEN", "ELEVEN", "TWELVE", "THIRTEEN", "FOURTEEN", "FIFTEEN", "SIXTEEN",
    "SEVENTEEN", "EIGHTEEN", "NINETEEN"
]

TENS = [
    "", "", "TWENTY", "THIRTY", "FORTY", "FIFTY", "SIXTY", "SEVENTY", "EIGHTY", "NINETY"
]

SCALES = ["", "THOUSAND", "MILLION", "BILLION", "TRILLION"]


def integer_to_words(n: int) -> str:
    """Converts a positive integer to uppercase words."""
    if n == 0:
        return "ZERO"
    if n < 0:
        return "NEGATIVE " + integer_to_words(abs(n))

    def convert_hundred(val: int) -> str:
        words = []
        if val >= 100:
            words.append(ONES[val // 100] + " HUNDRED")
            val %= 100
        if 1 <= val < 20:
            words.append(ONES[val])
        elif val >= 20:
            w = TENS[val // 10]
            if val % 10:
                w += " " + ONES[val % 10]
            words.append(w)
        return " ".join(words)

    groups = []
    scale_idx = 0
    temp = n
    while temp > 0:
        group_val = temp % 1000
        if group_val != 0:
            g_str = convert_hundred(group_val)
            if SCALES[scale_idx]:
                g_str += " " + SCALES[scale_idx]
            groups.insert(0, g_str)
        temp //= 1000
        scale_idx += 1

    return " ".join(groups).strip()


def amount_to_words(amount: float, currency_word: str = "PESOS") -> str:
    """
    Format: 'ONE MILLION PESOS AND 00/100'
    Centavos formatted as 'AND xx/100'.
    """
    if amount is None or math.isnan(amount):
        return ""
    total_cents = int(round(amount * 100))
    pesos = total_cents // 100
    cents = total_cents % 100

    pesos_words = integer_to_words(pesos)
    if not pesos_words:
        pesos_words = "ZERO"

    cur = currency_word.upper().strip()
    if cur:
        return f"{pesos_words} {cur} AND {cents:02d}/100"
    return f"{pesos_words} AND {cents:02d}/100"


def format_currency_figures(amount: float, symbol: str = "P") -> str:
    """e.g. 1250000.0 -> 'P 1,250,000.00' or '1,250,000.00'"""
    if amount is None or math.isnan(amount):
        return ""
    s = f"{amount:,.2f}"
    if symbol:
        return f"{symbol} {s}".strip()
    return s


def amount_words_and_figures(amount: float, currency_word: str = "PESOS", symbol: str = "P") -> str:
    """Convenience combination: 'ONE MILLION PESOS AND 00/100 (P 1,000,000.00)'"""
    words = amount_to_words(amount, currency_word)
    figs = format_currency_figures(amount, symbol)
    return f"{words} ({figs})"


def interest_rate_to_words_and_figures(rate: float) -> str:
    """
    Given rate (e.g. 0.0625 or 6.25):
    prints in words and figures: 'SIX PT. TWO HUNDRED FIFTY (6.250%)'
    """
    if rate is None or math.isnan(rate):
        return ""
    percent = rate * 100.0 if rate < 1.0 else rate
    thousandths = int(round(percent * 1000))
    int_part, dec_part = divmod(thousandths, 1000)
    percent = thousandths / 1000.0

    int_words = integer_to_words(int_part)
    if dec_part > 0:
        dec_words = integer_to_words(dec_part)
        words_part = f"{int_words} PT. {dec_words}"
    else:
        words_part = f"{int_words} PERCENT"

    figures_part = f"({percent:.3f}%)"
    return f"{words_part} {figures_part}"


def area_to_words_and_figures(sqm: float) -> str:
    """
    e.g. 36.0 -> 'THIRTY SIX (36) SQUARE METERS more or less.'
    """
    if sqm is None or math.isnan(sqm):
        return ""
    int_val = int(round(sqm))
    words = integer_to_words(int_val)
    return f"{words} ({int_val}) SQUARE METERS more or less."


def years_to_months(years: int) -> int:
    return int(years) * 12


def term_to_words_and_figures(years: int, unit: str = "years") -> str:
    """
    Returns e.g. 'THIRTY (30) YEARS' or 'THREE HUNDRED SIXTY (360) MONTHS'.
    """
    if unit == "months":
        months = years_to_months(years)
        return f"{integer_to_words(months)} ({months}) {'MONTH' if months == 1 else 'MONTHS'}"
    return f"{integer_to_words(years)} ({years}) {'YEAR' if years == 1 else 'YEARS'}"
