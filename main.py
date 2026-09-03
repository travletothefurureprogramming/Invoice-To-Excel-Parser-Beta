

import datetime as dt
import math
import os
import queue
import re
import tempfile
import threading
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable, Optional

import customtkinter as ctk
import pandas as pd
import pdfplumber
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# ============================================================================
# APPLICATION CONFIGURATION
# ============================================================================

APP_NAME = "Invoice2Excel Pro"
APP_VERSION = "1.0.0"

SUPPORTED_EXTENSIONS = {".pdf"}

OUTPUT_COLUMNS = [
    "Invoice Date",
    "Tax ID / VAT",
    "Net Amount (€/$)",
    "VAT Amount (€/$)",
    "Total Amount (€/$)",
    "File Name",
    "Status",
]


# ============================================================================
# TEXT NORMALIZATION
# ============================================================================

EMAIL_RE = re.compile(
    r"\b[A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,}\b",
    re.IGNORECASE,
)

URL_RE = re.compile(
    r'\b(?:https?://|www\.)[^\s<>"\']+',
    re.IGNORECASE,
)


def normalize_text(text: str) -> str:
    """Normalize Unicode and common OCR whitespace artifacts."""
    if not text:
        return ""

    text = unicodedata.normalize("NFKC", text)

    text = text.replace("\u00A0", " ")
    text = text.replace("\u200B", "")
    text = text.replace("\uFEFF", "")

    text = text.replace("\r\n", "\n")
    text = text.replace("\r", "\n")

    return "\n".join(
        re.sub(r"[ \t]+", " ", line).strip()
        for line in text.splitlines()
    )


def strip_emails_and_urls(text: str) -> str:
    """
    Remove email addresses and URLs before Tax ID extraction.

    This prevents:
        elon@elon.com
        accounts@example.de
        https://example.com/DE123456789
    from becoming false VAT candidates.
    """
    if not text:
        return ""

    text = EMAIL_RE.sub(" ", text)
    text = URL_RE.sub(" ", text)

    return text


def fold_text(text: str) -> str:
    """
    Case/diacritic/punctuation-insensitive representation.

    Examples:
        Φ.Π.Α.       -> φπα
        ΦΠΑ          -> φπα
        ΚΑΘΑΡΗ ΑΞΙΑ -> καθαρηαξια
    """
    value = unicodedata.normalize(
        "NFKD",
        text or "",
    )

    value = "".join(
        ch
        for ch in value
        if not unicodedata.combining(ch)
    )

    value = value.casefold()

    return re.sub(
        r"[^0-9a-zα-ω]+",
        "",
        value,
    )


# ============================================================================
# TAX ID / VAT
# ============================================================================

# Body length AFTER the country prefix.
EU_VAT_BODY_LENGTHS = {
    "AT": {9},
    "BE": {10},
    "BG": {9, 10},
    "CY": {9},
    "CZ": {8, 9, 10},
    "DE": {9},
    "DK": {8},
    "EE": {9},
    "EL": {9},
    "ES": {9},
    "FI": {8},
    "FR": {11},
    "GB": {9, 12},
    "GR": {9},
    "HR": {11},
    "HU": {8},
    "IE": {8},
    "IT": {11},
    "LT": {9, 12},
    "LU": {8},
    "LV": {11},
    "MT": {8},
    "NL": {12},
    "PL": {10},
    "PT": {9},
    "RO": set(range(2, 11)),
    "SE": {12},
    "SI": {8},
    "SK": {10},
}

TAX_ANCHOR_RE = re.compile(
    r"""
    (?:
        A\s*\.?\s*F\s*\.?\s*M\.?
        |
        Α\s*\.?\s*Φ\s*\.?\s*Μ\.?
        |
        V\s*A\s*T
        |
        U\s*S?\s*T\s*[-.]?\s*I\s*D\s*N\s*R\.?
        |
        T\s*V\s*A
        |
        E\s*I\s*N
        |
        T\s*A\s*X\s*I\s*D
    )
    \s*(?:NO|NUMBER|NUM|ID|NR)?
    \s*[#:\-]?\s*
    """,
    re.IGNORECASE | re.VERBOSE,
)

# Greek AFM:
# STRICTLY exactly 9 contiguous digits after the label.
#
# It intentionally does NOT allow:
#   123 456 789
#   123-456-789
#
# because the test requirement specifically says the anchored Greek AFM
# must match exactly \b\d{9}\b and must never concatenate adjacent IDs.
GREEK_AFM_ANCHORED_RE = re.compile(
    r"""
    (?:
        Α\s*\.?\s*Φ\s*\.?\s*Μ\.?
        |
        A\s*\.?\s*F\s*\.?\s*M\.?
    )
    \s*[#:\-]?\s*
    \b(\d{9})\b
    """,
    re.IGNORECASE | re.VERBOSE,
)

EIN_ANCHORED_RE = re.compile(
    r"""
    \bE\s*I\s*N\s*[#:\-]?\s*
    (\d{2})
    \s*-\s*
    (\d{7})
    \b
    """,
    re.IGNORECASE | re.VERBOSE,
)

EIN_PLAIN_RE = re.compile(
    r"""
    \bE\s*I\s*N\s*[#:\-]?\s*
    (\d{9})
    \b
    """,
    re.IGNORECASE | re.VERBOSE,
)


def compact_identifier(value: str) -> str:
    """
    Remove separators introduced by OCR.

    FR 12 345 678 901
        ->
    FR12345678901
    """
    value = unicodedata.normalize(
        "NFKC",
        value or "",
    )

    return re.sub(
        r"[\s.\-_/():;]+",
        "",
        value,
    ).upper()


def is_valid_tax_identifier(
    candidate: str,
) -> bool:
    """
    Validate a tax identifier structurally.

    Supported:
        - bare 9-digit AFM/EIN
        - EU/UK VAT with country prefix
    """
    candidate = compact_identifier(
        candidate
    )

    if not candidate:
        return False

    # Bare 9-digit tax identifier.
    if re.fullmatch(
        r"\d{9}",
        candidate,
    ):
        return True

    if len(candidate) < 4:
        return False

    prefix = candidate[:2]
    body = candidate[2:]

    if prefix not in EU_VAT_BODY_LENGTHS:
        return False

    if len(body) not in EU_VAT_BODY_LENGTHS[prefix]:
        return False

    return bool(
        re.fullmatch(
            r"[A-Z0-9]+",
            body,
        )
    )


def _scan_vat_after_position(
    text: str,
    start: int,
) -> Optional[str]:
    """
    Search a bounded region for an EU VAT prefix + body.

    The prefix is always preserved.
    """
    window = text[
        start:start + 40
    ]

    country_pattern = "|".join(
        sorted(
            EU_VAT_BODY_LENGTHS,
            key=len,
            reverse=True,
        )
    )

    country_re = re.compile(
        rf"\b({country_pattern})\b",
        re.IGNORECASE,
    )

    for match in country_re.finditer(
        window
    ):
        prefix = match.group(1).upper()

        raw = prefix

        remaining = window[
            match.end():
        ]

        for char in remaining:

            if char.isalnum():
                raw += char.upper()

            elif char in (
                " ",
                ".",
                "-",
                "_",
                "/",
                "(",
                ")",
                ":",
                ";",
                "\u00A0",
            ):
                continue

            else:
                break

            if len(raw) > 16:
                break

            candidate = compact_identifier(
                raw
            )

            if is_valid_tax_identifier(
                candidate
            ):
                return candidate

    return None


def _find_fallback_vat_ids(
    text: str,
) -> Iterable[str]:
    """
    Full document EU/UK VAT scan.

    Crucially, the country prefix is mandatory.
    """
    country_pattern = "|".join(
        sorted(
            EU_VAT_BODY_LENGTHS,
            key=len,
            reverse=True,
        )
    )

    country_re = re.compile(
        rf"\b({country_pattern})\b",
        re.IGNORECASE,
    )

    for match in country_re.finditer(
        text
    ):
        candidate = _scan_vat_after_position(
            text,
            match.start(),
        )

        if candidate:
            yield candidate


def _find_fallback_nine_digit_ids(
    text: str,
) -> Iterable[str]:
    """
    Fallback ONLY for exactly 9 contiguous digits.

    This prevents:
        123456789 987654321
    becoming:
        123456789987654321
    """
    for match in re.finditer(
        r"\b\d{9}\b",
        text,
    ):
        yield match.group(0)


def extract_tax_id(
    text: str,
) -> Optional[str]:
    """
    Extract Tax ID / VAT.

    Priority:

    1. Explicit Greek AFM.
    2. Explicit US EIN.
    3. Explicit EU/UK VAT anchor.
    4. Full-document EU/UK VAT.
    5. Standalone 9-digit fallback.

    Email addresses and URLs are stripped FIRST.
    """
    if not text:
        return None

    text = normalize_text(text)

    text = strip_emails_and_urls(
        text
    )

    if not text.strip():
        return None

    # ------------------------------------------------------------
    # 1. Greek AFM
    # ------------------------------------------------------------
    match = GREEK_AFM_ANCHORED_RE.search(
        text
    )

    if match:
        afm = match.group(1)

        if re.fullmatch(
            r"\d{9}",
            afm,
        ):
            return afm

    # ------------------------------------------------------------
    # 2. US EIN, canonical XX-XXXXXXX
    # ------------------------------------------------------------
    match = EIN_ANCHORED_RE.search(
        text
    )

    if match:
        return (
            f"{match.group(1)}-"
            f"{match.group(2)}"
        )

    # EIN written as 9 contiguous digits.
    match = EIN_PLAIN_RE.search(
        text
    )

    if match:
        raw = match.group(1)

        return (
            f"{raw[:2]}-"
            f"{raw[2:]}"
        )

    # ------------------------------------------------------------
    # 3. Explicit generic Tax/VAT anchor
    # ------------------------------------------------------------
    for anchor in TAX_ANCHOR_RE.finditer(
        text
    ):
        candidate = _scan_vat_after_position(
            text,
            anchor.end(),
        )

        if candidate:
            return candidate

        # Generic anchored bare tax IDs are still STRICTLY 9 digits.
        window = text[
            anchor.end():
            anchor.end() + 24
        ]

        bare_id = re.search(
            r"\b\d{9}\b",
            window,
        )

        if bare_id:
            return bare_id.group(0)

    # ------------------------------------------------------------
    # 4. EU/UK VAT fallback
    # ------------------------------------------------------------
    for candidate in _find_fallback_vat_ids(
        text
    ):
        return candidate

    # ------------------------------------------------------------
    # 5. Bare 9-digit fallback
    # ------------------------------------------------------------
    for candidate in _find_fallback_nine_digit_ids(
        text
    ):
        return candidate

    return None


# ============================================================================
# DATE EXTRACTION
# ============================================================================

MONTHS = {
    # English
    "january": 1,
    "jan": 1,
    "february": 2,
    "feb": 2,
    "march": 3,
    "mar": 3,
    "april": 4,
    "apr": 4,
    "may": 5,
    "june": 6,
    "jun": 6,
    "july": 7,
    "jul": 7,
    "august": 8,
    "aug": 8,
    "september": 9,
    "sep": 9,
    "sept": 9,
    "october": 10,
    "oct": 10,
    "november": 11,
    "nov": 11,
    "december": 12,
    "dec": 12,

    # German
    "januar": 1,
    "februar": 2,
    "marz": 3,
    "maerz": 3,
    "juni": 6,
    "juli": 7,
    "oktober": 10,
    "dezember": 12,

    # French
    "janvier": 1,
    "fevrier": 2,
    "mars": 3,
    "avril": 4,
    "mai": 5,
    "juin": 6,
    "juillet": 7,
    "aout": 8,
    "septembre": 9,
    "octobre": 10,
    "novembre": 11,
    "decembre": 12,

    # Spanish
    "enero": 1,
    "febrero": 2,
    "marzo": 3,
    "abril": 4,
    "mayo": 5,
    "junio": 6,
    "julio": 7,
    "agosto": 8,
    "septiembre": 9,
    "setiembre": 9,
    "octubre": 10,
    "noviembre": 11,
    "diciembre": 12,

    # Greek full names and abbreviations.
    "ιανουαριος": 1,
    "ιανουαριου": 1,
    "ιαν": 1,

    "φεβρουαριος": 2,
    "φεβρουαριου": 2,
    "φεβ": 2,

    "μαρτιος": 3,
    "μαρτιου": 3,
    "μαρ": 3,

    "απριλιος": 4,
    "απριλιου": 4,
    "απρ": 4,

    "μαιος": 5,
    "μαιου": 5,
    "μαι": 5,

    "ιουνιος": 6,
    "ιουνιου": 6,
    "ιουν": 6,

    "ιουλιος": 7,
    "ιουλιου": 7,
    "ιουλ": 7,

    "αυγουστος": 8,
    "αυγουστου": 8,
    "αυγ": 8,

    "σεπτεμβριος": 9,
    "σεπτεμβριου": 9,
    "σεπ": 9,

    "οκτωβριος": 10,
    "οκτωβριου": 10,
    "οκτ": 10,

    "νοεμβριος": 11,
    "νοεμβριου": 11,
    "νοε": 11,

    "δεκεμβριος": 12,
    "δεκεμβριου": 12,
    "δεκ": 12,
}


def _normalize_month_name(
    value: str,
) -> str:
    value = unicodedata.normalize(
        "NFKD",
        value or "",
    )

    value = "".join(
        ch
        for ch in value
        if not unicodedata.combining(ch)
    )

    return (
        value
        .casefold()
        .replace(".", "")
    )


def _lookup_month(
    value: str,
) -> Optional[int]:
    return MONTHS.get(
        _normalize_month_name(value)
    )


def _make_date(
    year: int,
    month: int,
    day: int,
) -> Optional[str]:

    if year < 100:
        year += (
            2000
            if year < 70
            else 1900
        )

    try:
        return dt.date(
            year,
            month,
            day,
        ).isoformat()

    except ValueError:
        return None


def parse_date(
    text: str,
) -> Optional[str]:
    """
    Supported:

        YYYY-MM-DD
        DD/MM/YYYY
        DD.MM.YYYY
        DD-MM-YY
        MM/DD/YYYY

        15 Ιανουαρίου 2026
        15 Ιαν 2026

        15 January 2026
        January 15, 2026
    """
    if not text:
        return None

    text = normalize_text(
        text
    )

    # ------------------------------------------------------------
    # ISO
    # ------------------------------------------------------------
    iso_re = re.compile(
        r"\b"
        r"(\d{4})-(\d{1,2})-(\d{1,2})"
        r"\b"
    )

    for match in iso_re.finditer(
        text
    ):
        result = _make_date(
            int(match.group(1)),
            int(match.group(2)),
            int(match.group(3)),
        )

        if result:
            return result

    # ------------------------------------------------------------
    # Numeric dates
    # ------------------------------------------------------------
    numeric_re = re.compile(
        r"\b"
        r"(\d{1,2})"
        r"([./-])"
        r"(\d{1,2})"
        r"\2"
        r"(\d{2,4})"
        r"\b"
    )

    for match in numeric_re.finditer(
        text
    ):
        first = int(
            match.group(1)
        )

        second = int(
            match.group(3)
        )

        year = int(
            match.group(4)
        )

        # Unambiguous US:
        # 08/25/2026
        if (
            first <= 12
            and second > 12
        ):
            month = first
            day = second

        # Unambiguous European:
        # 25/08/2026
        elif (
            first > 12
            and second <= 12
        ):
            day = first
            month = second

        # Ambiguous:
        # 03/04/2026
        #
        # Invoice-oriented European default.
        else:
            day = first
            month = second

        result = _make_date(
            year,
            month,
            day,
        )

        if result:
            return result

    # ------------------------------------------------------------
    # DD Month YYYY
    # ------------------------------------------------------------
    textual_re = re.compile(
        r"\b"
        r"(\d{1,2})"
        r"\s+"
        r"([A-Za-zÀ-ÿ"
        r"Α-Ωα-ω"
        r"άέήίόύώ"
        r"ϊΐϋΰ"
        r"äöüßÄÖÜ.]+)"
        r"\s+"
        r"(\d{2,4})"
        r"\b",
        re.IGNORECASE,
    )

    for match in textual_re.finditer(
        text
    ):
        day = int(
            match.group(1)
        )

        month = _lookup_month(
            match.group(2)
        )

        year = int(
            match.group(3)
        )

        if month is None:
            continue

        result = _make_date(
            year,
            month,
            day,
        )

        if result:
            return result

    # ------------------------------------------------------------
    # Month DD YYYY
    # ------------------------------------------------------------
    month_first_re = re.compile(
        r"\b"
        r"([A-Za-zÀ-ÿ"
        r"Α-Ωα-ω"
        r"άέήίόύώ"
        r"ϊΐϋΰ"
        r"äöüßÄÖÜ.]+)"
        r"\s+"
        r"(\d{1,2})"
        r",?\s+"
        r"(\d{2,4})"
        r"\b",
        re.IGNORECASE,
    )

    for match in month_first_re.finditer(
        text
    ):
        month = _lookup_month(
            match.group(1)
        )

        if month is None:
            continue

        result = _make_date(
            int(match.group(3)),
            month,
            int(match.group(2)),
        )

        if result:
            return result

    return None


# ============================================================================
# AMOUNT EXTRACTION
# ============================================================================

TOTAL_KEYWORDS = (
    "TOTAL DUE",
    "TOTAL AMOUNT",
    "GRAND TOTAL",
    "AMOUNT DUE",
    "TOTAL",

    "ΣΥΝΟΛΟ",
    "ΠΛΗΡΩΤΕΟ",
    "ΤΕΛΙΚΟ ΠΟΣΟ",

    "GESAMTBETRAG",
    "ZU ZAHLEN",
    "BRUTTOBETRAG",

    "MONTANT TTC",
    "TOTAL TTC",
)

NET_KEYWORDS = (
    "SUBTOTAL",
    "NET AMOUNT",
    "NET TOTAL",
    "NETTOBETRAG",
    "NETTO BETRAG",
    "NETTO",

    "ΚΑΘΑΡΗ ΑΞΙΑ",
    "ΚΑΘΑΡΗ",

    "MONTANT HT",
    "HORS TAXES",
)

VAT_KEYWORDS = (
    "VAT AMOUNT",
    "VAT",
    "MWST",
    "UST",
    "UMSATZSTEUER",

    "Φ.Π.Α.",
    "ΦΠΑ",

    "TVA",
    "TAXE SUR LA VALEUR AJOUTEE",
)

CURRENCY_RE = (
    r"(?:"
    r"€|EUR|"
    r"USD|US\$|\$|"
    r"£|GBP|"
    r"CHF|"
    r"CNY|¥|"
    r"円|"
    r"R\$"
    r")"
)

PERCENT_RE = re.compile(
    r"(?<!\d)"
    r"(\d{1,3}(?:[.,]\d+)?)"
    r"\s*%",
)

CURRENCY_AMOUNT_RE = re.compile(
    rf"""
    (?:
        (?P<prefix>{CURRENCY_RE})
        \s*
        (?P<prefix_amount>
            [-+]?
            (?:\d{{1,3}}(?:[ .,\u00A0]\d{{3}})+|\d+)
            (?:[.,]\d{{1,4}})?
        )
        |
        (?P<suffix_amount>
            [-+]?
            (?:\d{{1,3}}(?:[ .,\u00A0]\d{{3}})+|\d+)
            (?:[.,]\d{{1,4}})?
        )
        \s*
        (?P<suffix>{CURRENCY_RE})
    )
    """,
    re.IGNORECASE | re.VERBOSE,
)

GENERIC_AMOUNT_RE = re.compile(
    r"""
    (?<!\d)
    (
        [-+]?
        (?:\d{1,3}(?:[ .,\u00A0]\d{3})+|\d+)
        (?:[.,]\d{1,4})?
    )
    (?!\d)
    """,
    re.VERBOSE,
)


def _parse_numeric_amount(
    token: str,
) -> Optional[float]:
    """Parse European and US-style monetary numbers."""
    if token is None:
        return None

    value = str(token).strip()

    if not value:
        return None

    value = (
        value
        .replace("\u00A0", " ")
        .replace(" ", "")
        .replace("'", "")
    )

    negative_parentheses = (
        value.startswith("(")
        and value.endswith(")")
    )

    if negative_parentheses:
        value = value[1:-1]

    value = re.sub(
        r"[^0-9,.+\-]",
        "",
        value,
    )

    if not value:
        return None

    # 1.234,56
    # 1,234.56
    if "," in value and "." in value:

        if value.rfind(",") > value.rfind("."):
            value = (
                value
                .replace(".", "")
                .replace(",", ".")
            )
        else:
            value = (
                value
                .replace(",", "")
            )

    # Single separator.
    elif "," in value or "." in value:

        separator = (
            ","
            if "," in value
            else "."
        )

        pieces = value.split(
            separator
        )

        if len(pieces) > 2:

            value = (
                "".join(
                    pieces[:-1]
                )
                + "."
                + pieces[-1]
            )

        else:

            left, right = pieces

            # 1.234 => 1234
            # 1,234 => 1234
            if (
                len(right) == 3
                and len(left) <= 3
            ):
                value = (
                    left
                    + right
                )
            else:
                value = (
                    left
                    + "."
                    + right
                )

    try:
        result = float(
            value
        )

    except (
        ValueError,
        TypeError,
    ):
        return None

    if negative_parentheses:
        result = -result

    if not math.isfinite(
        result
    ):
        return None

    return result


def _line_amounts(
    line: str,
) -> list[float]:
    """
    Extract all monetary values from a line.

    Currency-aware first, including trailing currency:
        1.000,00 €
    """
    if not line:
        return []

    safe_line = PERCENT_RE.sub(
        " ",
        line,
    )

    values: list[float] = []

    currency_matches = list(
        CURRENCY_AMOUNT_RE.finditer(
            safe_line
        )
    )

    if currency_matches:

        for match in currency_matches:

            token = (
                match.group(
                    "prefix_amount"
                )
                or match.group(
                    "suffix_amount"
                )
            )

            if not token:
                continue

            amount = _parse_numeric_amount(
                token
            )

            if (
                amount is not None
                and amount >= 0
            ):
                values.append(
                    amount
                )

        return values

    for match in GENERIC_AMOUNT_RE.finditer(
        safe_line
    ):
        token = match.group(1)

        raw_digits = re.sub(
            r"\D",
            "",
            token,
        )

        # Prevent IDs / phone numbers / barcodes
        # from becoming amounts.
        if (
            len(raw_digits) >= 9
            and token.isdigit()
        ):
            continue

        amount = _parse_numeric_amount(
            token
        )

        if (
            amount is not None
            and amount >= 0
        ):
            values.append(
                amount
            )

    return values


def extract_amounts(
    text: str,
) -> list[float]:

    values: list[float] = []

    for line in normalize_text(
        text
    ).splitlines():

        values.extend(
            _line_amounts(line)
        )

    return values


def _line_matches_keywords(
    line: str,
    keywords: tuple[str, ...],
) -> bool:

    folded_line = fold_text(
        line
    )

    return any(
        fold_text(keyword)
        in folded_line
        for keyword in keywords
    )


def _keyword_line_amounts(
    text: str,
    keywords: tuple[str, ...],
) -> list[float]:

    values: list[float] = []

    for line in normalize_text(
        text
    ).splitlines():

        if not _line_matches_keywords(
            line,
            keywords,
        ):
            continue

        amounts = _line_amounts(
            line
        )

        if amounts:
            values.append(
                amounts[-1]
            )

    return values


def extract_tax_rate(
    text: str,
) -> Optional[float]:

    for line in normalize_text(
        text
    ).splitlines():

        if not _line_matches_keywords(
            line,
            VAT_KEYWORDS,
        ):
            continue

        match = PERCENT_RE.search(
            line
        )

        if not match:
            continue

        rate = _parse_numeric_amount(
            match.group(1)
        )

        if (
            rate is not None
            and 0 < rate <= 100
        ):
            return rate

    return None


def extract_amount_fields(
    text: str,
    warnings: list[str],
) -> tuple[
    Optional[float],
    Optional[float],
    Optional[float],
]:
    """
    Extract Net / VAT / Total.

    Required rules:

    - Total + VAT missing/zero:
        Net = Total
        VAT = 0.00

    - Total + VAT + Net:
        preserve explicit values and warn on mismatch.

    - Total + VAT + missing Net:
        Net = Total - VAT

    - Net + VAT + missing Total:
        Total = Net + VAT

    - Missing Total:
        largest invoice-like amount is used as fallback.
    """
    normalized = normalize_text(
        text
    )

    if not normalized:
        return (
            None,
            None,
            None,
        )

    net_values = _keyword_line_amounts(
        normalized,
        NET_KEYWORDS,
    )

    vat_values = _keyword_line_amounts(
        normalized,
        VAT_KEYWORDS,
    )

    total_values = _keyword_line_amounts(
        normalized,
        TOTAL_KEYWORDS,
    )

    net = (
        net_values[-1]
        if net_values
        else None
    )

    vat = (
        vat_values[-1]
        if vat_values
        else None
    )

    total = (
        total_values[-1]
        if total_values
        else None
    )

    # ------------------------------------------------------------
    # Total fallback.
    # ------------------------------------------------------------
    if total is None:

        amounts = extract_amounts(
            normalized
        )

        if amounts:
            total = max(
                amounts
            )

            warnings.append(
                "Total keyword not found; "
                "largest invoice-like amount used as Total."
            )

    # ------------------------------------------------------------
    # Required VAT=0/missing behavior.
    # ------------------------------------------------------------
    if (
        total is not None
        and (
            vat is None
            or abs(vat) < 0.005
        )
    ):
        net = round(
            total,
            2,
        )

        vat = 0.00

    # ------------------------------------------------------------
    # Total + VAT but missing Net.
    # ------------------------------------------------------------
    elif (
        total is not None
        and vat is not None
        and net is None
    ):
        derived_net = (
            total - vat
        )

        if derived_net >= -0.005:
            net = round(
                max(
                    derived_net,
                    0.0,
                ),
                2,
            )

            warnings.append(
                "Net Amount inferred as Total - VAT."
            )
        else:
            warnings.append(
                "Could not derive a valid Net Amount from Total - VAT."
            )

    # ------------------------------------------------------------
    # Total + Net + VAT consistency.
    # ------------------------------------------------------------
    elif (
        total is not None
        and net is not None
        and vat is not None
    ):
        difference = abs(
            (net + vat)
            - total
        )

        if difference > 0.05:
            warnings.append(
                "Net + VAT does not reconcile with Total within 0.05."
            )

    # ------------------------------------------------------------
    # Net + VAT, missing Total.
    # ------------------------------------------------------------
    if (
        total is None
        and net is not None
        and vat is not None
    ):
        total = round(
            net + vat,
            2,
        )

        warnings.append(
            "Total Amount inferred as Net + VAT."
        )

    # ------------------------------------------------------------
    # Net + known VAT rate, missing VAT/Total.
    # ------------------------------------------------------------
    if (
        total is None
        and net is not None
        and vat is None
    ):
        rate = extract_tax_rate(
            normalized
        )

        if rate is not None:
            vat = round(
                net * rate / 100.0,
                2,
            )

            total = round(
                net + vat,
                2,
            )

            warnings.append(
                f"VAT inferred from {rate:g}% tax rate."
            )

    # ------------------------------------------------------------
    # Final required invariant.
    # ------------------------------------------------------------
    if total is not None:

        total = round(
            float(total),
            2,
        )

        if (
            vat is None
            or abs(vat) < 0.005
        ):
            vat = 0.00
            net = total

        elif net is None:
            net = round(
                total - vat,
                2,
            )

    if net is not None:
        net = round(
            float(net),
            2,
        )

    if vat is not None:
        vat = round(
            float(vat),
            2,
        )

    return (
        net,
        vat,
        total,
    )


# ============================================================================
# PARSE RESULT
# ============================================================================

@dataclass
class ParseResult:
    file_name: str

    invoice_date: Optional[str] = None
    tax_id: Optional[str] = None

    net_amount: Optional[float] = None
    vat_amount: Optional[float] = None
    total_amount: Optional[float] = None

    warnings: list[str] = field(
        default_factory=list
    )

    fatal_error: Optional[str] = None

    @property
    def status(self) -> str:
        if self.fatal_error:
            return "Warning"

        return (
            "Success"
            if not self.warnings
            else "Warning"
        )

    def to_row(self) -> dict:
        return {
            "Invoice Date": (
                self.invoice_date
                or ""
            ),
            "Tax ID / VAT": (
                self.tax_id
                or ""
            ),
            "Net Amount (€/$)": (
                self.net_amount
            ),
            "VAT Amount (€/$)": (
                self.vat_amount
            ),
            "Total Amount (€/$)": (
                self.total_amount
            ),
            "File Name": self.file_name,
            "Status": self.status,
        }


# ============================================================================
# PDF EXTRACTION
# ============================================================================

def extract_pdf_text(
    pdf_path: Path,
    warnings: list[str],
) -> str:

    pages: list[str] = []

    try:

        with pdfplumber.open(
            str(pdf_path)
        ) as pdf:

            if not pdf.pages:
                warnings.append(
                    "PDF contains no pages."
                )

                return ""

            for page_number, page in enumerate(
                pdf.pages,
                start=1,
            ):

                try:

                    text = (
                        page.extract_text(
                            x_tolerance=1.5,
                            y_tolerance=3.0,
                        )
                        or ""
                    )

                    if not text.strip():

                        text = (
                            page.extract_text(
                                layout=True
                            )
                            or ""
                        )

                    if text.strip():
                        pages.append(
                            text
                        )

                except Exception as exc:
                    warnings.append(
                        (
                            f"Page {page_number} "
                            f"could not be extracted "
                            f"({exc.__class__.__name__})."
                        )
                    )

    except Exception as exc:

        raise RuntimeError(
            f"PDF could not be opened/read: {exc}"
        ) from exc

    full_text = "\n\n".join(
        pages
    )

    if not full_text.strip():

        warnings.append(
            "No text layer was extracted; "
            "scanned/image-only PDF requires OCR."
        )

    return full_text


def parse_invoice(
    pdf_path: Path,
) -> ParseResult:

    result = ParseResult(
        file_name=pdf_path.name
    )

    try:

        text = extract_pdf_text(
            pdf_path,
            result.warnings,
        )

        if not text.strip():

            result.warnings.append(
                "No readable text was available for parsing."
            )

            return result

        # Date.
        result.invoice_date = parse_date(
            text
        )

        if not result.invoice_date:
            result.warnings.append(
                "Invoice Date not found."
            )

        # Tax ID.
        result.tax_id = extract_tax_id(
            text
        )

        if not result.tax_id:
            result.warnings.append(
                "Tax ID / VAT not found."
            )

        # Amounts.
        (
            result.net_amount,
            result.vat_amount,
            result.total_amount,
        ) = extract_amount_fields(
            text,
            result.warnings,
        )

        if result.net_amount is None:
            result.warnings.append(
                "Net Amount not found."
            )

        if result.vat_amount is None:
            result.warnings.append(
                "VAT Amount not found."
            )

        if result.total_amount is None:
            result.warnings.append(
                "Total Amount not found."
            )

    except Exception as exc:

        result.fatal_error = (
            f"{exc.__class__.__name__}: "
            f"{exc}"
        )

        result.warnings.append(
            result.fatal_error
        )

    return result


# ============================================================================
# EXCEL EXPORT
# ============================================================================

def export_to_excel(
    results: list[ParseResult],
    requested_path: Path,
) -> Path:

    output_path = (
        requested_path
        .expanduser()
        .resolve()
    )

    if output_path.suffix.lower() != ".xlsx":
        output_path = (
            output_path.with_suffix(
                ".xlsx"
            )
        )

    output_path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    dataframe = pd.DataFrame(
        [
            result.to_row()
            for result in results
        ],
        columns=OUTPUT_COLUMNS,
    )

    temp_path: Optional[Path] = None

    try:

        with tempfile.NamedTemporaryFile(
            prefix=".invoice2excel_",
            suffix=".xlsx",
            dir=str(output_path.parent),
            delete=False,
        ) as temp_file:

            temp_path = Path(
                temp_file.name
            )

        dataframe.to_excel(
            temp_path,
            index=False,
            sheet_name="Invoices",
            engine="openpyxl",
        )

        workbook = load_workbook(
            temp_path
        )

        worksheet = workbook[
            "Invoices"
        ]

        header_fill = PatternFill(
            "solid",
            fgColor="1F2937",
        )

        header_font = Font(
            color="FFFFFF",
            bold=True,
        )

        success_fill = PatternFill(
            "solid",
            fgColor="DCFCE7",
        )

        warning_fill = PatternFill(
            "solid",
            fgColor="FEF3C7",
        )

        thin = Side(
            style="thin",
            color="D1D5DB",
        )

        border = Border(
            left=thin,
            right=thin,
            top=thin,
            bottom=thin,
        )

        # Header.
        for cell in worksheet[1]:

            cell.fill = header_fill
            cell.font = header_font
            cell.border = border

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
            )

        columns = {
            cell.value: cell.column
            for cell in worksheet[1]
        }

        # Body.
        for row in worksheet.iter_rows(
            min_row=2
        ):

            for cell in row:

                cell.border = border

                cell.alignment = Alignment(
                    vertical="center"
                )

        # Numbers.
        for name in (
            "Net Amount (€/$)",
            "VAT Amount (€/$)",
            "Total Amount (€/$)",
        ):

            column_index = columns[
                name
            ]

            for row_index in range(
                2,
                worksheet.max_row + 1,
            ):

                worksheet.cell(
                    row=row_index,
                    column=column_index,
                ).number_format = (
                    "#,##0.00"
                )

        # Status.
        status_column = columns[
            "Status"
        ]

        for row_index in range(
            2,
            worksheet.max_row + 1,
        ):

            cell = worksheet.cell(
                row=row_index,
                column=status_column,
            )

            cell.fill = (
                success_fill
                if cell.value == "Success"
                else warning_fill
            )

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
            )

        worksheet.freeze_panes = "A2"

        worksheet.auto_filter.ref = (
            worksheet.dimensions
        )

        worksheet.sheet_view.showGridLines = (
            False
        )

        worksheet.row_dimensions[
            1
        ].height = 24

        # Auto-fit widths.
        for column_index in range(
            1,
            worksheet.max_column + 1,
        ):

            max_length = 0

            for row in worksheet.iter_rows(
                min_col=column_index,
                max_col=column_index,
                min_row=1,
                max_row=worksheet.max_row,
            ):

                for cell in row:

                    value = (
                        ""
                        if cell.value is None
                        else str(cell.value)
                    )

                    max_length = max(
                        max_length,
                        len(value),
                    )

            worksheet.column_dimensions[
                get_column_letter(
                    column_index
                )
            ].width = min(
                max(
                    max_length + 2,
                    12,
                ),
                42,
            )

        workbook.save(
            temp_path
        )

        try:

            # Atomic replacement where supported.
            os.replace(
                temp_path,
                output_path,
            )

            temp_path = None

            return output_path

        except PermissionError as exc:

            fallback_path = (
                output_path.with_name(
                    f"{output_path.stem}_"
                    f"{dt.datetime.now():%Y%m%d_%H%M%S}"
                    f".xlsx"
                )
            )

            os.replace(
                temp_path,
                fallback_path,
            )

            temp_path = None

            raise RuntimeError(
                "Output workbook is locked/open. "
                f"Export was saved to: {fallback_path}"
            ) from exc

    finally:

        if temp_path is not None:

            try:
                temp_path.unlink(
                    missing_ok=True
                )
            except OSError:
                pass


# ============================================================================
# GUI
# ============================================================================

class Invoice2ExcelApp(ctk.CTk):

    def __init__(self) -> None:

        super().__init__()

        self.title(
            f"{APP_NAME} {APP_VERSION}"
        )

        self.geometry(
            "980x720"
        )

        self.minsize(
            860,
            620,
        )

        self.protocol(
            "WM_DELETE_WINDOW",
            self._on_close,
        )

        ctk.set_appearance_mode(
            "dark"
        )

        ctk.set_default_color_theme(
            "blue"
        )

        self.message_queue: queue.Queue = (
            queue.Queue()
        )

        self.processing = False

        self.worker: Optional[
            threading.Thread
        ] = None

        self.input_var = (
            ctk.StringVar()
        )

        self.output_var = ctk.StringVar(
            value=str(
                Path.home()
                / "invoice_export.xlsx"
            )
        )

        self.progress_var = (
            ctk.DoubleVar(
                value=0.0
            )
        )

        self.status_var = (
            ctk.StringVar(
                value="Ready."
            )
        )

        self._build_ui()

        self.after(
            100,
            self._poll_queue,
        )

    def _build_ui(self) -> None:

        self.grid_columnconfigure(
            0,
            weight=1,
        )

        self.grid_rowconfigure(
            5,
            weight=1,
        )

        title = ctk.CTkLabel(
            self,
            text=APP_NAME,
            font=ctk.CTkFont(
                size=28,
                weight="bold",
            ),
        )

        title.grid(
            row=0,
            column=0,
            padx=28,
            pady=(24, 4),
            sticky="w",
        )

        subtitle = ctk.CTkLabel(
            self,
            text=(
                "Local PDF invoice/receipt "
                "parser → polished XLSX export"
            ),
            text_color=(
                "gray60",
                "gray70",
            ),
            font=ctk.CTkFont(
                size=13
            ),
        )

        subtitle.grid(
            row=1,
            column=0,
            padx=30,
            pady=(0, 18),
            sticky="w",
        )

        controls = ctk.CTkFrame(
            self,
            corner_radius=14,
        )

        controls.grid(
            row=2,
            column=0,
            padx=24,
            pady=(0, 14),
            sticky="ew",
        )

        controls.grid_columnconfigure(
            1,
            weight=1,
        )

        # Input folder.
        ctk.CTkLabel(
            controls,
            text="Input PDF Folder",
        ).grid(
            row=0,
            column=0,
            padx=(18, 10),
            pady=(18, 10),
            sticky="w",
        )

        self.input_entry = ctk.CTkEntry(
            controls,
            textvariable=self.input_var,
        )

        self.input_entry.grid(
            row=0,
            column=1,
            padx=10,
            pady=(18, 10),
            sticky="ew",
        )

        self.input_button = ctk.CTkButton(
            controls,
            text="Browse…",
            width=110,
            command=self._browse_input,
        )

        self.input_button.grid(
            row=0,
            column=2,
            padx=(10, 18),
            pady=(18, 10),
        )

        # Output.
        ctk.CTkLabel(
            controls,
            text="Output XLSX",
        ).grid(
            row=1,
            column=0,
            padx=(18, 10),
            pady=10,
            sticky="w",
        )

        self.output_entry = ctk.CTkEntry(
            controls,
            textvariable=self.output_var,
        )

        self.output_entry.grid(
            row=1,
            column=1,
            padx=10,
            pady=10,
            sticky="ew",
        )

        self.output_button = ctk.CTkButton(
            controls,
            text="Save As…",
            width=110,
            command=self._browse_output,
        )

        self.output_button.grid(
            row=1,
            column=2,
            padx=(10, 18),
            pady=10,
        )

        # Start.
        self.start_button = ctk.CTkButton(
            controls,
            text="Start Processing",
            height=42,
            font=ctk.CTkFont(
                size=15,
                weight="bold",
            ),
            command=self._start_processing,
        )

        self.start_button.grid(
            row=2,
            column=0,
            columnspan=3,
            padx=18,
            pady=(10, 18),
            sticky="ew",
        )

        # Progress.
        self.progress = ctk.CTkProgressBar(
            self,
            variable=self.progress_var,
            height=12,
        )

        self.progress.grid(
            row=3,
            column=0,
            padx=28,
            pady=(4, 8),
            sticky="ew",
        )

        # Status.
        status = ctk.CTkLabel(
            self,
            textvariable=self.status_var,
            anchor="w",
            font=ctk.CTkFont(
                size=13,
                weight="bold",
            ),
        )

        status.grid(
            row=4,
            column=0,
            padx=28,
            pady=(2, 8),
            sticky="ew",
        )

        # Log.
        self.log_box = ctk.CTkTextbox(
            self,
            wrap="word",
            corner_radius=12,
        )

        self.log_box.grid(
            row=5,
            column=0,
            padx=24,
            pady=(0, 24),
            sticky="nsew",
        )

        self.log_box.configure(
            state="disabled"
        )

        self._append_log(
            (
                f"{APP_NAME} "
                f"{APP_VERSION} ready."
            )
        )

    def _append_log(
        self,
        message: str,
    ) -> None:

        self.log_box.configure(
            state="normal"
        )

        self.log_box.insert(
            "end",
            str(message).rstrip()
            + "\n",
        )

        self.log_box.see(
            "end"
        )

        self.log_box.configure(
            state="disabled"
        )

    def _browse_input(
        self,
    ) -> None:

        from tkinter import filedialog

        folder = (
            filedialog.askdirectory(
                title="Select PDF input folder"
            )
        )

        if folder:
            self.input_var.set(
                folder
            )

    def _browse_output(
        self,
    ) -> None:

        from tkinter import filedialog

        path = (
            filedialog.asksaveasfilename(
                title="Choose Excel output",
                defaultextension=".xlsx",
                filetypes=[
                    (
                        "Excel workbook",
                        "*.xlsx",
                    )
                ],
            )
        )

        if path:
            self.output_var.set(
                path
            )

    def _set_controls_enabled(
        self,
        enabled: bool,
    ) -> None:

        state = (
            "normal"
            if enabled
            else "disabled"
        )

        for widget in (
            self.start_button,
            self.input_button,
            self.output_button,
            self.input_entry,
            self.output_entry,
        ):
            widget.configure(
                state=state
            )

    def _start_processing(
        self,
    ) -> None:

        if self.processing:
            return

        input_text = (
            self.input_var
            .get()
            .strip()
        )

        output_text = (
            self.output_var
            .get()
            .strip()
        )

        if not input_text:
            self._append_log(
                "[ERROR] Select an input PDF folder."
            )

            self.status_var.set(
                "Input folder required."
            )

            return

        if not output_text:
            self._append_log(
                "[ERROR] Select an output XLSX path."
            )

            self.status_var.set(
                "Output path required."
            )

            return

        input_dir = (
            Path(input_text)
            .expanduser()
        )

        output_path = (
            Path(output_text)
            .expanduser()
        )

        if not input_dir.is_dir():

            self._append_log(
                "[ERROR] Input folder does not exist."
            )

            self.status_var.set(
                "Invalid input folder."
            )

            return

        if output_path.suffix.lower() != ".xlsx":

            output_path = (
                output_path.with_suffix(
                    ".xlsx"
                )
            )

            self.output_var.set(
                str(output_path)
            )

        try:

            pdf_files = sorted(
                [
                    path
                    for path
                    in input_dir.iterdir()
                    if (
                        path.is_file()
                        and
                        path.suffix.lower()
                        in SUPPORTED_EXTENSIONS
                    )
                ],
                key=lambda path: (
                    path.name.casefold()
                ),
            )

        except OSError as exc:

            self._append_log(
                (
                    "[ERROR] Cannot enumerate "
                    f"input folder: {exc}"
                )
            )

            self.status_var.set(
                "Cannot read input folder."
            )

            return

        if not pdf_files:

            self._append_log(
                "[ERROR] No PDF files found."
            )

            self.status_var.set(
                "No PDFs found."
            )

            return

        self.processing = True

        self._set_controls_enabled(
            False
        )

        self.progress_var.set(
            0.0
        )

        self.status_var.set(
            (
                f"Starting… 0 of "
                f"{len(pdf_files)} files."
            )
        )

        self._append_log(
            (
                f"Found {len(pdf_files)} "
                "PDF file(s)."
            )
        )

        self.worker = threading.Thread(
            target=self._worker_process,
            args=(
                pdf_files,
                output_path,
            ),
            daemon=True,
            name="Invoice2ExcelWorker",
        )

        self.worker.start()

    def _worker_process(
        self,
        pdf_files: list[Path],
        output_path: Path,
    ) -> None:

        results: list[
            ParseResult
        ] = []

        total = len(
            pdf_files
        )

        for index, pdf_path in enumerate(
            pdf_files,
            start=1,
        ):

            self.message_queue.put(
                (
                    "status",
                    (
                        f"Processing file "
                        f"{index} of {total}: "
                        f"{pdf_path.name}"
                    ),
                )
            )

            try:

                result = parse_invoice(
                    pdf_path
                )

            except Exception as exc:

                result = ParseResult(
                    file_name=pdf_path.name,
                    warnings=[
                        (
                            "Unexpected worker error: "
                            f"{exc.__class__.__name__}: "
                            f"{exc}"
                        )
                    ],
                    fatal_error=str(exc),
                )

            results.append(
                result
            )

            if result.status == "Success":

                self.message_queue.put(
                    (
                        "log",
                        (
                            f"[OK] "
                            f"{pdf_path.name}"
                        ),
                    )
                )

            else:

                details = (
                    " | ".join(
                        result.warnings
                    )
                    or "No details."
                )

                self.message_queue.put(
                    (
                        "log",
                        (
                            f"[WARN] "
                            f"{pdf_path.name}: "
                            f"{details}"
                        ),
                    )
                )

            self.message_queue.put(
                (
                    "progress",
                    index / total,
                )
            )

        try:

            self.message_queue.put(
                (
                    "status",
                    "Writing Excel workbook…",
                )
            )

            output = export_to_excel(
                results,
                output_path,
            )

            self.message_queue.put(
                (
                    "log",
                    (
                        "[DONE] Excel export created: "
                        f"{output}"
                    ),
                )
            )

            self.message_queue.put(
                (
                    "status",
                    (
                        f"Finished. Processed "
                        f"{total} file(s)."
                    ),
                )
            )

        except Exception as exc:

            self.message_queue.put(
                (
                    "log",
                    (
                        "[ERROR] Excel export failed: "
                        f"{exc.__class__.__name__}: "
                        f"{exc}"
                    ),
                )
            )

            self.message_queue.put(
                (
                    "status",
                    (
                        "Finished with "
                        "export errors."
                    ),
                )
            )

        finally:

            self.message_queue.put(
                (
                    "finished",
                    None,
                )
            )

    def _poll_queue(
        self,
    ) -> None:

        try:

            while True:

                kind, payload = (
                    self.message_queue
                    .get_nowait()
                )

                if kind == "log":

                    self._append_log(
                        payload
                    )

                elif kind == "status":

                    self.status_var.set(
                        payload
                    )

                elif kind == "progress":

                    self.progress_var.set(
                        max(
                            0.0,
                            min(
                                1.0,
                                float(payload),
                            ),
                        )
                    )

                elif kind == "finished":

                    self.processing = False

                    self._set_controls_enabled(
                        True
                    )

        except queue.Empty:
            pass

        except Exception as exc:

            try:
                self._append_log(
                    (
                        "[ERROR] GUI update error: "
                        f"{exc.__class__.__name__}: "
                        f"{exc}"
                    )
                )
            except Exception:
                pass

        finally:

            try:
                self.after(
                    100,
                    self._poll_queue,
                )
            except Exception:
                pass

    def _on_close(
        self,
    ) -> None:

        if self.processing:

            self._append_log(
                (
                    "[INFO] Processing is still "
                    "running. Please wait for "
                    "the current batch to complete."
                )
            )

            return

        self.destroy()


# ============================================================================
# ENTRY POINT
# ============================================================================

def main() -> None:

    try:

        app = Invoice2ExcelApp()

        app.mainloop()

    except Exception as exc:

        try:

            from tkinter import messagebox

            messagebox.showerror(
                APP_NAME,
                (
                    "Application startup failed:"
                    f"\n\n{exc}"
                ),
            )

        except Exception:
            pass


if __name__ == "__main__":
    main()
