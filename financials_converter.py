from __future__ import annotations

import re
import shutil
import tempfile
import uuid
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Iterable

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


STATEMENTS = ("Income Statement", "Balance Sheet", "Cash Flow Statement")

INCOME_KEYS = (
    "revenue",
    "sale",
    "sales",
    "income",
    "expense",
    "expenses",
    "cost",
    "profit",
    "loss",
    "tax",
    "earnings",
    "share",
    "comprehensive",
    "depreciation",
    "amortisation",
    "amortization",
    "finance",
)

BALANCE_KEYS = (
    "asset",
    "assets",
    "property",
    "equipment",
    "capital work",
    "investment",
    "goodwill",
    "intangible",
    "inventory",
    "inventories",
    "receivable",
    "cash",
    "bank",
    "equity",
    "capital",
    "liabilit",
    "borrowings",
    "lease",
    "payable",
    "provision",
    "deferred tax",
)

CASH_FLOW_KEYS = (
    "cash",
    "operating activities",
    "investing activities",
    "financing activities",
    "depreciation",
    "amortization",
    "amortisation",
    "stock-based",
    "inventories",
    "receivable",
    "payable",
    "purchase",
    "proceeds",
    "repayment",
    "borrowings",
    "dividend",
    "interest",
    "tax paid",
    "net cash generated",
    "net cash provided",
    "exchange",
    "lease",
)

SUMMARY_LABELS = {
    "Revenue": (
        "revenue from operations",
        "total net sales",
        "total revenue from operations",
        "net sales",
    ),
    "PBT": ("profit before tax", "income before income taxes", "before income taxes"),
    "PAT": ("profit for the year", "net income", "net profit for the"),
    "Total Assets": ("total assets",),
    "Total Equity": ("total equity", "total stockholders' equity", "total shareholders' equity"),
    "Operating Cash Flow": (
        "net cash generated from operating activities",
        "net cash provided by operating activities",
        "net cash provided by (used in) operating activities",
    ),
}


@dataclass
class ParsedPdf:
    company: str
    source_name: str
    source_path: Path
    periods: list[str] = field(default_factory=list)
    # statements[stmt][canonical_key] = {period: value}. The canonical key is used
    # only to MERGE the same line across years; the verbatim label shown in the
    # sheet is kept in display_labels, and the PDF row position in row_order.
    statements: dict[str, dict[str, dict[str, float]]] = field(default_factory=dict)
    display_labels: dict[str, dict[str, str]] = field(default_factory=dict)
    row_order: dict[str, dict[str, int]] = field(default_factory=dict)
    skipped_reason: str | None = None


@dataclass
class ConversionResult:
    output_paths: list[Path]
    summaries: dict[str, dict[str, dict[str, float | None]]]
    skipped: list[str]


def convert_pdfs(pdf_paths: Iterable[Path], output_dir: Path) -> ConversionResult:
    output_dir.mkdir(parents=True, exist_ok=True)
    parsed = [parse_pdf(Path(path)) for path in pdf_paths]

    valid = [item for item in parsed if not item.skipped_reason]
    skipped = [
        f"{item.source_name}: {item.skipped_reason}" for item in parsed if item.skipped_reason
    ]
    grouped: dict[str, list[ParsedPdf]] = {}
    for item in valid:
        grouped.setdefault(item.company, []).append(item)

    output_paths: list[Path] = []
    summaries: dict[str, dict[str, dict[str, float | None]]] = {}

    for company, filings in grouped.items():
        filings.sort(key=lambda item: _period_sort_key(item.periods))
        company_periods = _combined_periods(filings)
        # safe_filename the period tag too — a generic fallback like "Period 1"
        # contains a space, which the download route's secure_filename() would
        # rewrite to "_", breaking the download lookup.
        period_tag = safe_filename(company_periods[-1]) if company_periods else "Output"
        out_name = (
            f"{safe_filename(company)}_Consolidated_FS_MultiYear.xlsx"
            if len(company_periods) > 2 or len(filings) > 1
            else f"{safe_filename(company)}_Consolidated_FS_{period_tag}.xlsx"
        )
        out_path = output_dir / out_name
        build_company_workbook(company, filings, out_path)
        output_paths.append(out_path)
        summaries[company] = build_summary(company, filings)

    if len(grouped) >= 2:  # a master only makes sense across 2+ distinct companies
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        master_path = output_dir / f"MASTER_Consolidated_FS_{timestamp}.xlsx"
        build_master_workbook(grouped, master_path)
        output_paths.append(master_path)

    return ConversionResult(output_paths=output_paths, summaries=summaries, skipped=skipped)


# PDFs with more pages than this use the low-memory two-phase reader, so a big
# annual report (e.g. a 90-page US filing) doesn't run pdfminer over every page
# and get OOM-killed on a small instance. Smaller reports keep the original,
# proven full-pdfplumber path unchanged.
LARGE_PDF_PAGES = 30


def _pdf_page_count(path: Path) -> int:
    import pypdfium2 as pdfium

    pdf = pdfium.PdfDocument(str(path))
    try:
        return len(pdf)
    finally:
        pdf.close()


def _scan_text_pdfium(path: Path) -> list[tuple[int, str]]:
    """Cheap, low-memory full-document text scan via pypdfium2 (C-backed; each
    page is read and released immediately). Used only to locate statement pages
    and the company name in large PDFs — not for precise number parsing."""
    import pypdfium2 as pdfium

    pdf = pdfium.PdfDocument(str(path))
    out: list[tuple[int, str]] = []
    try:
        for i in range(len(pdf)):
            page = pdf.get_page(i)
            textpage = page.get_textpage()
            try:
                text = textpage.get_text_range() or ""
            finally:
                textpage.close()
                page.close()
            out.append((i + 1, text))
    finally:
        pdf.close()
    return out


def _extract_pages_pdfplumber(path: Path, page_numbers: set[int]) -> dict[int, str]:
    """Run pdfplumber's precise extraction on ONLY the given pages, flushing each
    page's cache so peak memory stays bounded regardless of document size."""
    result: dict[int, str] = {}
    if not page_numbers:
        return result
    with pdfplumber.open(path) as pdf:
        for i, page in enumerate(pdf.pages):
            page_no = i + 1
            if page_no in page_numbers:
                result[page_no] = page.extract_text(x_tolerance=1, y_tolerance=3) or ""
            page.flush_cache()
    return result


def parse_pdf(path: Path) -> ParsedPdf:
    try:
        page_count = _pdf_page_count(path)
    except Exception:
        page_count = 0  # fall through to the pdfplumber path, which surfaces a real error

    if page_count > LARGE_PDF_PAGES:
        # Low-memory path: cheap full scan to find the statement pages, then
        # deep-parse only those few pages with pdfplumber for accurate numbers.
        try:
            scan_pages = _scan_text_pdfium(path)
            statement_pages = find_statement_pages(scan_pages)
            needed = {p for page_set in statement_pages.values() for p in page_set}
            precise = _extract_pages_pdfplumber(path, needed)
        except Exception as exc:
            return ParsedPdf("Unknown Company", path.name, path, skipped_reason=f"could not read PDF ({exc})")
        # Use precise pdfplumber text on statement pages; the cheap scan text on
        # the rest (only consumed for company-name detection / preview).
        pages = [(page_no, precise.get(page_no, text)) for page_no, text in scan_pages]
    else:
        try:
            with pdfplumber.open(path) as pdf:
                pages = []
                for i, page in enumerate(pdf.pages):
                    text = page.extract_text(x_tolerance=1, y_tolerance=3) or ""
                    pages.append((i + 1, text))
                    # Release pdfplumber's cached layout objects for this page.
                    page.flush_cache()
        except Exception as exc:
            return ParsedPdf("Unknown Company", path.name, path, skipped_reason=f"could not read PDF ({exc})")
        statement_pages = find_statement_pages(pages)
    statement_preview = "\n".join(
        text
        for page_no, text in pages
        if any(page_no in page_set for page_set in statement_pages.values())
    )
    joined_first_pages = "\n".join(text for _, text in pages[:12])
    company = detect_company(statement_preview) or detect_company(joined_first_pages) or path.stem

    if not any(statement_pages.values()):
        return ParsedPdf(company, path.name, path, skipped_reason="no detectable consolidated financial statements")

    parsed = ParsedPdf(company=company, source_name=path.name, source_path=path)
    parsed.statements = {}

    all_periods: list[str] = []
    for statement_name, page_numbers in statement_pages.items():
        statement_text = "\n".join(text for page_no, text in pages if page_no in page_numbers)
        periods = detect_periods(statement_text, statement_name)
        if not periods:
            periods = detect_periods(joined_first_pages + "\n" + statement_text, statement_name)
        if periods:
            all_periods.extend(periods)
        rows, labels, order = parse_statement_rows(statement_text, periods, statement_name)
        if rows:
            parsed.statements[statement_name] = rows
            parsed.display_labels[statement_name] = labels
            parsed.row_order[statement_name] = order

    parsed.periods = _dedupe_periods(all_periods)
    if not parsed.periods:
        parsed.periods = _infer_periods_from_rows(parsed.statements)

    if not parsed.statements:
        parsed.skipped_reason = "consolidated statement pages found but no rows could be parsed"

    return parsed


def detect_company(text: str) -> str | None:
    normalized = normalize_text(text)
    annual_report_match = re.search(
        r"([A-Z][A-Za-z0-9&.,'’() -]{3,}?)\s+Annual Report\s+20\d{2}",
        normalized,
        flags=re.I,
    )
    if annual_report_match:
        return title_company(annual_report_match.group(1).strip(" -.,"))

    members_match = re.search(
        r"To the Members of\s+([A-Z][A-Za-z0-9&.,'’() -]{3,}?)(?:\s+Report on|\s+Basis for|\n)",
        normalized,
        flags=re.I,
    )
    if members_match:
        return title_company(members_match.group(1).strip(" -.,"))

    result_match = re.search(
        r"^([A-Z][A-Z0-9&.,'’() -]{3,}?(?:LIMITED|LTD\.?|INC\.?|CORPORATION|COMPANY|PLC))\s*$",
        normalized,
        flags=re.I | re.M,
    )
    if result_match:
        return title_company(result_match.group(1).strip(" -.,"))

    patterns = [
        r"([A-Z][A-Z0-9&.,'’() -]{3,}?(?:LIMITED|LTD\.?|INC\.?|CORPORATION|COMPANY|PLC))",
    ]
    for pattern in patterns:
        match = re.search(pattern, normalized, flags=re.I)
        if match:
            candidate = re.sub(r"\s+", " ", match.group(1)).strip(" -.,")
            candidate = re.sub(r"\bAnnual Report\b.*", "", candidate, flags=re.I).strip()
            if len(candidate) > 3:
                return title_company(candidate)
    return None


def detect_company(text: str) -> str | None:
    normalized = normalize_text(text)
    for line in normalized.splitlines():
        annual_report_match = re.search(
            r"^([A-Z][A-Za-z0-9&.,'() -]{3,}?)\s+Annual Report\s+20\d{2}",
            line.strip(),
            flags=re.I,
        )
        if annual_report_match:
            return title_company(annual_report_match.group(1).strip(" -.,"))

        exact_company = re.search(
            r"^([A-Z][A-Z0-9&.,'() -]{3,}?(?:LIMITED|LTD\.?|INC\.?|CORPORATION|COMPANY|PLC))\s*$",
            line.strip(),
            flags=re.I,
        )
        if exact_company:
            return title_company(exact_company.group(1).strip(" -.,"))

    members_match = re.search(
        r"To the Members of\s+([A-Z][A-Za-z0-9&.,'() -]{3,}?)(?:\s+Report on|\s+Basis for|\n)",
        normalized,
        flags=re.I,
    )
    if members_match:
        return title_company(members_match.group(1).strip(" -.,"))

    generic_match = re.search(
        r"([A-Z][A-Z0-9&.,'() -]{3,}?(?:LIMITED|LTD\.?|INC\.?|CORPORATION|COMPANY|PLC))",
        normalized,
        flags=re.I,
    )
    if generic_match:
        return title_company(generic_match.group(1).strip(" -.,"))
    return None


# Title phrases that, when they START a heading line, identify a primary
# consolidated statement page (covers both US 10-K and Indian annual-report
# wording).
_TITLE_PATTERNS = {
    "Income Statement": (
        r"consolidated statements? of operations",
        r"consolidated statements? of income",
        r"consolidated statement of profit and loss",
        r"consolidated statement of profit",
        r"statement of audited consolidated financial results",
    ),
    "Balance Sheet": (
        r"consolidated balance sheets?",
    ),
    "Cash Flow Statement": (
        r"consolidated statements? of cash flows?",
        r"consolidated cash flows? statement",
        r"consolidated cash flow statements?",
    ),
}

# Pages that mention the statement titles only in prose / listings — never the
# primary statement page itself.
_PAGE_EXCLUDE_MARKERS = (
    "index to consolidated financial statements",
    "report of independent registered public accounting firm",
    "we have audited",
    "item 8. financial statements",
    "item 15",
    "part iv",
)

_PERIOD_HEADER_RE = re.compile(
    r"(year ended|quarter ended|months ended|period ended|as at|as of|"
    r"(?:january|february|march|april|may|june|july|august|september|october|november|december)"
    r"\s+\d{1,2}|\b20\d{2}\b)",
    re.I,
)


def _title_heading_score(line: str, patterns: tuple[str, ...]) -> int:
    """Score a line as a statement-title heading. A title that starts the line
    and is followed by nothing / a date / "(in millions)" scores high; an index
    entry ("... 38") or a mid-sentence mention scores 0."""
    stripped = re.sub(r"\s+", " ", line).strip().lower().rstrip(".")
    if not stripped or len(stripped) > 90:
        return 0
    for pat in patterns:
        match = re.match(rf"{pat}\b", stripped)
        if not match:
            continue
        rest = stripped[match.end():].strip(" .,:-")
        if re.fullmatch(r"\d{1,4}", rest):  # index entry: title + page number
            return 0
        if rest == "" or rest.startswith("("):
            return 3
        if len(rest) <= 40 and _PERIOD_HEADER_RE.search(rest):  # "... as at March 31, 2023"
            return 3
        return 0  # title trails into a sentence -> notes / prose
    return 0


def _numeric_line_count(text: str) -> int:
    count = 0
    for line in text.splitlines():
        if len(re.findall(r"\d", line)) >= 3 and re.search(r"\d[\d,]*", line):
            count += 1
    return count


def find_statement_pages(pages: list[tuple[int, str]]) -> dict[str, set[int]]:
    found = {name: set() for name in STATEMENTS}
    best: dict[str, tuple[int, int | None]] = {name: (0, None) for name in STATEMENTS}

    for page_no, text in pages:
        lower = text.lower()
        if any(marker in lower for marker in _PAGE_EXCLUDE_MARKERS):
            continue
        numeric_lines = _numeric_line_count(text)
        if numeric_lines < 5:  # a real statement page is dense with numbers
            continue
        lines = text.splitlines()
        for name, patterns in _TITLE_PATTERNS.items():
            heading = max((_title_heading_score(line, patterns) for line in lines), default=0)
            if heading <= 0:
                continue
            score = heading * 100 + min(numeric_lines, 60)
            if score > best[name][0]:
                best[name] = (score, page_no)

    for name, (_, page_no) in best.items():
        if page_no is not None:
            found[name].add(page_no)

    # A statement can spill onto the following page (cash-flow financing section,
    # balance-sheet equity half). Include the next page when it continues the
    # table and is not itself another statement's title page.
    page_text = {pno: txt for pno, txt in pages}
    for name in ("Cash Flow Statement", "Balance Sheet"):
        for page_no in list(found[name]):
            nxt = page_text.get(page_no + 1, "")
            if not nxt:
                continue
            if any(
                _title_heading_score(line, pats)
                for pats in _TITLE_PATTERNS.values()
                for line in nxt.splitlines()
            ):
                continue
            low = nxt.lower()
            if _numeric_line_count(nxt) < 5:
                continue
            if name == "Cash Flow Statement":
                keys = ("financing activities", "net cash", "cash and cash equivalents", "continued")
                if any(key in low for key in keys):
                    found[name].add(page_no + 1)
            else:  # Balance Sheet — only the Indian two-page (assets | equity+liabilities) split,
                   # never the separate Statement of Stockholders' Equity roll-forward.
                if "stockholders" in low and "equity" in low and "statement" in low:
                    continue
                keys = ("total equity and liabilities", "total liabilities and equity", "continued")
                if any(key in low for key in keys):
                    found[name].add(page_no + 1)
    return found


def detect_periods(text: str, statement_name: str) -> list[str]:
    normalized = normalize_text(text)
    lines = [line.strip() for line in normalized.splitlines() if line.strip()]

    if "quarter ended" in normalized.lower() and "year ended" in normalized.lower():
        dated_years = re.findall(r"\b\d{1,2}[./-]\d{1,2}[./-](20\d{2})\b", normalized)
        if len(dated_years) >= 2:
            return [f"FY{year}" for year in _dedupe(dated_years[-2:])]

    for idx, line in enumerate(lines[:20]):
        if "year ended december 31" in line.lower():
            window = " ".join(lines[idx : idx + 3])
            years = re.findall(r"\b(20\d{2})\b", window)
            if years:
                return [f"FY{year}" for year in _dedupe(years)]

    # A column header anchored on a month/day date — "December 31," or
    # "As at March 31, 2023" — possibly with the years on the next line. Grab the
    # years from a short window. Covers US balance sheets and Indian statements
    # whose period wording isn't "year ended". (normalize_text can fuse the date
    # line and year line, so search a window rather than a single line.)
    month = r"(january|february|march|april|may|june|july|august|september|october|november|december)"
    for idx, line in enumerate(lines[:20]):
        if re.search(rf"{month}\s+\d{{1,2}}", line.lower()):
            window = " ".join(lines[idx : idx + 3])
            years = re.findall(r"\b(20\d{2})\b", window)
            if years:
                return [f"FY{year}" for year in _dedupe(years)]

    date_patterns = [
        r"\b\d{1,2}\s+(?:March|Mar|December|Dec)\s+(20\d{2})\b",
        r"\b(?:March|Mar|December|Dec)\s+\d{1,2},?\s+(20\d{2})\b",
        r"\b\d{1,2}[./-]\d{1,2}[./-](20\d{2})\b",
    ]
    years: list[str] = []
    header = "\n".join(lines[:30])
    period_search_space = normalized if statement_name == "Cash Flow Statement" else header
    for pattern in date_patterns:
        years.extend(re.findall(pattern, period_search_space, flags=re.I))
    years = _dedupe(years)

    if not years:
        bare_year_lines = [line for line in lines[:15] if re.fullmatch(r"(?:20\d{2}\s*){2,4}", line)]
        if bare_year_lines:
            years = re.findall(r"20\d{2}", bare_year_lines[0])

    if statement_name == "Balance Sheet" and len(years) > 2:
        years = years[-2:]

    return [f"FY{year}" for year in years]


def parse_statement_rows(
    text: str, periods: list[str], statement_name: str
) -> tuple[dict[str, dict[str, float]], dict[str, str], dict[str, int]]:
    periods = periods or []
    period_count = max(1, len(periods))
    rows: dict[str, dict[str, float]] = {}
    labels: dict[str, str] = {}  # canonical key -> clean display label
    order: dict[str, int] = {}   # canonical key -> first line position (PDF row order)

    for idx, raw_line in enumerate(normalize_text(text).splitlines()):
        line = raw_line.strip()
        if not line or len(line) < 4:
            continue
        lower = line.lower()
        if any(skip in lower for skip in ("see accompanying notes", "corporate overview", "statutory reports")):
            continue

        numbers = extract_numbers(line)
        # A real statement row has exactly the period columns (allow one stray,
        # e.g. a note reference). The statement page is detected precisely, so
        # there are no footnote tables here to capture extra numbers from.
        if not (period_count <= len(numbers) <= period_count + 1):
            continue
        text_part = strip_numbers_from_label(line)
        if not re.search(r"[A-Za-z]", text_part) or len(text_part.strip()) < 3:
            continue
        # Skip the period-header line itself ("Year Ended December 31, 2022 2023"),
        # whose "numbers" are just years/dates, not financial values.
        if re.search(_MONTH_RE, lower) and all(1900 <= abs(n) <= 2100 for n in numbers):
            continue
        if lower.startswith("year ended") or "months ended" in lower or lower.strip() in ("period", "particulars"):
            continue

        values = numbers[:period_count] if statement_name == "Balance Sheet" else numbers[-period_count:]
        raw_label = re.sub(r"\s+", " ", text_part).strip(" -:;*")
        key = standardize_label(raw_label, statement_name)  # merge key only
        if not key or len(key) < 3:
            continue
        if re.fullmatch(r"[ivxlcdm.\s()-]+", key.lower()):
            continue

        if not periods:
            periods = [f"Period {i + 1}" for i in range(period_count)]
        rows[key] = {period: value for period, value in zip(periods[-len(values) :], values)}
        labels[key] = clean_label(raw_label)  # polished Title Case for display
        order.setdefault(key, idx)

    return rows, labels, order


def normalize_text(text: str) -> str:
    text = text.replace("\u00a0", " ")
    text = text.replace("`", "₹")
    text = text.replace("’", "'")
    text = text.replace("—", "-").replace("–", "-")
    text = re.sub(r"(?<=\d)\s*,\s*(?=\d)", ",", text)
    text = re.sub(r"(?<=\d)\s*\.\s*(?=\d)", ".", text)
    text = re.sub(r"(?<=\d)\s+(?=\d,\d{3})", "", text)
    text = re.sub(r"(\.\d{2})(?=\d{1,3},\d{3}\.\d{2})", r"\1 ", text)
    return text


def extract_numbers(line: str) -> list[float]:
    matches = re.findall(r"\(?-?\d[\d,]*(?:\.\d+)?\)?", line)
    values: list[float] = []
    for token in matches:
        if re.fullmatch(r"\d{1,2}", token):
            continue
        negative = token.startswith("(") and token.endswith(")")
        clean = token.strip("()").replace(",", "")
        try:
            value = float(clean)
        except ValueError:
            continue
        values.append(-value if negative else value)
    return values


def strip_numbers_from_label(line: str) -> str:
    label = re.sub(r"\(?-?\d[\d,]*(?:\.\d+)?\)?", " ", line)
    label = re.sub(r"^\s*(?:[IVXLCDM]+|[A-Z]|\(?[a-z]\)?|S\.?No\.?)\s+", "", label, flags=re.I)
    label = re.sub(r"\s+", " ", label)
    label = label.replace("$", "").replace("₹", "").strip(" :-")
    return label


def standardize_label(label: str, statement_name: str) -> str:
    label = re.sub(r"\bNote\b", "", label, flags=re.I)
    label = re.sub(r"\bRefer note\b", "", label, flags=re.I)
    label = re.sub(r"\s+", " ", label).strip(" -:;")
    lower = label.lower()

    if statement_name == "Balance Sheet":
        # Grand total ("Total liabilities and stockholders' equity") must be
        # caught before the bare "total liabilities" check, or it gets mislabeled
        # as Total Liabilities with the wrong (= total assets) value.
        if "total liabilities and" in lower or "total equity and liabilities" in lower:
            return "Total Equity and Liabilities"
        if "total current assets" in lower:
            return "Total Current Assets"
        if "total current liabilities" in lower:
            return "Total Current Liabilities"
        if "total assets" in lower:
            return "Total Assets"
        if "total stockholders" in lower or "total shareholders" in lower or "total equity" in lower:
            return "Total Equity"
        if "total liabilities" in lower:
            return "Total Liabilities"
    if statement_name == "Income Statement":
        if "sale of goods" in lower:
            return "Sale of Goods"
        if "other operating revenues" in lower:
            return "Other Operating Revenues"
        # Canonicalize across year-to-year wording changes so the same concept
        # is ONE row (e.g. "Income (loss) before income taxes" == "Income before
        # income taxes"; "Net income (loss)" == "Net income").
        if "before income tax" in lower or "profit before tax" in lower or "profit/(loss) before tax" in lower:
            return "Profit Before Tax"
        if "for income tax" in lower and ("provision" in lower or "benefit" in lower):
            return "Provision for Income Taxes"
        if "technology and" in lower and ("content" in lower or "infrastructure" in lower):
            return "Technology and Infrastructure"
        if (
            "profit for the year" in lower
            or "net income" in lower
            or "net profit" in lower
            or "profit/(loss) for the" in lower
        ):
            return "Profit for the Year"

    if statement_name == "Cash Flow Statement":
        # Wording of this line drifts year to year ("... and other" vs
        # "... non-marketable investments, and other"); collapse to one row.
        if "acquisitions" in lower and "net of cash acquired" in lower:
            return "Acquisitions, Net of Cash Acquired, and Other"

    replacements = {
        "total revenue from operations": "Revenue from Operations",
        "total net sales": "Total Net Sales",
        "net income": "Net Income",
        "profit before tax": "Profit Before Tax",
        "income before income taxes": "Income Before Income Taxes",
        "profit for the year": "Profit for the Year",
        "net profit for the period / year": "Profit for the Year",
        "total assets": "Total Assets",
        "total equity": "Total Equity",
        "total liabilities": "Total Liabilities",
        "net cash provided by operating activities": "Net Cash Provided by Operating Activities",
        "net cash generated from operating activities": "Net Cash Generated from Operating Activities",
        "net cash provided by (used in) operating activities": "Net Cash Provided by Operating Activities",
        "net cash used in investing activities": "Net Cash Used in Investing Activities",
        "net cash used in financing activities": "Net Cash Used in Financing Activities",
    }
    for needle, replacement in replacements.items():
        if needle in lower:
            return replacement
    return title_label(label)


def title_company(value: str) -> str:
    value = re.sub(r"\s+", " ", value).strip()
    if value.isupper():
        value = value.title()
    value = value.replace("Limited", "Limited").replace("Inc.", "Inc.")
    value = value.replace("Britannia Industries Limited", "Britannia Industries Limited")
    value = value.replace("Amazon.Com", "Amazon.com")
    return value


_SMALL_WORDS = {"of", "and", "the", "for", "to", "in", "on", "a", "an", "or",
                "per", "by", "with", "from", "as", "at"}
_MONTH_RE = (
    r"(january|february|march|april|may|june|july|august|september|october|"
    r"november|december)\s+\d{1,2}"
)


def clean_label(label: str) -> str:
    """Polished finance Title Case for display, e.g. 'net product sales' ->
    'Net Product Sales', 'cost of sales' -> 'Cost of Sales'. Keeps small words
    lowercase (except first), preserves hyphen casing and parenthetical words."""
    label = re.sub(r"\s+", " ", label).strip(" -:;*,.")
    if not label:
        return label

    def fix(token: str, first: bool) -> str:
        out = []
        for piece in re.split(r"(-)", token):
            if piece == "-":
                out.append(piece)
                continue
            match = re.match(r"^([^A-Za-z]*)(.*)$", piece)
            lead, rest = match.group(1), match.group(2)
            if rest:
                low = rest.lower()
                if not first and low in _SMALL_WORDS:
                    rest = low
                else:
                    rest = rest[0].upper() + rest[1:].lower()
            out.append(lead + rest)
            first = False
        return "".join(out)

    words = label.split(" ")
    return " ".join(fix(w, i == 0) for i, w in enumerate(words))


def title_label(value: str) -> str:
    keep_upper = {"EPS", "PBT", "PAT", "OCI", "AWS"}
    words = []
    for word in value.split():
        clean = word.strip()
        if clean.upper() in keep_upper:
            words.append(clean.upper())
        elif clean.isupper() and len(clean) <= 4:
            words.append(clean)
        else:
            words.append(clean[:1].upper() + clean[1:])
    return " ".join(words)


def build_company_workbook(company: str, filings: list[ParsedPdf], out_path: Path) -> None:
    wb = Workbook()
    wb.remove(wb.active)
    periods = _combined_periods(filings)
    for statement in STATEMENTS:
        ws = wb.create_sheet(statement)
        write_statement_sheet(ws, company, filings, statement, periods)
    wb.save(out_path)


def build_master_workbook(grouped: dict[str, list[ParsedPdf]], out_path: Path) -> None:
    wb = Workbook()
    wb.remove(wb.active)
    for statement in STATEMENTS:
        ws = wb.create_sheet(statement)
        write_master_sheet(ws, grouped, statement)
    wb.save(out_path)


# ----- analyst-grade statement styling (ported from the A2E layout) -----
_AS_HEAD = PatternFill("solid", fgColor="1B3A5C")
_AS_SUBHEAD = PatternFill("solid", fgColor="FF9900")
_AS_SECTION = PatternFill("solid", fgColor="FFF3E0")
_AS_TOTAL = PatternFill("solid", fgColor="FFE0B2")
_AS_ALT = PatternFill("solid", fgColor="FAFAFA")
_AS_WHITE = PatternFill("solid", fgColor="FFFFFF")
_AS_NOTE = PatternFill("solid", fgColor="FFF9F0")
_AS_YEAR = PatternFill("solid", fgColor="B35900")

_AS_TITLE_FONT = Font(name="Calibri", bold=True, color="FFFFFF", size=14)
_AS_SUB_FONT = Font(name="Calibri", bold=True, color="1B3A5C", size=10)
_AS_YEAR_FONT = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
_AS_SEC_FONT = Font(name="Calibri", bold=True, color="1B3A5C", size=10)
_AS_NORM_FONT = Font(name="Calibri", color="000000", size=10)
_AS_NOTE_FONT = Font(name="Calibri", italic=True, color="595959", size=9)
_AS_THIN = Side(style="thin", color="FFB74D")
_AS_MED = Side(style="medium", color="FF9900")

# Ordered sections per statement and the label keywords that fall under each.
_SECTIONS: dict[str, list[tuple[str, tuple[str, ...]]]] = {
    "Income Statement": [
        ("I.  REVENUE", ("product sales", "service sales", "sale of goods", "revenue from operations",
                         "total net sales", "total revenue", "operating revenue", "net sales", "total income")),
        ("II. OPERATING EXPENSES", ("cost of sales", "cost of materials", "cost of goods", "purchases of stock",
                         "changes in inventories", "fulfillment", "technology", "content", "sales and marketing",
                         "general and administrative", "employee benefit", "depreciation", "other operating expense",
                         "other expense", "total operating expense", "total expense")),
        ("III. OPERATING & NON-OPERATING INCOME", ("operating income", "operating profit", "interest income",
                         "finance income", "interest expense", "finance cost", "other income",
                         "total non-operating", "exceptional")),
        ("IV. TAXES & BOTTOM LINE", ("before income tax", "before tax", "provision for income tax", "benefit",
                         "tax expense", "current tax", "deferred tax", "income tax", "equity-method",
                         "net income", "profit for the year", "profit after tax")),
        ("V.  COMPREHENSIVE INCOME", ("comprehensive",)),
        ("VI. EARNINGS PER SHARE", ("earnings per share", "per share", "weighted", "basic", "diluted")),
    ],
    "Balance Sheet": [
        ("A.  CURRENT ASSETS", ("cash and cash equivalent", "bank balance", "marketable securities",
                         "current investment", "inventories", "accounts receivable", "trade receivable",
                         "prepaid", "other current asset", "total current assets")),
        ("B.  NON-CURRENT ASSETS", ("property and equipment", "property, plant", "fixed asset", "right-of-use",
                         "operating lease", "goodwill", "intangible", "non-current investment",
                         "deferred tax asset", "other asset", "total non-current assets", "total assets")),
        ("C.  CURRENT LIABILITIES", ("accounts payable", "trade payable", "accrued", "unearned", "short-term debt",
                         "short-term borrowing", "current portion", "other current liabilit", "total current liabilities")),
        ("D.  NON-CURRENT LIABILITIES", ("long-term debt", "long-term borrowing", "long-term lease", "lease liabilit",
                         "deferred tax liabilit", "other long-term", "other non-current liabilit",
                         "total non-current liabilities", "total liabilities")),
        ("E.  STOCKHOLDERS' EQUITY", ("common stock", "preferred stock", "share capital", "additional paid-in",
                         "securities premium", "treasury stock", "retained earnings", "reserves", "accumulated",
                         "other equity", "total stockholders", "total shareholders", "total equity",
                         "total liabilities and", "total equity and liabilities")),
    ],
    "Cash Flow Statement": [
        ("A.  OPERATING ACTIVITIES", ("net income", "profit before tax", "profit for the year", "depreciation",
                         "amortization", "stock-based compensation", "deferred income tax", "non-operating expense",
                         "working capital", "changes in", "accounts receivable", "inventories", "accounts payable",
                         "accrued", "unearned", "other assets", "cash generated from operation", "income tax paid",
                         "operating lease assets", "net cash provided by operating", "net cash generated from operating",
                         "net cash from operating", "net cash used in operating", "net cash provided by (used in) operating")),
        ("B.  INVESTING ACTIVITIES", ("purchases of property", "purchase of property", "proceeds from property",
                         "purchases of marketable", "sales and maturities", "acquisitions", "purchase of investment",
                         "net cash used in investing", "net cash provided by (used in) investing",
                         "net cash from investing", "net cash generated from investing")),
        ("C.  FINANCING ACTIVITIES", ("repurchased", "buyback", "proceeds from long-term debt",
                         "repayments of long-term debt", "proceeds from short-term", "repayments of short-term",
                         "principal repayments of finance", "principal repayments of financing", "dividend",
                         "proceeds from issuance", "net cash used in financing",
                         "net cash provided by (used in) financing", "net cash from financing")),
        ("D.  NET CHANGE IN CASH", ("foreign currency effect", "net increase", "net decrease",
                         "cash, cash equivalents", "cash and cash equivalents")),
    ],
}


def _classify_section(statement: str, label: str) -> int:
    """Index of the section this line belongs to (longest keyword match wins).
    -1 means unclassified (rendered after the sections, ungrouped)."""
    lower = label.lower()
    best_idx, best_len = -1, 0
    for idx, (_, keywords) in enumerate(_SECTIONS.get(statement, [])):
        for kw in keywords:
            if kw in lower and len(kw) > best_len:
                best_idx, best_len = idx, len(kw)
    return best_idx


_DROP_LABELS = {"basic", "diluted", "period", "marketing", "december", "particulars",
                "issued shares - and", "outstanding shares - and"}


def _drop_row(label: str) -> bool:
    """Known terse footnote fragments that broadened extraction picks up."""
    return label.strip().lower() in _DROP_LABELS


def _is_noise_label(label: str) -> bool:
    low = label.strip().lower()
    if len(low) < 6 or len(low.split()) < 2:
        return True
    if low.endswith((" and", " - and", " -", " of", " or", ",")):
        return True
    return False


def _is_subtotal(label: str) -> bool:
    low = label.lower().strip()
    if low.startswith("total") or "net cash" in low:
        return True
    return low in {
        "operating income", "operating profit", "gross profit",
        "income (loss) before income taxes", "income before income taxes", "profit before tax",
        "net income", "net income (loss)", "profit for the year",
        "comprehensive income (loss)", "comprehensive income",
    }


def _as_row(ws, row, label, vals, indent=0, total=False, ncols=0):
    fill = _AS_TOTAL if total else (_AS_ALT if row % 2 == 0 else _AS_WHITE)
    font = _AS_SEC_FONT if total else _AS_NORM_FONT
    border = Border(top=_AS_THIN, bottom=_AS_MED) if total else Border(bottom=_AS_THIN)
    c1 = ws.cell(row, 1, "    " * indent + label)
    c1.fill, c1.font, c1.border = fill, font, border
    c1.alignment = Alignment(horizontal="left", vertical="center")
    for i, val in enumerate(vals, 2):
        c = ws.cell(row, i, val)
        c.fill, c.font, c.border = fill, font, border
        c.number_format = "#,##0"
        c.alignment = Alignment(horizontal="right", vertical="center")


def write_statement_sheet(ws, company: str, filings: list[ParsedPdf], statement: str, periods: list[str]) -> None:
    ws.sheet_view.showGridLines = False
    ncols = len(periods)
    last_col = get_column_letter(ncols + 1)
    ws.column_dimensions["A"].width = 56
    for col in range(2, ncols + 2):
        ws.column_dimensions[get_column_letter(col)].width = 15

    # Title block
    ws.merge_cells(f"A1:{last_col}1")
    ws.row_dimensions[1].height = 26
    c = ws.cell(1, 1, company.upper())
    c.fill, c.font = _AS_HEAD, _AS_TITLE_FONT
    c.alignment = Alignment(horizontal="center", vertical="center")
    for col in range(2, ncols + 2):
        ws.cell(1, col).fill = _AS_HEAD
    ws.merge_cells(f"A2:{last_col}2")
    c = ws.cell(2, 1, f"CONSOLIDATED {statement.upper()}")
    c.fill, c.font = _AS_SUBHEAD, _AS_SUB_FONT
    c.alignment = Alignment(horizontal="center", vertical="center")
    span = f"{periods[0]} to {periods[-1]}" if periods else ""
    ws.merge_cells(f"A3:{last_col}3")
    c = ws.cell(3, 1, f"Multi-Year {span}  |  Figures in millions as reported  |  {company}")
    c.fill, c.font = _AS_NOTE, _AS_NOTE_FONT
    c.alignment = Alignment(horizontal="center", vertical="center")

    # Column headers
    ws.row_dimensions[5].height = 26
    ws.freeze_panes = "B6"
    h = ws.cell(5, 1, "Particulars")
    h.fill, h.font = _AS_SUBHEAD, _AS_SUB_FONT
    h.alignment = Alignment(horizontal="center", vertical="center")
    h.border = Border(top=_AS_MED, bottom=_AS_MED)
    for i, period in enumerate(periods, 2):
        c = ws.cell(5, i, f"{period}\n(Mn)")
        c.fill, c.font = _AS_YEAR, _AS_YEAR_FONT
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = Border(top=_AS_MED, bottom=_AS_MED)

    keys = _ordered_keys(filings, statement)
    # group keys by section, preserving statement order within each
    sections = _SECTIONS.get(statement, [])
    buckets: dict[int, list[str]] = {i: [] for i in range(len(sections))}
    leftovers: list[str] = []
    for key in keys:
        label = _display_label(filings, statement, key)
        if _drop_row(label):
            continue
        idx = _classify_section(statement, label)
        if idx < 0 and _is_noise_label(label):
            continue  # unclassified footnote / wrapped-line fragment
        (buckets[idx] if idx >= 0 else leftovers).append(key)

    row = 6
    for sec_idx, (sec_label, _) in enumerate(sections):
        sec_keys = buckets[sec_idx]
        if not sec_keys:
            continue
        ws.row_dimensions[row].height = 16
        for col in range(1, ncols + 2):
            cell = ws.cell(row, col)
            cell.fill = _AS_SECTION
            if col == 1:
                cell.value = sec_label
                cell.font = _AS_SEC_FONT
                cell.alignment = Alignment(horizontal="left", vertical="center")
        row += 1
        for key in sec_keys:
            label = _display_label(filings, statement, key)
            vals = [_value_for_period(filings, statement, key, p) for p in periods]
            total = _is_subtotal(label)
            _as_row(ws, row, label, vals, indent=0 if total else 1, total=total, ncols=ncols)
            row += 1
    for key in leftovers:
        label = _display_label(filings, statement, key)
        vals = [_value_for_period(filings, statement, key, p) for p in periods]
        _as_row(ws, row, label, vals, indent=1, total=_is_subtotal(label), ncols=ncols)
        row += 1


def write_master_sheet(ws, grouped: dict[str, list[ParsedPdf]], statement: str) -> None:
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 56
    companies = sorted(grouped)

    # company -> its period columns; build the flat column plan (with a blank
    # spacer column between companies).
    plan: list[tuple[str, str, int]] = []  # (company, period, col)
    col = 2
    company_spans: list[tuple[str, int, int]] = []
    for company in companies:
        periods = _combined_periods(grouped[company])
        if not periods:
            continue
        start = col
        for period in periods:
            plan.append((company, period, col))
            ws.column_dimensions[get_column_letter(col)].width = 15
            col += 1
        company_spans.append((company, start, col - 1))
        col += 1  # spacer
    last_col_idx = max(col - 2, 2)
    last_col = get_column_letter(last_col_idx)

    # Title block
    ws.merge_cells(f"A1:{last_col}1")
    ws.row_dimensions[1].height = 26
    c = ws.cell(1, 1, f"MASTER — CONSOLIDATED {statement.upper()}")
    c.fill, c.font = _AS_HEAD, _AS_TITLE_FONT
    c.alignment = Alignment(horizontal="center", vertical="center")
    for cc in range(2, last_col_idx + 1):
        ws.cell(1, cc).fill = _AS_HEAD

    # Company group headers (row 3) + period headers (row 4)
    hc = ws.cell(3, 1, "Particulars")
    hc.fill, hc.font = _AS_SUBHEAD, _AS_SUB_FONT
    hc.alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(4, 1).fill = _AS_SUBHEAD
    for company, start, end in company_spans:
        ws.merge_cells(start_row=3, start_column=start, end_row=3, end_column=end)
        cc = ws.cell(3, start, company)
        cc.fill, cc.font = _AS_SUBHEAD, _AS_SUB_FONT
        cc.alignment = Alignment(horizontal="center", vertical="center")
    for company, period, ccol in plan:
        c = ws.cell(4, ccol, f"{period}\n(Mn)")
        c.fill, c.font = _AS_YEAR, _AS_YEAR_FONT
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[4].height = 26
    ws.freeze_panes = "B5"

    # ordered, classified keys (union across companies)
    keys: list[str] = []
    seen: set[str] = set()
    for company in companies:
        for key in _ordered_keys(grouped[company], statement):
            if key not in seen:
                seen.add(key)
                keys.append(key)

    def disp(key: str) -> str:
        return next((lbl for company in companies
                     if (lbl := _display_label(grouped[company], statement, key)) != key), key)

    sections = _SECTIONS.get(statement, [])
    buckets: dict[int, list[str]] = {i: [] for i in range(len(sections))}
    leftovers: list[str] = []
    for key in keys:
        label = disp(key)
        if _drop_row(label):
            continue
        idx = _classify_section(statement, label)
        if idx < 0 and _is_noise_label(label):
            continue
        (buckets[idx] if idx >= 0 else leftovers).append(key)

    def render(row: int, key: str) -> None:
        label = disp(key)
        total = _is_subtotal(label)
        fill = _AS_TOTAL if total else (_AS_ALT if row % 2 == 0 else _AS_WHITE)
        font = _AS_SEC_FONT if total else _AS_NORM_FONT
        border = Border(top=_AS_THIN, bottom=_AS_MED) if total else Border(bottom=_AS_THIN)
        c1 = ws.cell(row, 1, ("" if total else "    ") + label)
        c1.fill, c1.font, c1.border = fill, font, border
        c1.alignment = Alignment(horizontal="left", vertical="center")
        for company, period, ccol in plan:
            c = ws.cell(row, ccol, _value_for_period(grouped[company], statement, key, period))
            c.fill, c.font, c.border = fill, font, border
            c.number_format = "#,##0"
            c.alignment = Alignment(horizontal="right", vertical="center")

    row = 5
    for sec_idx, (sec_label, _) in enumerate(sections):
        if not buckets[sec_idx]:
            continue
        for cc in range(1, last_col_idx + 1):
            cell = ws.cell(row, cc)
            cell.fill = _AS_SECTION
            if cc == 1:
                cell.value = sec_label
                cell.font = _AS_SEC_FONT
        row += 1
        for key in buckets[sec_idx]:
            render(row, key)
            row += 1
    for key in leftovers:
        render(row, key)
        row += 1


def style_row(ws, row_idx: int, max_col: int, palette: "WorkbookPalette", total: bool = False) -> None:
    fill = palette.total if total else (palette.alt if row_idx % 2 == 0 else palette.white)
    for col_idx in range(1, max_col + 1):
        cell = ws.cell(row_idx, col_idx)
        cell.fill = fill
        cell.border = palette.thin_border
        if total:
            cell.font = Font(bold=True, color="1F3864")
        if col_idx > 1:
            cell.number_format = "#,##0.00"
            cell.alignment = Alignment(horizontal="right")


class WorkbookPalette:
    def __init__(self) -> None:
        self.dark = PatternFill("solid", fgColor="1F3864")
        self.mid = PatternFill("solid", fgColor="2E75B6")
        self.alt = PatternFill("solid", fgColor="F2F7FC")
        self.white = PatternFill("solid", fgColor="FFFFFF")
        self.total = PatternFill("solid", fgColor="BDD7EE")
        thin = Side(style="thin", color="B8CCE4")
        medium = Side(style="medium", color="2E75B6")
        self.thin_border = Border(bottom=thin)
        self.header_border = Border(top=medium, bottom=medium)


def build_summary(company: str, filings: list[ParsedPdf]) -> dict[str, dict[str, float | None]]:
    periods = _combined_periods(filings)
    summary = {metric: {} for metric in SUMMARY_LABELS}
    for metric, needles in SUMMARY_LABELS.items():
        for period in periods:
            summary[metric][period] = find_summary_value(filings, metric, needles, period)
    return summary


def find_summary_value(
    filings: list[ParsedPdf], metric: str, needles: tuple[str, ...], period: str
) -> float | None:
    if metric == "Revenue":
        direct = _find_labeled_value(filings, ("Income Statement",), needles, period)
        if direct is not None:
            return direct
        sales = _find_labeled_value(filings, ("Income Statement",), ("sale of goods", "net product sales"), period)
        other_revenue = _find_labeled_value(
            filings,
            ("Income Statement",),
            ("other operating revenues", "net service sales"),
            period,
        )
        if sales is not None and other_revenue is not None:
            return sales + other_revenue
        return sales

    statement_names = {
        "Revenue": ("Income Statement",),
        "PBT": ("Income Statement",),
        "PAT": ("Income Statement",),
        "Total Assets": ("Balance Sheet",),
        "Total Equity": ("Balance Sheet",),
        "Operating Cash Flow": ("Cash Flow Statement",),
    }[metric]
    return _find_labeled_value(filings, statement_names, needles, period)


def _find_labeled_value(
    filings: list[ParsedPdf], statement_names: tuple[str, ...], needles: tuple[str, ...], period: str
) -> float | None:
    for filing in reversed(filings):
        for statement in statement_names:
            for label, values in filing.statements.get(statement, {}).items():
                label_lower = label.lower()
                if period in values and any(needle in label_lower for needle in needles):
                    return values[period]
    return None


def _combined_periods(filings: list[ParsedPdf]) -> list[str]:
    periods: list[str] = []
    for filing in filings:
        periods.extend(filing.periods)
        for statement_rows in filing.statements.values():
            for values in statement_rows.values():
                periods.extend(values.keys())
    return sorted(_dedupe_periods(periods), key=_period_value)


def _dedupe_periods(periods: list[str]) -> list[str]:
    clean = []
    for period in periods:
        if not period:
            continue
        match = re.search(r"20\d{2}", period)
        # Prefer a clean FYxxxx label when a year is present, but keep generic
        # labels (e.g. "Period 1") so values still render when year detection
        # couldn't identify the column headers.
        clean.append(f"FY{match.group(0)}" if match else period)
    return _dedupe(clean)


def _infer_periods_from_rows(statements: dict[str, dict[str, dict[str, float]]]) -> list[str]:
    periods: list[str] = []
    for rows in statements.values():
        for values in rows.values():
            periods.extend(values.keys())
    return _dedupe_periods(periods)


def _period_value(period: str) -> int:
    match = re.search(r"20\d{2}", period)
    return int(match.group(0)) if match else 0


def _period_sort_key(periods: list[str]) -> int:
    values = [_period_value(period) for period in periods]
    return max(values) if values else 0


def _value_for_period(
    filings: list[ParsedPdf], statement: str, label: str, period: str
) -> float | None:
    for filing in reversed(filings):
        value = filing.statements.get(statement, {}).get(label, {}).get(period)
        if value is not None:
            return value
    return None


# Ordered templates that lay each statement out top-to-bottom the way it reads
# on the page. A label is ordered by the LONGEST template phrase it contains, so
# overlaps resolve correctly (e.g. "non-operating income" -> "total non-operating",
# not "operating income").
_STATEMENT_ORDER = {
    "Income Statement": [
        "revenue from operations", "sale of goods", "net product sales",
        "net service sales", "other operating revenue",
        "total net sales", "total revenue", "total income",
        "cost of sales", "cost of materials", "cost of goods", "purchases of stock-in-trade",
        "changes in inventories", "employee benefit", "fulfillment", "technology",
        "sales and marketing", "general and administrative", "depreciation and amorti",
        "other operating expense", "other expense",
        "total operating expense", "total expense",
        "operating income", "operating profit",
        "interest income", "finance income",
        "interest expense", "finance cost",
        "other income", "total non-operating",
        "before exceptional", "exceptional",
        "profit before tax", "before income tax",
        "provision for income tax", "tax expense", "current tax", "deferred tax", "income tax",
        "equity-method", "profit after tax",
        "profit for the year", "net income",
        "other comprehensive", "total comprehensive",
        "basic earnings", "diluted earnings", "earnings per share",
    ],
    "Balance Sheet": [
        "cash and cash equivalent", "bank balance", "marketable securities", "current investment",
        "inventories", "accounts receivable", "trade receivable", "other current asset",
        "total current assets",
        "property and equipment", "property, plant", "fixed asset", "right-of-use", "operating lease",
        "goodwill", "intangible", "non-current investment", "deferred tax asset", "other asset",
        "total non-current assets",
        "total assets",
        "accounts payable", "trade payable", "accrued", "short-term debt", "short-term borrowing",
        "current portion", "other current liabilit", "total current liabilities",
        "long-term debt", "long-term borrowing", "long-term lease", "lease liabilit",
        "deferred tax liabilit", "other long-term", "other non-current liabilit",
        "total non-current liabilities", "total liabilities",
        "common stock", "share capital", "additional paid-in", "securities premium", "treasury stock",
        "retained earnings", "reserves", "accumulated", "other equity",
        "total stockholders", "total shareholders", "total equity", "total equity and liabilities",
    ],
    "Cash Flow Statement": [
        "profit before tax", "net income", "profit for the year",
        "depreciation", "amortization", "stock-based compensation", "stock based compensation",
        "changes in", "accounts receivable", "inventories", "accounts payable",
        "cash generated from operation", "income tax paid",
        "net cash provided by operating", "net cash generated from operating",
        "net cash from operating", "net cash used in operating",
        "purchases of property", "purchase of property", "purchases of marketable", "sales and maturities",
        "acquisitions", "purchase of investment",
        "net cash used in investing", "net cash provided by (used in) investing",
        "net cash from investing", "net cash generated from investing",
        "proceeds from long-term debt", "repayments of long-term debt",
        "proceeds from short-term", "repayments of short-term",
        "principal repayments of finance", "principal repayments of financing", "dividend", "buyback",
        "net cash used in financing", "net cash provided by (used in) financing", "net cash from financing",
        "foreign currency effect", "net increase", "net decrease",
        "cash, cash equivalents", "cash and cash equivalents",
    ],
}


def _ordered_keys(filings: list[ParsedPdf], statement: str) -> list[str]:
    """Canonical keys ordered by each line's AVERAGE normalized position across
    the filings it appears in. Using the average (rather than a single filing's
    order) keeps the PDF's row order even when a given year's report didn't parse
    every line, so nothing gets dumped at the bottom."""
    positions: dict[str, list[float]] = {}
    for filing in filings:
        order = filing.row_order.get(statement, {})
        if not order:
            continue
        span = max(max(order.values()), 1)
        for key, idx in order.items():
            positions.setdefault(key, []).append(idx / span)
    return sorted(positions, key=lambda k: sum(positions[k]) / len(positions[k]))


def _display_label(filings: list[ParsedPdf], statement: str, key: str) -> str:
    """The verbatim label as written on the PDF, preferring the newest filing."""
    for filing in reversed(filings):
        label = filing.display_labels.get(statement, {}).get(key)
        if label:
            return label
    return key


def _preferred_label_order(statement: str, label: str) -> tuple[int, str]:
    template = _STATEMENT_ORDER.get(statement, [])
    lower = label.lower()
    best_idx, best_len = None, -1
    for idx, phrase in enumerate(template):
        if phrase in lower and len(phrase) > best_len:
            best_idx, best_len = idx, len(phrase)
    if best_idx is None:
        return (len(template) + 1, lower)
    return (best_idx, lower)


def _is_total_label(label: str) -> bool:
    lower = label.lower()
    return lower.startswith("total") or "net cash" in lower or lower in {
        "profit for the year",
        "net income",
        "profit before tax",
        "income before income taxes",
    }


def safe_filename(name: str) -> str:
    name = re.sub(r"[^A-Za-z0-9]+", "_", name).strip("_")
    return name or f"Company_{uuid.uuid4().hex[:8]}"


def _dedupe(items: Iterable[str]) -> list[str]:
    seen = set()
    out = []
    for item in items:
        if item not in seen:
            seen.add(item)
            out.append(item)
    return out


def create_processing_dir(base: Path) -> Path:
    path = base / uuid.uuid4().hex
    path.mkdir(parents=True, exist_ok=True)
    return path


def cleanup_dir(path: Path) -> None:
    if path.exists():
        shutil.rmtree(path, ignore_errors=True)


def temporary_workspace() -> Path:
    return Path(tempfile.mkdtemp(prefix="pdf_financials_"))
