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
    "PBT": ("profit before tax", "income before income taxes"),
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
    statements: dict[str, dict[str, dict[str, float]]] = field(default_factory=dict)
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
        out_name = (
            f"{safe_filename(company)}_Consolidated_FS_MultiYear.xlsx"
            if len(company_periods) > 2 or len(filings) > 1
            else f"{safe_filename(company)}_Consolidated_FS_{company_periods[-1] if company_periods else 'Output'}.xlsx"
        )
        out_path = output_dir / out_name
        build_company_workbook(company, filings, out_path)
        output_paths.append(out_path)
        summaries[company] = build_summary(company, filings)

    if len(valid) >= 2:
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
        rows = parse_statement_rows(statement_text, periods, statement_name)
        if rows:
            parsed.statements[statement_name] = rows

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


def find_statement_pages(pages: list[tuple[int, str]]) -> dict[str, set[int]]:
    found = {name: set() for name in STATEMENTS}
    for page_no, text in pages:
        lower = text.lower()
        head = lower[:1400]
        has_cash_flow_anywhere = (
            "consolidated statement of cash flow" in lower
            or "consolidated statements of cash flow" in lower
        )
        if (
            "consolidated" not in head
            and "audited consolidated financial results" not in head
            and not has_cash_flow_anywhere
        ):
            continue
        if "independent auditor" in head:
            continue
        has_statement_header = any(
            marker in head
            for marker in (
                "consolidated balance sheet",
                "consolidated balance sheets",
                "consolidated statement of profit",
                "consolidated statements of operations",
                "statement of audited consolidated financial results",
            )
        ) or has_cash_flow_anywhere
        if "notes to consolidated financial statements" in head and not has_statement_header:
            continue
        if (
            "consolidated statement of profit" in head
            or "consolidated statements of operations" in head
            or "statement of audited consolidated financial results" in head
            or "consolidated statement of income" in head
        ):
            found["Income Statement"].add(page_no)
        if "consolidated balance sheet" in head or "consolidated balance sheets" in head:
            found["Balance Sheet"].add(page_no)
        if has_cash_flow_anywhere:
            found["Cash Flow Statement"].add(page_no)
            if page_no + 1 <= len(pages):
                next_text = pages[page_no][1].lower() if page_no < len(pages) else ""
                if "continued" in next_text or "financing activities" in next_text:
                    found["Cash Flow Statement"].add(page_no + 1)

    for name, page_set in list(found.items()):
        expanded = set(page_set)
        for page_no in page_set:
            if name == "Cash Flow Statement" and page_no + 1 <= len(pages):
                next_text = pages[page_no][1].lower() if page_no < len(pages) else ""
                if any(key in next_text for key in ("financing activities", "net cash", "cash and cash equivalents")):
                    expanded.add(page_no + 1)
        found[name] = expanded
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


def parse_statement_rows(text: str, periods: list[str], statement_name: str) -> dict[str, dict[str, float]]:
    periods = periods or []
    period_count = max(1, len(periods))
    rows: dict[str, dict[str, float]] = {}
    keys = {
        "Income Statement": INCOME_KEYS,
        "Balance Sheet": BALANCE_KEYS,
        "Cash Flow Statement": CASH_FLOW_KEYS,
    }[statement_name]

    for raw_line in normalize_text(text).splitlines():
        line = raw_line.strip()
        if not line or len(line) < 4:
            continue
        lower = line.lower()
        if not any(key in lower for key in keys):
            continue
        if any(skip in lower for skip in ("see accompanying notes", "corporate overview", "statutory reports")):
            continue

        numbers = extract_numbers(line)
        if len(numbers) < period_count:
            continue

        values = numbers[:period_count] if statement_name == "Balance Sheet" else numbers[-period_count:]
        label = strip_numbers_from_label(line)
        label = standardize_label(label, statement_name)
        if not label or len(label) < 3:
            continue
        if re.fullmatch(r"[ivxlcdm.\s()-]+", label.lower()):
            continue

        if not periods:
            periods = [f"Period {idx + 1}" for idx in range(period_count)]
        rows[label] = {period: value for period, value in zip(periods[-len(values) :], values)}

    return rows


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
        if "total equity and liabilities" in lower:
            return "Total Equity and Liabilities"
        if "total assets" in lower:
            return "Total Assets"
        if "total equity" in lower:
            return "Total Equity"
        if "total liabilities" in lower:
            return "Total Liabilities"
    if statement_name == "Income Statement":
        if "sale of goods" in lower:
            return "Sale of Goods"
        if "other operating revenues" in lower:
            return "Other Operating Revenues"
        if "profit before tax" in lower or "income before income taxes" in lower:
            return "Profit Before Tax"
        if "profit for the year" in lower or lower == "net income" or "net profit for the" in lower:
            return "Profit for the Year"

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


def write_statement_sheet(ws, company: str, filings: list[ParsedPdf], statement: str, periods: list[str]) -> None:
    palette = WorkbookPalette()
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "B5"
    ws.column_dimensions["A"].width = 54
    for col_idx in range(2, len(periods) + 2):
        ws.column_dimensions[get_column_letter(col_idx)].width = 16

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(2, len(periods) + 1))
    title = ws.cell(1, 1, company)
    title.fill = palette.dark
    title.font = Font(bold=True, color="FFFFFF", size=14)
    title.alignment = Alignment(horizontal="center")

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max(2, len(periods) + 1))
    subtitle = ws.cell(2, 1, f"Consolidated {statement}")
    subtitle.fill = palette.mid
    subtitle.font = Font(bold=True, color="FFFFFF")
    subtitle.alignment = Alignment(horizontal="center")

    headers = ["Line Item", *periods]
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(4, col_idx, header)
        cell.fill = palette.mid
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center")
        cell.border = palette.header_border

    labels = sorted(
        {
            label
            for filing in filings
            for label in filing.statements.get(statement, {}).keys()
        },
        key=lambda label: _preferred_label_order(statement, label),
    )

    for row_idx, label in enumerate(labels, 5):
        ws.cell(row_idx, 1, label)
        for period in periods:
            value = _value_for_period(filings, statement, label, period)
            ws.cell(row_idx, periods.index(period) + 2, value)
        style_row(ws, row_idx, len(periods) + 1, palette, total=_is_total_label(label))


def write_master_sheet(ws, grouped: dict[str, list[ParsedPdf]], statement: str) -> None:
    palette = WorkbookPalette()
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "B4"
    ws.column_dimensions["A"].width = 54

    companies = sorted(grouped)
    labels = sorted(
        {
            label
            for filings in grouped.values()
            for filing in filings
            for label in filing.statements.get(statement, {}).keys()
        },
        key=lambda label: _preferred_label_order(statement, label),
    )

    ws.cell(1, 1, f"Master Consolidated {statement}")
    ws.cell(1, 1).font = Font(bold=True, size=14, color="FFFFFF")
    ws.cell(1, 1).fill = palette.dark
    ws.cell(3, 1, "Line Item")
    ws.cell(3, 1).font = Font(bold=True, color="FFFFFF")
    ws.cell(3, 1).fill = palette.mid

    col = 2
    for company in companies:
        periods = _combined_periods(grouped[company])
        if not periods:
            continue
        start_col = col
        end_col = col + len(periods) - 1
        ws.merge_cells(start_row=2, start_column=start_col, end_row=2, end_column=end_col)
        cell = ws.cell(2, start_col, company)
        cell.fill = palette.dark
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center")
        for period in periods:
            ws.cell(3, col, period)
            ws.cell(3, col).fill = palette.mid
            ws.cell(3, col).font = Font(bold=True, color="FFFFFF")
            ws.cell(3, col).alignment = Alignment(horizontal="center")
            ws.column_dimensions[get_column_letter(col)].width = 16
            col += 1
        col += 1

    for row_idx, label in enumerate(labels, 4):
        ws.cell(row_idx, 1, label)
        col = 2
        for company in companies:
            filings = grouped[company]
            for period in _combined_periods(filings):
                ws.cell(row_idx, col, _value_for_period(filings, statement, label, period))
                col += 1
            col += 1
        style_row(ws, row_idx, max(1, col - 1), palette, total=_is_total_label(label))


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


def _preferred_label_order(statement: str, label: str) -> tuple[int, str]:
    order = {
        "Income Statement": [
            "revenue",
            "sales",
            "income",
            "expense",
            "profit before",
            "tax",
            "profit for",
            "net income",
            "comprehensive",
            "earnings",
        ],
        "Balance Sheet": [
            "asset",
            "current asset",
            "total assets",
            "equity",
            "liabilit",
            "total equity",
        ],
        "Cash Flow Statement": [
            "operating",
            "net cash generated",
            "net cash provided",
            "investing",
            "financing",
            "cash and cash equivalents",
        ],
    }[statement]
    lower = label.lower()
    for idx, needle in enumerate(order):
        if needle in lower:
            return (idx, label)
    return (99, label)


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
