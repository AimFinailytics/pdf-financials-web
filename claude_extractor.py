"""A2E-grade extraction via the Claude API.

The manual "A2E" gold standard IS Claude reading the report and structuring it,
so this uses the Anthropic API to reproduce that quality. Same interface and
output shape as gemini_fallback.extract(); when ANTHROPIC_API_KEY is set this is
the preferred engine.

Configure with:
  ANTHROPIC_API_KEY -> key from https://console.anthropic.com/settings/keys
  CLAUDE_MODEL      -> optional, defaults to "claude-sonnet-4-6"
                       (set "claude-opus-4-8" to match the chat's A2E output exactly)

Degrades to a no-op (returns None) when the key or SDK is absent.
"""

from __future__ import annotations

import os

# Reuse the lenient JSON loader + number coercion; this module has its own
# compact-format shaper.
from gemini_fallback import _loads_lenient, _to_number

STATEMENTS = ("Income Statement", "Balance Sheet", "Cash Flow Statement")


def _to_structured(data: dict) -> dict:
    """Parse the structured layout rows into the ParsedPdf shape + a render layout.

    Each statement is an ordered list of rows, each row a typed array:
      ["S", "SECTION NAME"]                section header
      ["H", "Sub-group header"]            sub-group header (no numbers)
      ["L", "Line item", v1, v2, ...]      line item, values in `periods` order
      ["T", "Total label", v1, v2, ...]    subtotal / total line
    """
    periods = [str(p) for p in (data.get("periods") or []) if p]
    out: dict = {"periods": periods, "statements": {}, "layout": {}}
    statements = data.get("statements") or data
    for stmt in STATEMENTS:
        items = statements.get(stmt) or []
        rows: dict[str, dict[str, float]] = {}
        labels: dict[str, str] = {}
        order: dict[str, int] = {}
        layout: list[tuple[str, str]] = []
        seq = 0
        for item in items:
            if not item or not isinstance(item, (list, tuple)):
                continue
            tag = str(item[0]).strip().upper()
            if tag in ("S", "H"):
                text = str(item[1]).strip() if len(item) > 1 else ""
                if text:
                    layout.append((tag, text.upper() if tag == "S" else text))
                continue
            # default any other tag to a line item
            if tag not in ("L", "T") or len(item) < 2:
                # tolerate a bare [label, v1, v2...] row (no type marker)
                if len(item) >= 2 and not str(item[0]).strip().upper() in ("S", "H"):
                    tag, body = "L", list(item)
                else:
                    continue
            else:
                body = item[1:]
            label = str(body[0]).strip()
            if not label:
                continue
            values: dict[str, float] = {}
            for j, period in enumerate(periods):
                if j + 1 < len(body):
                    num = _to_number(body[j + 1])
                    if num is not None:
                        values[period] = num
            key = label.lower()
            # keep first occurrence's key stable; suffix dupes so both render
            if key in rows:
                key = f"{key}#{seq}"
            rows[key] = values
            labels[key] = label
            order[key] = seq
            layout.append((tag, key))
            seq += 1
        if rows:
            out["statements"][stmt] = {"rows": rows, "labels": labels, "order": order}
            out["layout"][stmt] = layout
    return out

_SYSTEM = """You reproduce an equity analyst's hand-built consolidated financial model
from the raw text of one company's annual/quarterly report.

For EACH statement (Income Statement, Balance Sheet, Cash Flow Statement), output an
ORDERED list that reproduces the statement top-to-bottom, grouped into the standard
analyst sections. Every element is ONE typed array:
  ["S", "SECTION NAME"]                 a section header (UPPERCASE)
  ["H", "Sub-group header"]             a grouping header that has NO numbers
  ["L", "Line item label", v1, v2, ...] a normal line item, one value per period
  ["T", "Total/Subtotal label", v1, ...]a subtotal or total line

Rules:
- Use ONLY consolidated figures. Never invent, compute, or merge lines. Include
  EVERY line item that appears in the source.
- Values are in the SAME order as "periods" (oldest -> newest); use null for a
  period with no value. Output plain numbers (no currency symbols / thousands
  commas; negatives as -1234.56). Parentheses or a dash = negative/zero.
- Read digits split by spaces in a column ("4 69.21" -> 469.21) in context. For a
  quarterly page with several columns, take the full-year / period-end columns.
- Use clean Title-Case labels ("Revenues from Franchised Restaurants",
  "Depreciation and Amortization"), NOT raw lowercase text.
- Mark every total/subtotal ("Total ...", "Net income", "Operating income",
  net-cash lines, etc.) as "T", not "L".
- Section order to use:
  * INCOME STATEMENT: REVENUES; OPERATING COSTS AND EXPENSES; OPERATING &
    NON-OPERATING INCOME; TAXES & BOTTOM LINE; COMPREHENSIVE INCOME (FOLD IN the
    separate Statement of Comprehensive Income if the text contains it); PER-SHARE
    DATA (EPS basic/diluted, dividends per share, weighted-average shares).
  * BALANCE SHEET: CURRENT ASSETS; NON-CURRENT ASSETS; CURRENT LIABILITIES;
    NON-CURRENT LIABILITIES; SHAREHOLDERS' EQUITY. Add ["H", ...] sub-group headers
    where the report groups lines (e.g. "Company-Owned Restaurant Expenses").
  * CASH FLOW: OPERATING ACTIVITIES; INVESTING ACTIVITIES; FINANCING ACTIVITIES;
    NET CHANGE IN CASH.

Return STRICT JSON only (no prose, no code fences), exactly:
{
  "periods": ["FY2023", "FY2024", "FY2025"],
  "statements": {
    "Income Statement": [
      ["S", "REVENUES"],
      ["L", "Revenues from Franchised Restaurants", 15437, 15715, 16548],
      ["T", "Total Revenues", 25494, 25920, 26885]
    ],
    "Balance Sheet": [],
    "Cash Flow Statement": []
  }
}"""


def is_configured() -> bool:
    return bool(os.environ.get("ANTHROPIC_API_KEY", "").strip())


def extract(statement_texts: dict[str, str], periods: list[str]) -> dict | None:
    """Run Claude extraction over the statement-region text. Returns the same
    shape as gemini_fallback.extract(), or None if unavailable/failed."""
    if not is_configured():
        return None
    body = "\n\n".join(f"=== {n} ===\n{t}" for n, t in statement_texts.items() if t).strip()
    if not body:
        return None
    try:
        import anthropic

        client = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"].strip())
        model = os.environ.get("CLAUDE_MODEL", "claude-sonnet-4-6").strip()
        user = (
            f"Detected periods (may be empty): {', '.join(periods) or 'none detected'}\n\n"
            f"STATEMENT TEXT:\n{body[:160_000]}"
        )
        resp = client.messages.create(
            model=model,
            max_tokens=8192,
            temperature=0,
            # Cache the long instruction block so repeat conversions are cheaper.
            system=[{"type": "text", "text": _SYSTEM, "cache_control": {"type": "ephemeral"}}],
            messages=[{"role": "user", "content": user + "\n\nReturn only the JSON object."}],
        )
        raw = "".join(block.text for block in resp.content if getattr(block, "type", "") == "text")
        data = _loads_lenient(raw)
        if not data:
            return None
        return _to_structured(data)
    except Exception as exc:  # noqa: BLE001 — never break the request path
        print(f"[claude_extractor] extraction unavailable: {exc}")
        return None
