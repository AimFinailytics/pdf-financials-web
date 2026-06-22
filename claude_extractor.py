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


def _to_shape_compact(data: dict) -> dict:
    """Parse the compact `[label, v1, v2, ...]` rows into the ParsedPdf shape."""
    periods = [str(p) for p in (data.get("periods") or []) if p]
    out: dict = {"periods": periods, "statements": {}}
    for stmt in STATEMENTS:
        items = data.get(stmt) or []
        rows: dict[str, dict[str, float]] = {}
        labels: dict[str, str] = {}
        order: dict[str, int] = {}
        for i, item in enumerate(items):
            if not item or not isinstance(item, (list, tuple)):
                continue
            label = str(item[0]).strip()
            if not label:
                continue
            values: dict[str, float] = {}
            for j, period in enumerate(periods):
                if j + 1 < len(item):
                    num = _to_number(item[j + 1])
                    if num is not None:
                        values[period] = num
            if not values:
                continue
            key = label.lower()
            rows[key] = values
            labels[key] = label
            order[key] = i
        if rows:
            out["statements"][stmt] = {"rows": rows, "labels": labels, "order": order}
    return out

_SYSTEM = """You are a meticulous financial-statements data extractor.

You receive the raw text of the CONSOLIDATED financial-statement pages from one
company's annual or quarterly report. Extract EVERY line item with its numeric
value for each reporting period, exactly as the manual analyst gold-standard does.

Rules:
- Use ONLY the consolidated figures. Never invent, compute, or infer new lines.
- Find all three statements even if the source text is messy, multi-column, or the
  digits are split by spaces (e.g. "4 69.21" in a table column means 469.21 — read
  it in the column's context).
- A quarterly results page may show several columns (quarter + year-to-date); take
  the FULL-YEAR / period-end columns, not the quarter-only ones.
- Preserve sign: parentheses or a dash mean negative/zero; a bare "-" is 0.
- Output plain numbers (strip currency symbols and thousands commas; negatives as
  -1234.56).
- Map each period to a label like "FY2024" using the column's year; order periods
  oldest -> newest.
- Classify each line into exactly one of: "Income Statement", "Balance Sheet",
  "Cash Flow Statement".
- Keep the original line-item wording as the label (cleaned of note references).
- Maintain the statement's top-to-bottom order.

Return STRICT JSON only (no prose, no code fences), in this COMPACT shape — each
line is [label, value_for_period_1, value_for_period_2, ...] with the values in
the SAME order as "periods" (use null for a missing value):
{
  "periods": ["FY2023", "FY2024"],
  "Income Statement": [["Revenue from Operations", 16300.55, 16769.27]],
  "Balance Sheet": [],
  "Cash Flow Statement": []
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
        return _to_shape_compact(data)
    except Exception as exc:  # noqa: BLE001 — never break the request path
        print(f"[claude_extractor] extraction unavailable: {exc}")
        return None
