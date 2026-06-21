"""Opt-in AI-assisted extraction via the free Google AI Studio (Gemini) tier.

This is a FALLBACK only. The deterministic parser in financials_converter.py runs
first and on-server; this module is invoked solely when (a) the user explicitly
ticked "AI-assisted extraction" AND (b) the deterministic parse came back with low
confidence. It sends only the already-extracted *text* of the detected statement
pages (never the raw PDF, never the whole document) to Gemini Flash, and returns
structured line items that get merged back into the workbook builder.

Configure with one environment variable:
  GEMINI_API_KEY   -> a free key from https://aistudio.google.com/apikey
  GEMINI_MODEL     -> optional, defaults to "gemini-2.5-flash" (free-tier Flash model)

If the key or library is absent, every function here degrades to a no-op (returns
None), so the app runs perfectly fine without AI configured.
"""

from __future__ import annotations

import json
import os
import re

STATEMENTS = ("Income Statement", "Balance Sheet", "Cash Flow Statement")

_PROMPT = """You are a meticulous financial-statements data extractor.

Below is the raw text of the CONSOLIDATED financial statement pages from one
company's annual or quarterly report. Extract EVERY line item with its numeric
value for each reporting period.

Rules:
- Use ONLY the consolidated figures shown. Do not invent or compute new lines.
- Preserve the sign: values in parentheses or shown as a dash are negative/zero.
  A bare dash "-" means 0.
- Keep numbers exactly as printed (strip currency symbols and thousands commas;
  output plain numbers, negatives as -1234.56).
- Map each period to a label like "FY2024" using the column's year. Order periods
  oldest -> newest.
- Classify each line into exactly one statement: "Income Statement",
  "Balance Sheet", or "Cash Flow Statement".
- Keep the original line-item wording as the label (cleaned of note references).
- Maintain the top-to-bottom order in which lines appear.

Detected periods (may be empty): {periods}

Return STRICT JSON only, no prose, in this exact shape:
{{
  "periods": ["FY2023", "FY2024"],
  "statements": {{
    "Income Statement": [
      {{"label": "Revenue from Operations", "values": {{"FY2023": 16300.55, "FY2024": 16769.27}}}}
    ],
    "Balance Sheet": [],
    "Cash Flow Statement": []
  }}
}}

STATEMENT TEXT:
{body}
"""


def is_configured() -> bool:
    return bool(os.environ.get("GEMINI_API_KEY", "").strip())


def extract(statement_texts: dict[str, str], periods: list[str]) -> dict | None:
    """Run Gemini extraction over the detected statement-page text.

    `statement_texts` maps statement name -> concatenated page text (only the
    pages the deterministic detector already flagged). Returns a dict shaped like
      {statement: {"rows": {key:{period:val}}, "labels": {key:label}, "order": {key:int}}, ...}
      with a top-level "periods" list,
    ready to merge into ParsedPdf, or None if AI is unavailable / failed.
    """
    if not is_configured():
        return None
    body = "\n\n".join(f"=== {name} ===\n{text}" for name, text in statement_texts.items() if text).strip()
    if not body:
        return None

    try:
        from google import genai
        from google.genai import types

        client = genai.Client(api_key=os.environ["GEMINI_API_KEY"].strip())
        model_name = os.environ.get("GEMINI_MODEL", "gemini-2.5-flash").strip()
        prompt = _PROMPT.format(periods=", ".join(periods) or "none detected", body=body[:120_000])
        resp = client.models.generate_content(
            model=model_name,
            contents=prompt,
            config=types.GenerateContentConfig(
                temperature=0, response_mime_type="application/json"
            ),
        )
        raw = (resp.text or "").strip()
        data = _loads_lenient(raw)
        if not data:
            return None
        return _to_parsed_shape(data)
    except Exception as exc:  # noqa: BLE001 — fallback must never break a request
        print(f"[gemini_fallback] extraction unavailable: {exc}")
        return None


def _loads_lenient(raw: str) -> dict | None:
    raw = raw.strip()
    if raw.startswith("```"):  # strip ```json ... ``` fences if present
        raw = re.sub(r"^```[a-zA-Z]*\n?", "", raw)
        raw = re.sub(r"\n?```$", "", raw).strip()
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        match = re.search(r"\{.*\}", raw, flags=re.S)
        if match:
            try:
                return json.loads(match.group(0))
            except json.JSONDecodeError:
                return None
    return None


def _to_parsed_shape(data: dict) -> dict:
    periods = [str(p) for p in data.get("periods", []) if p]
    out: dict = {"periods": periods, "statements": {}}
    for stmt in STATEMENTS:
        items = data.get("statements", {}).get(stmt, []) or []
        rows: dict[str, dict[str, float]] = {}
        labels: dict[str, str] = {}
        order: dict[str, int] = {}
        for idx, item in enumerate(items):
            label = str(item.get("label", "")).strip()
            if not label:
                continue
            values_raw = item.get("values", {}) or {}
            values: dict[str, float] = {}
            for period, val in values_raw.items():
                num = _to_number(val)
                if num is not None:
                    values[str(period)] = num
            if not values:
                continue
            key = label.lower()
            rows[key] = values
            labels[key] = label
            order[key] = idx
        if rows:
            out["statements"][stmt] = {"rows": rows, "labels": labels, "order": order}
    return out


def _to_number(val) -> float | None:
    if isinstance(val, (int, float)):
        return float(val)
    if not isinstance(val, str):
        return None
    s = val.strip().replace(",", "").replace("$", "").replace("₹", "")
    if s in ("", "-", "—", "–"):
        return 0.0
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()")
    try:
        num = float(s)
    except ValueError:
        return None
    return -num if neg else num
