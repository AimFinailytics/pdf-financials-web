# AimFiinsight · Annual Report → Excel

Free, branded web tool: upload annual/quarterly report PDFs → download clean,
sectioned, multi-year consolidated Excel workbooks (Income Statement, Balance
Sheet, Cash Flow). Free for everyone after a one-time email signup; optional
donation QR for support.

## Run locally

```powershell
python -m pip install -r requirements.txt
python app.py
```

Open `http://127.0.0.1:5050`. Everything works with **no configuration** — email
signups fall back to a local `signups.csv`, and AI assist stays off.

## Output

- One Excel workbook per company; a master workbook when 2+ companies are uploaded.
- Each workbook has three sheets: `Income Statement`, `Balance Sheet`, `Cash Flow Statement`.
- Multi-year reports stack every period side-by-side.

## Commercial features

| Feature | How it works |
| --- | --- |
| **One-time email gate** | First use prompts for an email (`/api/signup`). Stored, then unlimited free use. Session cookie remembers the visitor. |
| **Email list** | Signups append to a **Google Sheet** (your mailing list). Falls back to local `signups.csv` if Sheets isn't configured. |
| **Donation QR** | Optional "support us" panel. Drop your UPI QR at `static/img/donate-qr.png`. |
| **Privacy** | Default conversion is 100% on-server — nothing leaves. Advertised as a trust badge. |
| **Opt-in AI assist** | A checkbox sends only statement-page **text** to the free Gemini Flash tier, used only when the on-server parser has low confidence. |

## Configuration (all optional)

Copy `.env.example` → `.env` (local) or set these as environment variables on Render:

- `FLASK_SECRET_KEY` — stable random string so the signup cookie survives restarts.
- `SHEET_ID` + `GOOGLE_SERVICE_ACCOUNT_JSON` — Google Sheets email storage
  (enable Sheets API → service account → share the sheet with its client_email).
- `GEMINI_API_KEY` (+ optional `GEMINI_MODEL`) — free key from
  https://aistudio.google.com/apikey for opt-in AI assist.

See `.env.example` for step-by-step setup notes.

## Brand assets

- `static/img/aimfiinsight-logo.png` — the emblem (header + signup modal + favicon).
- `static/img/donate-qr.png` — **add your UPI donation QR here** (a graceful
  placeholder shows until it exists).

## Deploy (Render)

`Procfile` runs gunicorn. Add the env vars above in the Render dashboard. The
free web-service tier works (cold-starts after idle); `/health` is the uptime ping.
