from fastapi import FastAPI, BackgroundTasks, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
import httpx
import json
import os
import re
import asyncio
from datetime import datetime, date, timedelta
from typing import Optional
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from groq import Groq as GroqClient
from pydantic import BaseModel
from dotenv import load_dotenv
import io

try:
    from pypdf import PdfReader
    PYPDF_AVAILABLE = True
except ImportError:
    try:
        from PyPDF2 import PdfReader
        PYPDF_AVAILABLE = True
    except ImportError:
        PYPDF_AVAILABLE = False
        print("WARNING: pypdf not installed. PDF extraction disabled. Run: pip install pypdf")

load_dotenv()

app = FastAPI(title="NSE/BSE Announcement Tracker")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

GROQ_API_KEY = os.getenv("GROQ_API_KEY", "")
OUTPUT_DIR = "./output"
os.makedirs(OUTPUT_DIR, exist_ok=True)

NSE_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Accept": "application/json",
    "Accept-Language": "en-US,en;q=0.9",
    "Referer": "https://www.nseindia.com/",
}

BSE_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Accept": "application/json",
    "Accept-Language": "en-US,en;q=0.9",
    "Referer": "https://www.bseindia.com/",
}


class FetchRequest(BaseModel):
    from_date: Optional[str] = None  # DD-MM-YYYY
    to_date: Optional[str] = None    # DD-MM-YYYY


# ─── HELPERS ─────────────────────────────────────────────────────────────────

def safe_str(val) -> str:
    """Safely convert any value to a plain string for Excel."""
    if val is None:
        return ""
    if isinstance(val, dict):
        for key in ("LTP", "CurrRate", "lastPrice", "price", "value"):
            if key in val:
                return str(val[key])
        return str(list(val.values())[0]) if val else ""
    if isinstance(val, (list, tuple)):
        return str(val[0]) if val else ""
    return str(val)


# ─── NSE FETCHER ─────────────────────────────────────────────────────────────

async def fetch_nse_announcements(from_date: str, to_date: str) -> list:
    """Fetch NSE announcements for the given date range (DD-MM-YYYY)."""
    announcements = []
    async with httpx.AsyncClient(timeout=30, follow_redirects=True) as client:
        try:
            # NSE requires a session cookie — warm up first
            await client.get("https://www.nseindia.com", headers=NSE_HEADERS)
            url = (
                f"https://www.nseindia.com/api/corporate-announcements"
                f"?index=equities&from_date={from_date}&to_date={to_date}"
            )
            resp = await client.get(url, headers=NSE_HEADERS)
            if resp.status_code == 200:
                data = resp.json()
                items = data if isinstance(data, list) else data.get("data", [])
                for item in items:
                    announcements.append({
                        "exchange": "NSE",
                        "company":  item.get("comp", item.get("symbol", "")),
                        "symbol":   item.get("symbol", ""),
                        "subject":  item.get("subject", item.get("desc", "")),
                        "date":     item.get("an_dt", item.get("date", "")),
                        "link": (
                            f"https://www.nseindia.com/api/corporate-announcements"
                            f"?symbol={item.get('symbol','')}&an_num={item.get('an_num','')}"
                        ),
                        "raw": item,
                    })
            else:
                print(f"NSE fetch returned HTTP {resp.status_code}")
        except Exception as e:
            print(f"NSE fetch error: {e}")

    print(f"[NSE] Fetched {len(announcements)} announcements")
    return announcements


# ─── BSE FETCHER ─────────────────────────────────────────────────────────────

async def fetch_bse_announcements(from_date: str, to_date: str) -> list:
    """Fetch BSE announcements day-by-day (BSE paginates poorly over ranges)."""
    announcements = []
    try:
        start_dt = datetime.strptime(from_date, "%d-%m-%Y")
        end_dt   = datetime.strptime(to_date,   "%d-%m-%Y")
    except Exception:
        print("BSE: bad date format")
        return []

    async with httpx.AsyncClient(timeout=30, follow_redirects=True) as client:
        current_dt = start_dt
        while current_dt <= end_dt:
            d_str = current_dt.strftime("%Y%m%d")
            pageno = 1
            while True:
                try:
                    url = (
                        f"https://api.bseindia.com/BseIndiaAPI/api/AnnSubCategoryGetData/w"
                        f"?pageno={pageno}&strCat=-1&strPrevDate={d_str}&strScrip=&strSearch=P"
                        f"&strToDate={d_str}&strType=C&subcategory=-1"
                    )
                    resp = await client.get(url, headers=BSE_HEADERS)
                    if resp.status_code == 200:
                        data  = resp.json()
                        items = data.get("Table", data if isinstance(data, list) else [])

                        # ── DEBUG: log raw keys once so we know the exact field names ──
                        if items and current_dt == start_dt and pageno == 1:
                            print(f"[BSE DEBUG] Raw keys: {list(items[0].keys())}")

                        if not items:
                            break

                        for item in items:
                            scrip = item.get("SCRIP_CD", "")
                            attach = item.get("ATTACHMENTNAME", "")
                            announcements.append({
                                "exchange": "BSE",
                                "company":  item.get("SLONGNAME", item.get("short_name", "")),
                                "symbol":   str(scrip),
                                "subject":  item.get("HEADLINE", item.get("subject", "")),
                                "date":     item.get("NEWS_DT", item.get("date", "")),
                                "link": (
                                    f"https://www.bseindia.com/xml-data/corpfiling/AttachHis/{attach}"
                                    if attach else ""
                                ),
                                "raw": item,
                            })
                            
                        if len(items) < 50:
                            break
                        pageno += 1
                    else:
                        print(f"[BSE] HTTP {resp.status_code} for date {d_str} page {pageno}")
                        break
                except Exception as e:
                    print(f"[BSE] Fetch error for {d_str} page {pageno}: {e}")
                    break

            current_dt += timedelta(days=1)

    print(f"[BSE] Fetched {len(announcements)} announcements")
    return announcements


# ─── STOCK DATA ───────────────────────────────────────────────────────────────

async def fetch_stock_data(symbol: str, exchange: str) -> dict:
    """Fetch CMP, MCap, Sector for a given symbol."""
    result = {"cmp": "", "mcap": "", "sector": ""}
    if not symbol:
        return result

    async with httpx.AsyncClient(timeout=15, follow_redirects=True) as client:
        try:
            if exchange == "NSE":
                await client.get("https://www.nseindia.com", headers=NSE_HEADERS)
                url  = f"https://www.nseindia.com/api/quote-equity?symbol={symbol.upper()}"
                resp = await client.get(url, headers=NSE_HEADERS)
                if resp.status_code == 200:
                    d          = resp.json()
                    price_info = d.get("priceInfo", {})
                    meta       = d.get("metadata", {})
                    mc_raw     = d.get("securityInfo", {}).get("marketCap", "")
                    try:
                        mcap_str = f"{round(float(safe_str(mc_raw)) / 1e7, 2)} Cr" if mc_raw else ""
                    except Exception:
                        mcap_str = safe_str(mc_raw)
                    result = {
                        "cmp":    safe_str(price_info.get("lastPrice", "")),
                        "mcap":   mcap_str,
                        "sector": safe_str(meta.get("industry", "")),
                    }

            elif exchange == "BSE":
                url  = (
                    f"https://api.bseindia.com/BseIndiaAPI/api/getScripHeaderData/w"
                    f"?Debtflag=&scripcode={symbol}&seriesid="
                )
                resp = await client.get(url, headers=BSE_HEADERS)
                if resp.status_code == 200:
                    d = resp.json()
                    result = {
                        "cmp":    safe_str(d.get("CurrRate", d.get("Curt_Rate", d.get("LTP", "")))),
                        "mcap":   safe_str(d.get("Mktcap",  d.get("MktCap", ""))),
                        "sector": safe_str(d.get("Indust",  d.get("Industry", ""))),
                    }
        except Exception as e:
            print(f"[Stock] Fetch error for {symbol} ({exchange}): {e}")

    return result


# ─── EXCHANGE CATEGORY CHECK ──────────────────────────────────────────────────

# BSE subcategory strings that definitively indicate an auth capital announcement.
# We use substring matching (not equality) to handle minor API variations.
BSE_AUTH_SUBCAT_KEYWORDS = [
    "alteration of capital",
    "increase in authorised capital",
    "increase in authorized capital",
    "sub-division of shares",
    "subdivision of shares",
    "capital change",
]

# NSE sub_type / desc keywords
NSE_AUTH_SUBTYPE_KEYWORDS = [
    "capital change",
    "alteration of capital",
    "increase in authorised",
    "increase in authorized",
]


def is_exchange_auth_category(ann: dict) -> bool:
    """
    Return True if the exchange itself explicitly tagged this announcement
    as being about an authorised capital change.

    BSE: checks SUBCATEGORYNAME and SUBCATNAME (both field names observed in the wild).
    NSE: checks sub_type and desc fields from the raw item.
    """
    raw      = ann.get("raw", {})
    exchange = ann.get("exchange", "")

    if exchange == "BSE":
        # Try every field name BSE has ever used for this
        subcat_raw = (
            raw.get("SUBCATEGORYNAME")
            or raw.get("SUBCATNAME")
            or raw.get("SubCategoryName")
            or raw.get("subcategoryname")
            or ""
        )
        subcat = str(subcat_raw).strip().lower()

        print(f"[BSE CAT] company='{ann.get('company','')}' subcat='{subcat}'")

        return any(kw in subcat for kw in BSE_AUTH_SUBCAT_KEYWORDS)

    elif exchange == "NSE":
        sub_type = str(raw.get("sub_type", "") or "").strip().lower()
        desc     = str(raw.get("desc",     "") or "").strip().lower()
        subject  = str(raw.get("subject",  "") or "").strip().lower()

        combined = f"{sub_type} {desc} {subject}"
        return any(kw in combined for kw in NSE_AUTH_SUBTYPE_KEYWORDS)

    return False


# ─── KEYWORD FILTER ───────────────────────────────────────────────────────────

AUTH_CAPITAL_KEYWORDS = [
    "authorised capital",
    "authorized capital",
    "increase in authorised",
    "increase in authorized",
    "alteration of capital clause",
    "increase of capital clause",
    "memorandum of association",
    "moa alteration",
    "increase in auth",
    "raising of authorised",
    "change in authorised capital",
    "change in authorized capital",
    "alteration in capital clause",
    "sub-division",
    "subdivision",
    "increase in the authorised",
    "increase in the authorized",
]

AUTH_CAPITAL_BLACKLIST = [
    "outcome of board meeting",
    "outcome of the board meeting",
    "board meeting outcome",
    "financial results",
    "dividend",
    "buyback",
    "merger",
    "amalgamation",
    "acquisition",
    "change in management",
    "resignation",
    "appointment",
    "loss of share certificate",
    "trading window",
    "compliances",
    "credit rating",
    "press release",
    "investor presentation",
    "analyst meet",
    "agm",
    "egm",
    "annual general meeting",
    "extraordinary general meeting",
    "allotment of shares",
    "issue of shares",
    "debentures",
    "rights issue",
    "preferential allotment",
    "esop",
    "stock split",
    "bonus issue",
]


def is_auth_capital_keyword_match(subject: str) -> bool:
    """Return True only if the subject strongly suggests an auth capital increase."""
    subj = subject.lower().strip()
    if any(bl in subj for bl in AUTH_CAPITAL_BLACKLIST):
        return False
    return any(kw in subj for kw in AUTH_CAPITAL_KEYWORDS)


# ─── AI CLASSIFIER ───────────────────────────────────────────────────────────

def classify_with_ai(announcements: list) -> list:
    """
    Classify announcements as auth capital or not.

    Pipeline:
      1. Exchange category check  → definitive YES from the exchange itself
      2. Keyword pre-filter       → must have auth capital language in subject
      3. Groq AI confirmation     → verifies and extracts fields for candidates
      4. Final gate               → exchange match overrides AI rejection
    """

    # ── Step 1: Pre-classify ──────────────────────────────────────────────────
    for ann in announcements:
        subj = ann.get("subject", "")
        ann["_keyword_match"]  = is_auth_capital_keyword_match(subj)
        ann["_exchange_match"] = is_exchange_auth_category(ann)

        ann.setdefault("is_auth_capital",   False)
        ann.setdefault("board_approval",    "")
        ann.setdefault("dobm",              "")
        ann.setdefault("existing_auth_cap", "")
        ann.setdefault("new_auth_cap",      "")
        ann.setdefault("proposed_increase", "")
        ann.setdefault("remark_positive",   "")
        ann.setdefault("remark_negative",   "")
        ann.setdefault("action",            "NEUTRAL")

    # ── Step 2: Split into candidates vs non-candidates ───────────────────────
    candidates     = [a for a in announcements if a["_exchange_match"] or a["_keyword_match"]]
    non_candidates = [a for a in announcements if not (a["_exchange_match"] or a["_keyword_match"])]

    print(
        f"[Classify] Exchange-matched: {sum(1 for a in announcements if a['_exchange_match'])} | "
        f"Keyword-matched: {sum(1 for a in announcements if a['_keyword_match'])} | "
        f"Total candidates: {len(candidates)} | Skipping: {len(non_candidates)}"
    )

    for ann in non_candidates:
        ann["is_auth_capital"] = False

    if not candidates:
        return announcements

    # ── Step 3: Fallback if no Groq key ──────────────────────────────────────
    if not GROQ_API_KEY:
        print("[Classify] No GROQ_API_KEY — using keyword/exchange match only")
        for ann in candidates:
            ann["is_auth_capital"] = ann["_exchange_match"] or ann["_keyword_match"]
            ann["action"] = "WATCH" if ann["is_auth_capital"] else "NEUTRAL"
        return announcements

    # ── Step 4: Groq AI confirmation in batches ───────────────────────────────
    client = GroqClient(api_key=GROQ_API_KEY)

    SYSTEM_PROMPT = (
        "You are a senior Indian equity market analyst specializing in NSE/BSE corporate filings. "
        "You receive announcements pre-filtered as likely about Increase in Authorised/Authorized Equity Capital. "
        "Confirm if they truly are and extract structured details. "
        "Respond ONLY with a valid JSON array — no explanation, no markdown fences, no extra text."
    )

    batch_size = 10
    for i in range(0, len(candidates), batch_size):
        batch = candidates[i : i + batch_size]

        batch_text = json.dumps(
            [
                {"idx": j, "company": a.get("company", ""), "subject": a.get("subject", "")}
                for j, a in enumerate(batch)
            ],
            indent=2,
        )

        user_prompt = f"""These announcements have been pre-filtered as LIKELY about Increase in Authorised Equity Capital.
CONFIRM each one and extract details.

CRITICAL RULE: is_auth_capital = true ONLY if the announcement is SPECIFICALLY about:
- Increase in Authorised/Authorized Share Capital
- Alteration of Capital Clause in MOA
- Raising the authorized equity capital limit
DO NOT mark true for: Board Meeting outcomes, financial results, dividends, share allotments, or any other type.

For EACH announcement extract:
1. is_auth_capital (bool)
2. board_approval: Board approval date DD-MM-YYYY or ""
3. dobm: Date of Board Meeting DD-MM-YYYY or ""
4. existing_auth_cap: e.g. "10,00,00,000" or ""
5. new_auth_cap: e.g. "20,00,00,000" or ""
6. proposed_increase: new minus existing or ""
7. remark_positive: 1-line positive remark if is_auth_capital=true else ""
8. remark_negative: 1-line risk remark if is_auth_capital=true else ""
9. action: "BUY on dip"/"ACCUMULATE"/"WATCH"/"NEUTRAL"/"AVOID" — only if is_auth_capital=true else "NEUTRAL"

Announcements:
{batch_text}

Reply ONLY with JSON array:
[{{"idx":0,"is_auth_capital":true,"board_approval":"","dobm":"","existing_auth_cap":"","new_auth_cap":"","proposed_increase":"","remark_positive":"","remark_negative":"","action":""}}]"""

        try:
            response = client.chat.completions.create(
                model="llama-3.3-70b-versatile",
                messages=[
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user",   "content": user_prompt},
                ],
                temperature=0.0,
                max_tokens=3000,
            )
            raw_text = response.choices[0].message.content.strip()
            raw_text = re.sub(r"^```json\s*", "", raw_text)
            raw_text = re.sub(r"^```\s*",     "", raw_text)
            raw_text = re.sub(r"\s*```$",     "", raw_text)

            results = json.loads(raw_text)
            for r in results:
                idx = r.get("idx", 0)
                if idx >= len(batch):
                    continue
                ann = batch[idx]

                ai_says_auth = bool(r.get("is_auth_capital", False))

                # ── Final gate ────────────────────────────────────────────────
                # Exchange match = definitive YES (exchange is source of truth).
                # Without exchange match, BOTH AI and keyword must agree.
                if ann["_exchange_match"]:
                    final_is_auth = True
                else:
                    final_is_auth = ai_says_auth and ann["_keyword_match"]

                print(
                    f"[Classify] '{ann.get('company','')}' | exchange={ann['_exchange_match']} "
                    f"keyword={ann['_keyword_match']} ai={ai_says_auth} → FINAL={final_is_auth}"
                )

                ann.update({
                    "is_auth_capital":   final_is_auth,
                    "board_approval":    r.get("board_approval",    ""),
                    "dobm":              r.get("dobm",              ""),
                    "existing_auth_cap": r.get("existing_auth_cap", ""),
                    "new_auth_cap":      r.get("new_auth_cap",      ""),
                    "proposed_increase": r.get("proposed_increase", ""),
                    "remark_positive":   r.get("remark_positive",   "") if final_is_auth else "",
                    "remark_negative":   r.get("remark_negative",   "") if final_is_auth else "",
                    "action":            r.get("action", "NEUTRAL")     if final_is_auth else "NEUTRAL",
                })

        except Exception as e:
            print(f"[Classify] Groq error (batch {i}): {e}")
            # Fallback: trust exchange/keyword match
            for ann in batch:
                ann["is_auth_capital"] = ann["_exchange_match"] or ann["_keyword_match"]
                ann.setdefault("action", "WATCH" if ann["is_auth_capital"] else "NEUTRAL")

    return announcements


# ─── PDF EXTRACTION ──────────────────────────────────────────────────────────

async def download_pdf_text(url: str, exchange: str) -> str:
    """Download a PDF and return up to 8 000 chars of extracted text."""
    if not url or url.strip() in ("", "—"):
        return ""

    headers = NSE_HEADERS if exchange == "NSE" else BSE_HEADERS

    try:
        async with httpx.AsyncClient(timeout=30, follow_redirects=True) as client:
            if exchange == "NSE":
                await client.get("https://www.nseindia.com", headers=headers)

            resp = await client.get(url, headers=headers)
            if resp.status_code != 200:
                print(f"[PDF] Download failed (HTTP {resp.status_code}): {url}")
                return ""

            content_type = resp.headers.get("content-type", "")
            if "pdf" not in content_type.lower() and not url.lower().endswith(".pdf"):
                print(f"[PDF] Not a PDF ({content_type}): {url}")
                return ""

            if not PYPDF_AVAILABLE:
                print("[PDF] pypdf not available — skipping")
                return ""

            reader     = PdfReader(io.BytesIO(resp.content))
            pages_text = []
            for page in reader.pages:
                try:
                    pages_text.append(page.extract_text() or "")
                except Exception:
                    pass

            full_text = "\n".join(pages_text).strip()
            print(f"[PDF] Extracted {len(full_text)} chars from {url}")
            return full_text[:8000]

    except Exception as e:
        print(f"[PDF] Extraction error for {url}: {e}")
        return ""


def extract_fields_from_pdf_text(pdf_text: str, company: str, subject: str) -> dict:
    """Use Groq to extract auth capital fields from the full PDF text."""
    empty = {
        "board_approval": "", "dobm": "",
        "existing_auth_cap": "", "new_auth_cap": "",
        "proposed_increase": "", "remark_positive": "",
        "remark_negative": "", "action": "WATCH",
    }
    if not pdf_text or not GROQ_API_KEY:
        return empty

    client = GroqClient(api_key=GROQ_API_KEY)

    prompt = f"""You are a financial analyst reading an NSE/BSE corporate announcement PDF.
Company: {company}
Subject: {subject}

PDF Content:
{pdf_text}

Extract EXACTLY these fields. Return ONLY a JSON object (no explanation, no markdown):
{{
  "board_approval": "DD-MM-YYYY or empty string",
  "dobm": "DD-MM-YYYY or empty string",
  "existing_auth_cap": "e.g. 10,00,00,000 or empty string",
  "new_auth_cap": "e.g. 20,00,00,000 or empty string",
  "proposed_increase": "new minus existing or empty string",
  "remark_positive": "1-line positive remark",
  "remark_negative": "1-line risk remark",
  "action": "BUY on dip / ACCUMULATE / WATCH / NEUTRAL / AVOID"
}}"""

    try:
        response = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.0,
            max_tokens=1000,
        )
        raw = response.choices[0].message.content.strip()
        raw = re.sub(r"^```json\s*", "", raw)
        raw = re.sub(r"^```\s*",     "", raw)
        raw = re.sub(r"\s*```$",     "", raw)
        result = json.loads(raw)
        for k in empty:
            result.setdefault(k, "")
        return result
    except Exception as e:
        print(f"[PDF AI] Field extraction error: {e}")
        return empty


async def enrich_auth_announcements_from_pdf(auth_anns: list) -> list:
    """Download PDFs and extract fields for all auth capital announcements."""
    if not auth_anns:
        return auth_anns

    sem = asyncio.Semaphore(3)

    async def process_one(ann: dict) -> dict:
        async with sem:
            url      = ann.get("link", "")
            exchange = ann.get("exchange", "NSE")
            company  = ann.get("company", "")
            subject  = ann.get("subject", "")

            print(f"[PDF] Processing {company}: {url}")
            pdf_text = await download_pdf_text(url, exchange)

            if pdf_text:
                loop   = asyncio.get_running_loop()
                fields = await loop.run_in_executor(
                    None, extract_fields_from_pdf_text, pdf_text, company, subject
                )
                ann["board_approval"]    = fields.get("board_approval",    ann.get("board_approval", ""))
                ann["dobm"]              = fields.get("dobm",              ann.get("dobm", ""))
                ann["existing_auth_cap"] = fields.get("existing_auth_cap", ann.get("existing_auth_cap", ""))
                ann["new_auth_cap"]      = fields.get("new_auth_cap",      ann.get("new_auth_cap", ""))
                ann["proposed_increase"] = fields.get("proposed_increase", ann.get("proposed_increase", ""))
                ann["remark_positive"]   = fields.get("remark_positive",   ann.get("remark_positive", ""))
                ann["remark_negative"]   = fields.get("remark_negative",   ann.get("remark_negative", ""))
                ann["action"]            = fields.get("action",            ann.get("action", "WATCH"))
                ann["_pdf_extracted"]    = True
            else:
                ann["_pdf_extracted"] = False
                print(f"[PDF] No text extracted for {company}")

            return ann

    results = await asyncio.gather(*[process_one(a) for a in auth_anns], return_exceptions=True)
    final   = []
    for i, r in enumerate(results):
        if isinstance(r, Exception):
            print(f"[PDF] Enrichment error for announcement {i}: {r}")
            final.append(auth_anns[i])
        else:
            final.append(r)
    return final


# ─── EXCEL GENERATOR ─────────────────────────────────────────────────────────

COLUMNS = [
    ("Sr.no",                  8),
    ("Date of Entry",         13),
    ("Name of the Company",   28),
    ("Board Approval",        14),
    ("DOBM",                  12),
    ("Exst Auth Eq Cap (INR)",18),
    ("New Auth Eq Cap (INR)", 18),
    ("Proposed Increase (INR)",18),
    ("CMP",                    9),
    ("M Cap (in Cr)",         13),
    ("Sector",                18),
    ("Remark Positive",       28),
    ("Remark Negative",       28),
    ("Action",                14),
    ("Link",                  35),
]


def style_header_cell(cell, bg_color: str, font_color: str = "FFFFFF"):
    cell.font      = Font(bold=True, color=font_color, size=10, name="Calibri")
    cell.fill      = PatternFill("solid", fgColor=bg_color)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin           = Side(style="thin", color="CCCCCC")
    cell.border    = Border(left=thin, right=thin, top=thin, bottom=thin)


def style_data_cell(cell, wrap=False):
    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=wrap)
    thin           = Side(style="thin", color="E0E0E0")
    cell.border    = Border(left=thin, right=thin, top=thin, bottom=thin)
    cell.font      = Font(size=9, name="Calibri")


def write_sheet(ws, rows: list, header_color: str, title: str):
    # Title row
    ws.merge_cells(f"A1:{get_column_letter(len(COLUMNS))}1")
    tc = ws["A1"]
    tc.value     = title
    tc.font      = Font(bold=True, size=13, color="FFFFFF", name="Calibri")
    tc.fill      = PatternFill("solid", fgColor="1A1A2E")
    tc.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    # Header row
    for col_idx, (col_name, col_width) in enumerate(COLUMNS, 1):
        cell = ws.cell(row=2, column=col_idx, value=col_name)
        style_header_cell(cell, header_color)
        ws.column_dimensions[get_column_letter(col_idx)].width = col_width
    ws.row_dimensions[2].height = 35

    # Data rows
    for row_idx, ann in enumerate(rows, 3):
        ws.row_dimensions[row_idx].height = 55
        stock = ann.get("_stock", {})
        if not isinstance(stock, dict):
            stock = {}

        values = [
            row_idx - 2,
            safe_str(ann.get("date", "")),
            safe_str(ann.get("company", "")),
            safe_str(ann.get("board_approval", "")),
            safe_str(ann.get("dobm", "")),
            safe_str(ann.get("existing_auth_cap", "")),
            safe_str(ann.get("new_auth_cap", "")),
            safe_str(ann.get("proposed_increase", "")),
            safe_str(stock.get("cmp", "")),
            safe_str(stock.get("mcap", "")),
            safe_str(stock.get("sector", ann.get("sector", ""))),
            safe_str(ann.get("remark_positive", "")),
            safe_str(ann.get("remark_negative", "")),
            safe_str(ann.get("action", "")),
            safe_str(ann.get("link", "")),
        ]

        for col_idx, value in enumerate(values, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            style_data_cell(cell, wrap=(col_idx in [3, 12, 13, 15]))
            if row_idx % 2 == 0:
                cell.fill = PatternFill("solid", fgColor="F8F9FF")
            if col_idx == 14 and value:
                v      = str(value).upper()
                colors = {
                    "BUY":        "C6EFCE",
                    "ACCUMULATE": "C6EFCE",
                    "AVOID":      "FFC7CE",
                    "NEUTRAL":    "FFEB9C",
                    "WATCH":      "FFEB9C",
                }
                for key, clr in colors.items():
                    if key in v:
                        cell.fill = PatternFill("solid", fgColor=clr)
                        break

    ws.freeze_panes = "A3"


def generate_excel(auth_anns: list, other_anns: list, date_str: str) -> str:
    wb = openpyxl.Workbook()

    ws_auth       = wb.active
    ws_auth.title = "Auth Capital"
    write_sheet(ws_auth, auth_anns, "1565C8",
                f"Authorised Capital Increase Announcements — {date_str}")

    ws_other       = wb.create_sheet("Other Announcements")
    write_sheet(ws_other, other_anns, "C85215",
                f"Other Announcements — {date_str}")

    filename = f"announcements_{date_str.replace(' ', '_').replace('-', '')}.xlsx"
    filepath = os.path.join(OUTPUT_DIR, filename)
    wb.save(filepath)
    print(f"[Excel] Saved: {filename}")
    return filename


# ─── PIPELINE STATUS ──────────────────────────────────────────────────────────

fetch_status: dict = {"status": "idle", "message": "", "progress": 0, "filename": ""}


# ─── ENDPOINTS ───────────────────────────────────────────────────────────────

@app.get("/status")
async def get_status():
    return fetch_status


@app.post("/fetch")
async def trigger_fetch(req: FetchRequest, background_tasks: BackgroundTasks):
    if fetch_status["status"] == "running":
        raise HTTPException(status_code=409, detail="Fetch already in progress")
    today     = date.today().strftime("%d-%m-%Y")
    from_date = req.from_date or today
    to_date   = req.to_date   or today
    background_tasks.add_task(run_full_pipeline, from_date, to_date)
    return {"message": "Fetch started", "from_date": from_date, "to_date": to_date}


async def run_full_pipeline(from_date: str, to_date: str):
    global fetch_status
    try:
        fetch_status = {
            "status": "running", "message": "Fetching NSE announcements...",
            "progress": 10, "filename": "",
        }

        nse_anns = await fetch_nse_announcements(from_date, to_date)
        fetch_status = {
            **fetch_status,
            "message":  f"NSE done ({len(nse_anns)} announcements). Fetching BSE...",
            "progress": 25,
        }

        bse_anns = await fetch_bse_announcements(from_date, to_date)
        fetch_status = {
            **fetch_status,
            "message":  f"BSE done ({len(bse_anns)} announcements). Running AI classification...",
            "progress": 45,
        }

        all_anns = nse_anns + bse_anns
        print(f"[Pipeline] Total: {len(all_anns)} (NSE: {len(nse_anns)}, BSE: {len(bse_anns)})")

        if not all_anns:
            fetch_status = {
                "status": "done",
                "message": "No announcements found for this date range. Try a different date.",
                "progress": 100, "filename": "",
                "auth_count": 0, "other_count": 0, "total": 0, "announcements": [],
            }
            return

        loop     = asyncio.get_running_loop()
        all_anns = await loop.run_in_executor(None, classify_with_ai, all_anns)

        auth_anns  = [a for a in all_anns if a.get("is_auth_capital")]
        other_anns = [a for a in all_anns if not a.get("is_auth_capital")]
        print(f"[Pipeline] Auth capital: {len(auth_anns)}, Other: {len(other_anns)}")

        if auth_anns:
            fetch_status = {
                **fetch_status,
                "message":  f"Classified. Reading PDFs for {len(auth_anns)} auth capital announcements...",
                "progress": 55,
            }
            auth_anns = await enrich_auth_announcements_from_pdf(auth_anns)

        fetch_status = {**fetch_status, "message": "PDF extraction done. Fetching stock prices...", "progress": 70}

        all_anns = auth_anns + other_anns

        sem = asyncio.Semaphore(5)

        async def fetch_with_sem(symbol, exchange):
            async with sem:
                return await fetch_stock_data(symbol, exchange)

        stock_results = await asyncio.gather(
            *[fetch_with_sem(a["symbol"], a["exchange"]) for a in all_anns],
            return_exceptions=True,
        )
        for ann, stock in zip(all_anns, stock_results):
            ann["_stock"] = stock if isinstance(stock, dict) else {}

        fetch_status = {**fetch_status, "message": "Generating Excel file...", "progress": 88}

        date_label = f"{from_date} to {to_date}" if from_date != to_date else from_date
        filename   = await loop.run_in_executor(None, generate_excel, auth_anns, other_anns, date_label)

        fetch_status = {
            "status":        "done",
            "message":       f"Done! {len(auth_anns)} auth capital + {len(other_anns)} other announcements.",
            "progress":      100,
            "filename":      filename,
            "auth_count":    len(auth_anns),
            "other_count":   len(other_anns),
            "total":         len(all_anns),
            "announcements": all_anns,
        }

    except Exception as e:
        import traceback
        traceback.print_exc()
        fetch_status = {
            "status": "error", "message": f"Error: {str(e)}",
            "progress": 0, "filename": "",
        }


@app.get("/download/{filename}")
async def download_file(filename: str):
    filepath = os.path.join(OUTPUT_DIR, filename)
    if not os.path.exists(filepath):
        raise HTTPException(status_code=404, detail="File not found")
    return FileResponse(
        filepath,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=filename,
    )


@app.get("/announcements")
async def get_announcements():
    anns = fetch_status.get("announcements", [])
    return {"data": anns, "total": len(anns)}


@app.get("/health")
async def health():
    return {"status": "ok", "api_key_set": bool(GROQ_API_KEY)}


# ─── FRONTEND ─────────────────────────────────────────────────────────────────

from fastapi.staticfiles import StaticFiles

frontend_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "frontend")
if os.path.exists(frontend_path):
    @app.get("/")
    async def serve_frontend():
        index_path = os.path.join(frontend_path, "index.html")
        if os.path.exists(index_path):
            return FileResponse(index_path)
        raise HTTPException(status_code=404, detail="Frontend index.html not found")