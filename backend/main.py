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
from fastapi.staticfiles import StaticFiles

try:
    from pypdf import PdfReader
    PYPDF_AVAILABLE = True
except ImportError:
    try:
        from PyPDF2 import PdfReader
        PYPDF_AVAILABLE = True
    except ImportError:
        PYPDF_AVAILABLE = False
        print("WARNING: pypdf not installed. Run: pip install pypdf")

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
    from_date: Optional[str] = None
    to_date: Optional[str] = None


# ══════════════════════════════════════════════════════════════════════════════
#  NSE FETCHER — FIXED PDF LINK
# ══════════════════════════════════════════════════════════════════════════════

async def fetch_nse_announcements(from_date: str, to_date: str) -> list:
    announcements = []
    async with httpx.AsyncClient(timeout=30, follow_redirects=True) as client:
        try:
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
                    attchmnt_text = str(
                        item.get("attchmnt_text", "")
                        or item.get("attachment", "")
                        or item.get("attchmntText", "")
                        or item.get("ATTACHMENT", "")
                    ).strip()

                    symbol = item.get("symbol", "")
                    an_num = item.get("an_num", "")

                    pdf_link = item.get("attchmntFile") or ""
                    if pdf_link and not pdf_link.startswith("http"):
                        if pdf_link.startswith("/"):
                            pdf_link = f"https://nsearchives.nseindia.com{pdf_link}"
                        else:
                            pdf_link = f"https://nsearchives.nseindia.com/corporate/{pdf_link}"

                    web_link = (
                        f"https://www.nseindia.com/companies/corporate-announcements"
                        f"?symbol={symbol}&id={an_num}"
                    )

                    announcements.append({
                        "exchange": "NSE",
                        "company": item.get("comp", item.get("symbol", "")),
                        "symbol": symbol,
                        "subject": item.get("subject", item.get("desc", "")),
                        "date": item.get("an_dt", item.get("date", "")),
                        "link": pdf_link,
                        "web_link": web_link,
                        "attchmnt_text": attchmnt_text,
                        "nse_sm_name": str(item.get("sm_name", "")),
                        "nse_sub_type": str(item.get("sub_type", "")),
                        "nse_an_type": str(item.get("an_type", "")),
                        "nse_sm_user_name": str(item.get("sm_user_name", "")),
                        "nse_desc": str(item.get("desc", "")),
                        "nse_ex_category": str(item.get("ex_category", "")),
                        "raw": item,
                    })

                print(f"  NSE: {len(announcements)} announcements fetched")
                
                if items:
                    print("\n── RAW NSE ITEM PREVIEW ──")
                    print(json.dumps(items[0], indent=2))
                    print("──────────────────────────\n")

                for a in announcements[:5]:
                    print(
                        f"    Sample: [{a['exchange']}] {a.get('company','')[:30]} "
                        f"| subject='{a.get('subject','')[:80]}' "
                        f"| attchmnt='{a.get('attchmnt_text','')[:40]}'"
                    )

        except Exception as e:
            print(f"NSE fetch error: {e}")
    return announcements


# ══════════════════════════════════════════════════════════════════════════════
#  BSE FETCHER — IMPROVED FIELD CAPTURE
# ══════════════════════════════════════════════════════════════════════════════

async def fetch_bse_announcements(from_date: str, to_date: str) -> list:
    announcements = []
    try:
        start_dt = datetime.strptime(from_date, "%d-%m-%Y")
        end_dt   = datetime.strptime(to_date,   "%d-%m-%Y")
    except Exception:
        return []

    async with httpx.AsyncClient(timeout=30, follow_redirects=True) as client:
        current_dt = start_dt
        while current_dt <= end_dt:
            d_str = current_dt.strftime("%Y%m%d")
            try:
                url = (
                    f"https://api.bseindia.com/BseIndiaAPI/api/AnnSubCategoryGetData/w"
                    f"?strCat=-1&strPrevDate={d_str}&strScrip=&strSearch=P"
                    f"&strToDate={d_str}&strType=C&subcategory=-1"
                )
                resp = await client.get(url, headers=BSE_HEADERS)
                if resp.status_code == 200:
                    data = resp.json()
                    items = data.get("Table", data if isinstance(data, list) else [])
                    for item in items:
                        scrip = item.get("SCRIP_CD", "")
                        attachment = str(
                            item.get("ATTACHMENTNAME", "")
                            or item.get("ATTACHMENT", "")
                            or item.get("FILENAME", "")
                        ).strip()

                        if attachment:
                            pdf_link = f"https://www.bseindia.com/xml-data/corpfiling/AttachLive/{attachment}"
                        else:
                            pdf_link = ""

                        announcements.append({
                            "exchange": "BSE",
                            "company": item.get("SLONGNAME", item.get("short_name", "")),
                            "symbol": str(scrip),
                            "subject": item.get("HEADLINE", item.get("subject", "")),
                            "date": item.get("NEWS_DT", item.get("date", "")),
                            "link": pdf_link,
                            "bse_attachment": attachment,
                            "bse_category": str(item.get("CATEGORYNAME", "")),
                            "bse_subcategory": str(item.get("SUBCATEGORYNAME", "")),
                            "bse_subcatname": str(item.get("SUBCATNAME", "")),
                            "raw": item,
                        })
            except Exception as e:
                print(f"BSE fetch error for {d_str}: {e}")
            current_dt += timedelta(days=1)

    print(f"  BSE: {len(announcements)} announcements fetched")
    for a in announcements[:5]:
        raw = a.get("raw", {})
        print(
            f"    Sample: [{a['exchange']}] {a.get('company','')[:30]} "
            f"| subject='{a.get('subject','')[:80]}' "
            f"| CAT='{raw.get('CATEGORYNAME','')[:30]}' "
            f"| SUBCAT='{raw.get('SUBCATEGORYNAME','')[:30]}' "
            f"| attachment='{a.get('bse_attachment','')[:40]}'"
        )

    return announcements


# ══════════════════════════════════════════════════════════════════════════════
#  STOCK DATA
# ══════════════════════════════════════════════════════════════════════════════

def safe_str(val) -> str:
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


async def fetch_stock_data(symbol: str, exchange: str) -> dict:
    result = {"cmp": "", "mcap": "", "sector": ""}
    if not symbol:
        return result
    async with httpx.AsyncClient(timeout=15, follow_redirects=True) as client:
        try:
            if exchange == "NSE":
                await client.get("https://www.nseindia.com", headers=NSE_HEADERS)
                resp = await client.get(
                    f"https://www.nseindia.com/api/quote-equity?symbol={symbol.upper()}",
                    headers=NSE_HEADERS,
                )
                if resp.status_code == 200:
                    d = resp.json()
                    price_info = d.get("priceInfo", {})
                    mc_raw = d.get("securityInfo", {}).get("marketCap", "")
                    try:
                        mcap_str = (
                            f"{round(float(safe_str(mc_raw))/1e7, 2)} Cr"
                            if mc_raw else ""
                        )
                    except Exception:
                        mcap_str = safe_str(mc_raw)
                    result = {
                        "cmp": safe_str(price_info.get("lastPrice", "")),
                        "mcap": mcap_str,
                        "sector": safe_str(d.get("metadata", {}).get("industry", "")),
                    }
            elif exchange == "BSE":
                resp = await client.get(
                    f"https://api.bseindia.com/BseIndiaAPI/api/getScripHeaderData/w"
                    f"?Debtflag=&scripcode={symbol}&seriesid=",
                    headers=BSE_HEADERS,
                )
                if resp.status_code == 200:
                    d = resp.json()
                    result = {
                        "cmp": safe_str(
                            d.get("CurrRate", d.get("Curt_Rate", d.get("LTP", "")))
                        ),
                        "mcap": safe_str(d.get("Mktcap", d.get("MktCap", ""))),
                        "sector": safe_str(d.get("Indust", d.get("Industry", ""))),
                    }
        except Exception as e:
            print(f"Stock data error for {symbol} ({exchange}): {e}")
    return result


# ══════════════════════════════════════════════════════════════════════════════
#  CLASSIFICATION — COMPREHENSIVE FIX
# ══════════════════════════════════════════════════════════════════════════════

LIGHT_FILTER = [
    "authorised",
    "authorized",
    "share capital",
    "memorandum",
    "moa",
    "board meeting",
    "egm",
    "postal ballot",
    "special resolution",
]

AUTH_SEARCH_TERMS = [
    "authorised share capital",
    "authorized share capital",
    "increase in authorised share capital",
    "increase in authorized share capital",
    "alteration of capital clause",
    "capital clause",
    "memorandum of association",
    "amendment to memorandum",
    "preferential issue",
    "preferential allotment",
    "rights issue",
    "fund raising",
    "fundraising",
    "qualified institutions placement",
    "qip",
    "bonus issue",
    "stock split",
]

STRICT_AUTH_TERMS = [
    "increase in authorised share capital",
    "increase in authorized share capital",
    "existing authorised share capital",
    "existing authorized share capital",
    "authorised share capital of",
    "authorized share capital of",
    "alteration of capital clause",
    "preferential allotment",
    "preferential issue",
    "rights issue",
    "rights offer",
    "qualified institutions placement",
    "qip",
    "issue of warrants",
    "convertible warrants",
    "bonus issue",
    "stock split",
    "sub-division of shares",
]


def _any_match(text: str, patterns: list) -> bool:
    t = str(text).lower().strip()
    return any(p in t for p in patterns)


def get_relevant_chunks(text: str) -> str:
    if not text:
        return ""
    text_lower = text.lower()
    relevant_chunks = []
    
    # NEW EXTRACTION STRATEGY
    for keyword in AUTH_SEARCH_TERMS:
        start_search = 0
        while True:
            idx = text_lower.find(keyword, start_search)
            if idx == -1:
                break
            
            # extract 1500 chars before + 2500 after
            start = max(0, idx - 1500)
            end = min(len(text), idx + 2500)
            relevant_chunks.append(text[start:end])
            
            # Move search pointer to avoid infinite loop or redundant small overlaps
            start_search = idx + len(keyword)
            if len(relevant_chunks) > 15: # Safety cap
                break
    
    if not relevant_chunks:
        return text[:6000]
    
    # Merge all matches and cap
    return "\n\n".join(relevant_chunks)[:6000]


def classify_announcements(announcements: list) -> list:
    candidates = []
    for ann in announcements:
        # Initialize default fields
        for field in [
            "board_approval", "dobm", "existing_auth_cap", 
            "new_auth_cap", "proposed_increase", "remark_positive", 
            "remark_negative", "action"
        ]:
            ann.setdefault(field, "")
        ann.setdefault("action", "WATCH")
        ann.setdefault("is_auth_capital", False)

        subj = ann.get("subject", "").lower()
        if any(k in subj for k in LIGHT_FILTER):
            candidates.append(ann)
    return candidates


# ══════════════════════════════════════════════════════════════════════════════
#  PDF DOWNLOAD — WITH FALLBACK URLs
# ══════════════════════════════════════════════════════════════════════════════

async def download_pdf_text(url: str, exchange: str, ann: dict = None) -> str:
    if not url or not url.strip():
        return ""
    if not PYPDF_AVAILABLE:
        return ""

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Referer": "https://www.nseindia.com/" if exchange == "NSE" else "https://www.bseindia.com/",
    }

    print(f"\n📥 DOWNLOADING PDF: {url[:100]}")
    try:
        async with httpx.AsyncClient(timeout=45, follow_redirects=True, verify=False) as client:
            resp = await client.get(url, headers=headers)
            print(f"  Status: {resp.status_code}")
            print(f"  Content-Type: {resp.headers.get('content-type')}")

            if resp.status_code != 200:
                return ""

            ct = resp.headers.get("content-type", "").lower()
            is_pdf = "pdf" in ct or url.lower().endswith(".pdf") or resp.content[:4] == b"%PDF"
            if not is_pdf:
                print("  ❌ Not a PDF response")
                return ""

            reader = PdfReader(io.BytesIO(resp.content))
            pages = []
            # Increase page limit to find buried info
            for page in reader.pages[:50]:
                try:
                    txt = page.extract_text()
                    if txt:
                        pages.append(txt)
                except Exception:
                    pass
            text = "\n".join(pages).strip()
            print(f"  ✅ Extracted PDF raw text length: {len(text)}")
            
            # Apply smart chunking
            chunked_text = get_relevant_chunks(text)
            print(f"  ✂️ Chunked text length: {len(chunked_text)}")
            return chunked_text

    except Exception as e:
        print(f"  ❌ PDF extraction failed: {e}")
        return ""


async def extract_fields_from_pdf(pdf_text: str, company: str, subject: str, retry_count=0) -> dict:
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

Extract these fields. Return ONLY valid JSON, no markdown:
{{
  "board_approval": "DD-MM-YYYY date when Board approved the capital increase. Look for 'Board of Directors at their meeting held on', 'approved by the Board on'. Empty string if not found.",
  "dobm": "DD-MM-YYYY date of Board Meeting. Usually same as board_approval. Look for 'meeting held on', 'board meeting dated'. Empty string if not found.",
  "existing_auth_cap": "Existing Authorised Equity Capital in INR with commas e.g. '10,00,00,000'. Look for 'existing authorised capital', 'present authorised capital', 'from Rs.'. Empty if not found.",
  "new_auth_cap": "New Authorised Equity Capital in INR with commas e.g. '20,00,00,000'. Look for 'increased to', 'new authorised capital', 'to Rs.'. Empty if not found.",
  "proposed_increase": "Increase amount = New minus Existing in INR. Look for 'by Rs.', 'increase of Rs.'. Calculate if both found. Empty if not found.",
  "remark_positive": "1 line positive remark based on PDF content — expansion plans, fundraising flexibility, growth signal.",
  "remark_negative": "1 line risk remark — equity dilution risk, overleveraging, shareholder value concern.",
  "action": "One of: BUY on dip / ACCUMULATE / WATCH / NEUTRAL / AVOID — based on context and magnitude."
}}"""
    try:
        loop = asyncio.get_running_loop()
        def call_groq():
            return client.chat.completions.create(
                model="llama-3.3-70b-versatile",
                messages=[{"role": "user", "content": prompt}],
                temperature=0.0,
                max_tokens=1000,
            )
        
        resp = await loop.run_in_executor(None, call_groq)
        raw = resp.choices[0].message.content.strip()
        raw = re.sub(r"^```json\s*", "", raw)
        raw = re.sub(r"^```\s*",     "", raw)
        raw = re.sub(r"\s*```$",     "", raw)
        result = json.loads(raw)
        for k in empty:
            result.setdefault(k, "")
        return result
    except Exception as e:
        if "429" in str(e) and retry_count < 2:
            print(f"  ⏳ Rate limited (extract). Sleeping 12s... (Retry {retry_count+1})")
            await asyncio.sleep(12)
            return await extract_fields_from_pdf(pdf_text, company, subject, retry_count + 1)
        print(f"  PDF AI extraction error: {e}")
        return empty


# ══════════════════════════════════════════════════════════════════════════════
#  DEEP PDF SCAN — EXPANDED
# ══════════════════════════════════════════════════════════════════════════════

async def ai_classify_with_pdf(ann: dict, retry_count=0) -> bool:
    if not GROQ_API_KEY:
        return False

    subject = ann.get("subject", "")
    company = ann.get("company", "")
    exchange = ann.get("exchange", "NSE")
    url = ann.get("link", "")

    pdf_text = await download_pdf_text(url, exchange, ann)
    if pdf_text:
        ann["_pdf_text_cache"] = pdf_text

    # CRITICAL FIX — HARD RULE BEFORE AI
    pdf_lower = pdf_text.lower()
    
    hard_match = any(term in pdf_lower for term in STRICT_AUTH_TERMS)

    if not hard_match:
        print(f"❌ HARD FILTER REJECTED: {company}")
        return False

    # FOURTH FIX — BETTER PROMPT
    prompt = f"""
You are a stock market corporate announcement classifier.

Determine whether this announcement specifically involves:
1. Increase in authorised/authorized share capital
2. Preferential allotment
3. Rights issue
4. QIP
5. Fund raising through issuance of securities
6. Bonus issue
7. Stock split/sub-division
8. Alteration of capital clause in MOA

DO NOT classify normal:
- financial results
- expansion plans
- capex
- operational updates
- newspaper publication
- analyst calls
- investor presentations
- manufacturing investments
as YES.

Respond ONLY with:
YES
or
NO

SUBJECT:
{subject}

TEXT:
{pdf_text[:8000]}
"""

    client = GroqClient(api_key=GROQ_API_KEY)
    try:
        loop = asyncio.get_running_loop()
        def call_groq():
            resp = client.chat.completions.create(
                model="llama-3.3-70b-versatile",
                messages=[{"role": "user", "content": prompt}],
                temperature=0.0,
                max_tokens=10,
            )
            return resp.choices[0].message.content.strip()

        result = await loop.run_in_executor(None, call_groq)
        is_auth = "YES" in result.upper()
        
        if is_auth:
            print(f"\n✅ AI YES: {company}")
            for term in STRICT_AUTH_TERMS:
                if term in pdf_lower:
                    print(f"   Matched term: {term}")
            ann["is_auth_capital"] = True
            return True
        else:
            return False
            
    except Exception as e:
        if "429" in str(e) and retry_count < 2:
            print(f"  ⏳ Rate limited (classify). Sleeping 12s... (Retry {retry_count+1})")
            await asyncio.sleep(12)
            return await ai_classify_with_pdf(ann, retry_count + 1)
        print(f"  AI classification error for {company}: {e}")
        return False


# ══════════════════════════════════════════════════════════════════════════════
#  PDF ENRICHMENT
# ══════════════════════════════════════════════════════════════════════════════

async def enrich_from_pdf(auth_anns: list) -> list:
    if not auth_anns:
        return auth_anns

    final = []
    for i, ann in enumerate(auth_anns):
        print(f"  Enriching {i+1}/{len(auth_anns)}: {ann.get('company')}")
        if "_pdf_text_cache" in ann:
            pdf_text = ann.get("_pdf_text_cache")
            print(f"    Using cached PDF text")
        else:
            print(f"    Downloading PDF for enrichment: {ann.get('link', '')[:60]}")
            pdf_text = await download_pdf_text(
                ann.get("link", ""), ann.get("exchange", "NSE"), ann
            )

        if pdf_text:
            fields = await extract_fields_from_pdf(
                pdf_text, ann.get("company", ""), ann.get("subject", ""),
            )
            for k in [
                "board_approval", "dobm", "existing_auth_cap",
                "new_auth_cap", "proposed_increase",
                "remark_positive", "remark_negative", "action",
            ]:
                if fields.get(k):
                    ann[k] = fields[k]
            ann["_pdf_extracted"] = True
        else:
            ann["_pdf_extracted"] = False
        final.append(ann)
    
    return final


# ══════════════════════════════════════════════════════════════════════════════
#  EXCEL GENERATION
# ══════════════════════════════════════════════════════════════════════════════

COLUMNS = [
    ("Sr.no",                    8),
    ("Date of Entry",           13),
    ("Name of the Company",     28),
    ("Board Approval",          14),
    ("DOBM",                    12),
    ("Exst Auth Eq Cap (INR)",  18),
    ("New Auth Eq Cap (INR)",   18),
    ("Proposed Increase (INR)", 18),
    ("CMP",                      9),
    ("M Cap (in Cr)",           13),
    ("Sector",                  18),
    ("Remark Positive",         28),
    ("Remark Negative",         28),
    ("Action",                  14),
    ("Link",                    35),
]


def style_header(cell, bg: str):
    cell.font = Font(bold=True, color="FFFFFF", size=10, name="Calibri")
    cell.fill = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    s = Side(style="thin", color="CCCCCC")
    cell.border = Border(left=s, right=s, top=s, bottom=s)


def style_data(cell, wrap=False):
    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=wrap)
    s = Side(style="thin", color="E0E0E0")
    cell.border = Border(left=s, right=s, top=s, bottom=s)
    cell.font = Font(size=9, name="Calibri")


def write_sheet(ws, rows: list, header_color: str, title: str):
    ws.merge_cells(f"A1:{get_column_letter(len(COLUMNS))}1")
    tc = ws["A1"]
    tc.value = title
    tc.font = Font(bold=True, size=13, color="FFFFFF", name="Calibri")
    tc.fill = PatternFill("solid", fgColor="1A1A2E")
    tc.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    for ci, (name, width) in enumerate(COLUMNS, 1):
        cell = ws.cell(row=2, column=ci, value=name)
        style_header(cell, header_color)
        ws.column_dimensions[get_column_letter(ci)].width = width
    ws.row_dimensions[2].height = 35

    for ri, ann in enumerate(rows, 3):
        ws.row_dimensions[ri].height = 55
        stock = ann.get("_stock", {}) or {}
        vals = [
            ri - 2,
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
            safe_str(ann.get("web_link", ann.get("link", ""))),
        ]
        for ci, val in enumerate(vals, 1):
            cell = ws.cell(row=ri, column=ci, value=val)
            style_data(cell, wrap=(ci in [3, 12, 13, 15]))
            if ri % 2 == 0:
                cell.fill = PatternFill("solid", fgColor="F8F9FF")
            if ci == 14 and val:
                v = str(val).upper()
                for kw, clr in [
                    ("BUY", "C6EFCE"), ("ACCUMULATE", "C6EFCE"),
                    ("AVOID", "FFC7CE"), ("NEUTRAL", "FFEB9C"), ("WATCH", "FFEB9C"),
                ]:
                    if kw in v:
                        cell.fill = PatternFill("solid", fgColor=clr)
                        break
    ws.freeze_panes = "A3"


def generate_excel(auth_anns: list, other_anns: list, date_str: str) -> str:
    wb = openpyxl.Workbook()
    ws_auth = wb.active
    ws_auth.title = "Auth Capital"
    write_sheet(
        ws_auth, auth_anns, "1565C8",
        f"Authorised Capital Increase Announcements — {date_str}",
    )
    ws_other = wb.create_sheet("Other Announcements")
    write_sheet(
        ws_other, other_anns, "C85215",
        f"Other Announcements — {date_str}",
    )
    filename = f"announcements_{date_str.replace(' ','_').replace('-','')}.xlsx"
    wb.save(os.path.join(OUTPUT_DIR, filename))
    return filename


# ══════════════════════════════════════════════════════════════════════════════
#  MAIN PIPELINE
# ══════════════════════════════════════════════════════════════════════════════

fetch_status = {
    "status": "idle",
    "message": "",
    "progress": 0,
    "filename": "",
}


@app.get("/status")
async def get_status():
    return fetch_status


@app.post("/fetch")
async def trigger_fetch(req: FetchRequest, background_tasks: BackgroundTasks):
    if fetch_status["status"] == "running":
        raise HTTPException(status_code=409, detail="Fetch already in progress")
    today = date.today().strftime("%d-%m-%Y")
    from_date = req.from_date or today
    to_date   = req.to_date   or today
    background_tasks.add_task(run_pipeline, from_date, to_date)
    return {"message": "Fetch started", "from_date": from_date, "to_date": to_date}


async def run_pipeline(from_date: str, to_date: str):
    global fetch_status
    loop = asyncio.get_running_loop()
    try:
        fetch_status = {
            "status": "running",
            "message": "Fetching NSE announcements...",
            "progress": 10,
            "filename": "",
        }

        nse_anns = await fetch_nse_announcements(from_date, to_date)
        fetch_status = {
            **fetch_status,
            "message": f"NSE: {len(nse_anns)} found. Fetching BSE...",
            "progress": 25,
        }

        bse_anns = await fetch_bse_announcements(from_date, to_date)
        fetch_status = {
            **fetch_status,
            "message": f"BSE: {len(bse_anns)} found. Classifying...",
            "progress": 40,
        }

        all_anns = nse_anns + bse_anns
        
        # Deduplication
        seen = set()
        unique_anns = []
        for ann in all_anns:
            key = (
                ann.get("exchange"),
                ann.get("company"),
                ann.get("subject"),
                ann.get("date", "")[:10],
            )
            if key not in seen:
                seen.add(key)
                unique_anns.append(ann)
        all_anns = unique_anns

        print(f"\n{'='*60}")
        print(f"TOTAL UNIQUE FETCHED: {len(all_anns)} (NSE={len(nse_anns)}, BSE={len(bse_anns)})")
        print(f"{'='*60}")

        if not all_anns:
            fetch_status = {
                "status": "done",
                "message": "No announcements found. Try a different date.",
                "progress": 100,
                "filename": "",
                "auth_count": 0,
                "other_count": 0,
                "total": 0,
                "announcements": [],
            }
            return

        candidates = classify_announcements(all_anns)
        print(f"\nShortlisted {len(candidates)} candidates from {len(all_anns)} total.")

        if candidates and GROQ_API_KEY:
            fetch_status = {
                **fetch_status,
                "message": f"AI Classifying {len(candidates)} candidates...",
                "progress": 45,
            }
            # Sequential processing to respect Groq rate limits
            for i, c in enumerate(candidates):
                print(f"  Processing candidate {i+1}/{len(candidates)}")
                await ai_classify_with_pdf(c)

        auth_anns  = [a for a in all_anns if a.get("is_auth_capital")]
        other_anns = [a for a in all_anns if not a.get("is_auth_capital")]
        print(f"\nFinal Classification: Auth={len(auth_anns)}, Other={len(other_anns)}")

        if auth_anns:
            fetch_status = {
                **fetch_status,
                "message": (
                    f"Found {len(auth_anns)} auth capital announcements. "
                    f"Reading PDFs..."
                ),
                "progress": 60,
            }
            auth_anns = await enrich_from_pdf(auth_anns)

            # ══════════════════════════════════════════════════════════════════════════
            # FINAL NUMERIC VALIDATION LAYER
            # ══════════════════════════════════════════════════════════════════════════
            validated_auth = []
            FUND_RAISE_TERMS = [
                "preferential allotment",
                "preferential issue",
                "rights issue",
                "qip",
                "qualified institutions placement",
                "warrants",
                "convertible warrants",
                "issue of securities",
                "equity shares",
            ]

            for ann in auth_anns:
                existing_cap = str(ann.get("existing_auth_cap", "")).strip()
                new_cap = str(ann.get("new_auth_cap", "")).strip()
                
                # Check the PDF text cache we stored during classification
                pdf_text = str(ann.get("_pdf_text_cache", "")).lower()
                
                has_existing = bool(existing_cap)
                has_new = bool(new_cap)

                strict_cap_raise = (
                    has_existing
                    and has_new
                    and existing_cap != new_cap
                )

                strict_fundraise = (
                    "preferential allotment" in pdf_text
                    or "rights issue" in pdf_text
                    or "qualified institutions placement" in pdf_text
                    or "qip" in pdf_text
                )

                if strict_cap_raise or strict_fundraise:
                    validated_auth.append(ann)
                else:
                    print(f"❌ Removed false positive: {ann.get('company')}")
                    # Update its flag so it doesn't appear in the Auth sheet
                    ann["is_auth_capital"] = False

            auth_anns = validated_auth
            # Refresh other_anns after some were demoted to false positive
            other_anns = [a for a in all_anns if not a.get("is_auth_capital")]

            # Cleanup PDF text cache to keep JSON response small
            for a in all_anns:
                a.pop("_pdf_text_cache", None)

        fetch_status = {
            **fetch_status,
            "message": "Fetching stock prices...",
            "progress": 75,
        }

        all_anns = auth_anns + other_anns
        sem = asyncio.Semaphore(5)

        async def fetch_s(sym, exch):
            async with sem:
                return await fetch_stock_data(sym, exch)

        stock_results = await asyncio.gather(
            *[fetch_s(a["symbol"], a["exchange"]) for a in all_anns],
            return_exceptions=True,
        )
        for ann, stock in zip(all_anns, stock_results):
            ann["_stock"] = stock if isinstance(stock, dict) else {}

        fetch_status = {
            **fetch_status,
            "message": "Generating Excel...",
            "progress": 88,
        }

        date_label = (
            f"{from_date} to {to_date}" if from_date != to_date else from_date
        )
        filename = await loop.run_in_executor(
            None, generate_excel, auth_anns, other_anns, date_label
        )
        print(f"\nExcel generated: {filename}")

        fetch_status = {
            "status": "done",
            "message": (
                f"Done! {len(auth_anns)} auth capital + "
                f"{len(other_anns)} other announcements."
            ),
            "progress": 100,
            "filename": filename,
            "auth_count": len(auth_anns),
            "other_count": len(other_anns),
            "total": len(all_anns),
            "announcements": all_anns,
        }

    except Exception as e:
        import traceback
        traceback.print_exc()
        fetch_status = {
            "status": "error",
            "message": f"Error: {str(e)}",
            "progress": 0,
            "filename": "",
        }


# ══════════════════════════════════════════════════════════════════════════════
#  API ENDPOINTS
# ══════════════════════════════════════════════════════════════════════════════

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


# ══════════════════════════════════════════════════════════════════════════════
#  DEBUG ENDPOINTS — TO SEE WHAT THE API ACTUALLY RETURNS
# ══════════════════════════════════════════════════════════════════════════════

@app.get("/debug/subjects")
async def debug_subjects():
    anns = fetch_status.get("announcements", [])
    if not anns:
        return {"message": "No announcements fetched yet. Run /fetch first."}

    subjects = {}
    for a in anns:
        subj = a.get("subject", "")
        if subj not in subjects:
            subjects[subj] = {
                "count": 0,
                "exchange": a.get("exchange", ""),
                "sample_company": a.get("company", ""),
                "is_auth": a.get("is_auth_capital", False),
            }
        subjects[subj]["count"] += 1

    return {
        "total_announcements": len(anns),
        "unique_subjects": len(subjects),
        "subjects": subjects,
    }


@app.get("/debug/raw")
async def debug_raw(limit: int = 5):
    anns = fetch_status.get("announcements", [])
    if not anns:
        return {"message": "No announcements fetched yet. Run /fetch first."}

    results = []
    for a in anns[:limit]:
        raw = a.get("raw", {})
        serializable = {}
        for k, v in raw.items():
            try:
                json.dumps(v)
                serializable[k] = v
            except (TypeError, ValueError):
                serializable[k] = str(v)

        results.append({
            "exchange": a.get("exchange"),
            "company": a.get("company"),
            "subject": a.get("subject"),
            "link": a.get("link"),
            "attchmnt_text": a.get("attchmnt_text", ""),
            "bse_attachment": a.get("bse_attachment", ""),
            "nse_sm_name": a.get("nse_sm_name", ""),
            "nse_sub_type": a.get("nse_sub_type", ""),
            "nse_an_type": a.get("nse_an_type", ""),
            "bse_category": a.get("bse_category", ""),
            "bse_subcategory": a.get("bse_subcategory", ""),
            "bse_subcatname": a.get("bse_subcatname", ""),
            "is_auth_capital": a.get("is_auth_capital", False),
            "raw_fields": serializable,
        })

    return {
        "showing": len(results),
        "total_available": len(anns),
        "announcements": results,
    }


@app.get("/debug/search")
async def debug_search(q: str = "affle"):
    anns = fetch_status.get("announcements", [])
    if not anns:
        return {"message": "No announcements fetched yet. Run /fetch first."}

    q_lower = q.lower()
    matches = []
    for a in anns:
        if (
            q_lower in a.get("company", "").lower()
            or q_lower in a.get("subject", "").lower()
            or q_lower in a.get("symbol", "").lower()
        ):
            raw = a.get("raw", {})
            serializable = {}
            for k, v in raw.items():
                try:
                    json.dumps(v)
                    serializable[k] = v
                except (TypeError, ValueError):
                    serializable[k] = str(v)

            matches.append({
                "exchange": a.get("exchange"),
                "company": a.get("company"),
                "symbol": a.get("symbol"),
                "subject": a.get("subject"),
                "date": a.get("date"),
                "link": a.get("link"),
                "attchmnt_text": a.get("attchmnt_text", ""),
                "bse_attachment": a.get("bse_attachment", ""),
                "is_auth_capital": a.get("is_auth_capital", False),
                "needs_deep_scan": a.get("_needs_deep_scan", False),
                "raw_fields": serializable,
            })

    return {
        "query": q,
        "matches": len(matches),
        "results": matches,
    }

# Serve frontend files
app.mount("/frontend", StaticFiles(directory="frontend"), name="frontend")

# Homepage route
@app.get("/")
async def serve_home():
    return FileResponse("frontend/index.html")