from fastapi import FastAPI, BackgroundTasks, HTTPException, Depends, Security
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.security import APIKeyHeader
import httpx
import json
import os
import re
import asyncio
from datetime import datetime, date
from typing import Optional
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from groq import AsyncGroq
import io
import pdfplumber
from pydantic import BaseModel
from dotenv import load_dotenv
import sqlite3

# Load .env file automatically
load_dotenv()

app = FastAPI(title="NSE/BSE Announcement Tracker")

app.add_middleware(
    CORSMiddleware,
    allow_origins=os.getenv("ALLOWED_ORIGINS", "*").split(","),
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

GROQ_API_KEY = os.getenv("GROQ_API_KEY", "")

# Security Setup
API_KEY = os.getenv("API_KEY", "admin123")
api_key_header = APIKeyHeader(name="X-API-Key")

def verify_api_key(api_key: str = Security(api_key_header)):
    if api_key != API_KEY:
        raise HTTPException(status_code=403, detail="Invalid API Key")
    return api_key

# Storage Setup (Supports Render persistent disks)
IS_PRODUCTION = os.getenv("RENDER") == "true"
DATA_DIR = "/data" if IS_PRODUCTION else "."

OUTPUT_DIR = os.path.join(DATA_DIR, "output")
os.makedirs(OUTPUT_DIR, exist_ok=True)
DB_PATH = os.path.join(DATA_DIR, "announcements.db")

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

def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS announcements (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            exchange TEXT,
            company TEXT,
            symbol TEXT,
            subject TEXT,
            date TEXT,
            link TEXT,
            is_auth_capital BOOLEAN,
            board_approval TEXT,
            dobm TEXT,
            existing_auth_cap TEXT,
            new_auth_cap TEXT,
            proposed_increase TEXT,
            cmp TEXT,
            mcap TEXT,
            sector TEXT,
            remark_positive TEXT,
            remark_negative TEXT,
            action TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(exchange, symbol, date, subject)
        )
    """)
    conn.commit()
    conn.close()

init_db()

def save_to_db(announcements: list):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    for a in announcements:
        stock = a.get("_stock", {})
        try:
            c.execute("""
                INSERT OR REPLACE INTO announcements (
                    exchange, company, symbol, subject, date, link,
                    is_auth_capital, board_approval, dobm,
                    existing_auth_cap, new_auth_cap, proposed_increase,
                    cmp, mcap, sector, remark_positive, remark_negative, action
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                a.get("exchange"), a.get("company"), a.get("symbol"), a.get("subject"), a.get("date"), a.get("link"),
                a.get("is_auth_capital", False), a.get("board_approval", ""), a.get("dobm", ""),
                a.get("existing_auth_cap", ""), a.get("new_auth_cap", ""), a.get("proposed_increase", ""),
                stock.get("cmp", ""), stock.get("mcap", ""), stock.get("sector", a.get("sector", "")),
                a.get("remark_positive", ""), a.get("remark_negative", ""), a.get("action", "NEUTRAL")
            ))
        except Exception as e:
            # Duplicate or other error
            pass
    conn.commit()
    conn.close()

def get_from_db(from_date: str = None, to_date: str = None) -> list:
    """
    Fetch announcements from the DB.
    When from_date/to_date are supplied (DD-MM-YYYY) we filter records whose
    `created_at` timestamp falls within the last 24 h of the pipeline run so
    we only return what was just ingested, not all historical records.
    Falls back to the 500 most-recent rows when no date range is given.
    """
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()

    if from_date and to_date:
        # Parse the DD-MM-YYYY dates the pipeline uses
        try:
            dt_from = datetime.strptime(from_date, "%d-%m-%Y")
            dt_to   = datetime.strptime(to_date,   "%d-%m-%Y")
            # Include the whole to_date day
            ts_from = dt_from.strftime("%Y-%m-%d 00:00:00")
            ts_to   = dt_to.strftime(  "%Y-%m-%d 23:59:59")
            c.execute(
                "SELECT * FROM announcements WHERE created_at BETWEEN ? AND ? "
                "ORDER BY created_at DESC LIMIT 2000",
                (ts_from, ts_to),
            )
        except Exception:
            # Fallback: most recent 500
            c.execute("SELECT * FROM announcements ORDER BY created_at DESC LIMIT 500")
    else:
        c.execute("SELECT * FROM announcements ORDER BY created_at DESC LIMIT 500")

    rows = [dict(r) for r in c.fetchall()]
    conn.close()

    # Map DB fields back to _stock dict for frontend compatibility
    for r in rows:
        r["_stock"] = {"cmp": r.pop("cmp"), "mcap": r.pop("mcap"), "sector": r.pop("sector")}
        r["is_auth_capital"] = bool(r["is_auth_capital"])
    return rows


class FetchRequest(BaseModel):
    from_date: Optional[str] = None  # DD-MM-YYYY
    to_date: Optional[str] = None    # DD-MM-YYYY


# ─── NSE FETCHER ─────────────────────────────────────────────────────────────

async def fetch_nse_announcements(from_date: str, to_date: str) -> list:
    announcements = []
    async with httpx.AsyncClient(timeout=30, follow_redirects=True) as client:
        try:
            # Warm up session (NSE requires cookie)
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
                        "company": item.get("comp", item.get("symbol", "")),
                        "symbol": item.get("symbol", ""),
                        "subject": item.get("subject") or item.get("desc") or item.get("attchmntText") or "",
                        "date": item.get("an_dt") or item.get("date") or "",
                        "link": item.get("attchmntFile") or f"https://www.nseindia.com/api/corporate-announcements?symbol={item.get('symbol','')}&an_num={item.get('an_num','')}",
                        "raw": item,
                    })
        except Exception as e:
            print(f"NSE fetch error: {e}")
    return announcements


async def fetch_bse_announcements(from_date: str, to_date: str) -> list:
    announcements = []
    # Convert DD-MM-YYYY → YYYYMMDD for BSE
    try:
        d_from = datetime.strptime(from_date, "%d-%m-%Y").strftime("%Y%m%d")
        d_to   = datetime.strptime(to_date,   "%d-%m-%Y").strftime("%Y%m%d")
    except Exception:
        d_from = from_date.replace("-", "")
        d_to   = to_date.replace("-", "")

    async with httpx.AsyncClient(timeout=30, follow_redirects=True) as client:
        try:
            url = (
                f"https://api.bseindia.com/BseIndiaAPI/api/AnnSubCategoryGetData/w"
                f"?strCat=-1&strPrevDate={d_from}&strScrip=&strSearch=P"
                f"&strToDate={d_to}&strType=C&subcategory=-1"
            )
            resp = await client.get(url, headers=BSE_HEADERS)
            if resp.status_code == 200:
                data = resp.json()
                items = data.get("Table", data if isinstance(data, list) else [])
                for item in items:
                    scrip = item.get("SCRIP_CD", "")
                    announcements.append({
                        "exchange": "BSE",
                        "company": item.get("SLONGNAME", item.get("short_name", "")),
                        "symbol": str(scrip),
                        "subject": item.get("HEADLINE", item.get("subject", "")),
                        "date": item.get("NEWS_DT", item.get("date", "")),
                        "link": f"https://www.bseindia.com/xml-data/corpfiling/AttachHis/{item.get('ATTACHMENTNAME', '')}",
                        "raw": item,
                    })
        except Exception as e:
            print(f"BSE fetch error: {e}")
    return announcements


# ─── STOCK DATA FETCHER ───────────────────────────────────────────────────────

def safe_str(val) -> str:
    """Safely convert any value to a plain string for Excel."""
    if val is None:
        return ""
    if isinstance(val, dict):
        # BSE sometimes returns nested dicts like {'LTP': '151.60', ...}
        # Try common price keys first
        for key in ("LTP", "CurrRate", "lastPrice", "price", "value"):
            if key in val:
                return str(val[key])
        return str(list(val.values())[0]) if val else ""
    if isinstance(val, (list, tuple)):
        return str(val[0]) if val else ""
    return str(val)


async def fetch_stock_data(client: httpx.AsyncClient, symbol: str, exchange: str, semaphore: asyncio.Semaphore) -> dict:
    """Fetch CMP, MCap, Sector for a given symbol with concurrency limit"""
    result = {"cmp": "", "mcap": "", "sector": ""}
    if not symbol:
        return result
    
    async with semaphore:
        try:
            if exchange == "NSE":
                url = f"https://www.nseindia.com/api/quote-equity?symbol={symbol.upper()}"
                resp = await client.get(url, headers=NSE_HEADERS)
                # If we get 401 or 403, try warming up again (rare but possible during batch)
                if resp.status_code in (401, 403):
                    await client.get("https://www.nseindia.com", headers=NSE_HEADERS)
                    resp = await client.get(url, headers=NSE_HEADERS)
                
                if resp.status_code == 200:
                    d = resp.json()
                    price_info = d.get("priceInfo", {})
                    meta       = d.get("metadata", {})
                    cmp        = safe_str(price_info.get("lastPrice", ""))
                    # Try proper mcap field
                    mc_raw = d.get("securityInfo", {}).get("marketCap", "")
                    try:
                        mcap_str = f"{round(float(safe_str(mc_raw))/1e7, 2)} Cr" if mc_raw else ""
                    except Exception:
                        mcap_str = safe_str(mc_raw)
                    result = {
                        "cmp":    cmp,
                        "mcap":   mcap_str,
                        "sector": safe_str(meta.get("industry", "")),
                    }
            elif exchange == "BSE":
                url = (
                    f"https://api.bseindia.com/BseIndiaAPI/api/getScripHeaderData/w"
                    f"?Debtflag=&scripcode={symbol}&seriesid="
                )
                resp = await client.get(url, headers=BSE_HEADERS)
                if resp.status_code == 200:
                    d = resp.json()
                    cmp_raw    = d.get("CurrRate", d.get("Curt_Rate", d.get("LTP", "")))
                    mcap_raw   = d.get("Mktcap",   d.get("MktCap",   ""))
                    sector_raw = d.get("Indust",   d.get("Industry", ""))
                    result = {
                        "cmp":    safe_str(cmp_raw),
                        "mcap":   safe_str(mcap_raw),
                        "sector": safe_str(sector_raw),
                    }
        except Exception as e:
            print(f"Stock data fetch error for {symbol} ({exchange}): {e}")
    return result


# ─── AI CLASSIFIER ───────────────────────────────────────────────────────────




def extract_capital_fallback(text: str) -> dict:
    res = {"existing_auth_cap": "", "new_auth_cap": "", "proposed_increase": "", "action": "NEUTRAL", "board_approval": "", "dobm": "", "remark_positive": "", "remark_negative": ""}
    if not text: return res
    pattern = r'(?i)from\s+(?:rs\.?|rupees|inr)?\s*([\d,.]+)\s*(?:crores?|lakhs?)?\s*to\s+(?:rs\.?|rupees|inr)?\s*([\d,.]+)'
    match = re.search(pattern, text)
    if match:
        res["existing_auth_cap"] = match.group(1)
        res["new_auth_cap"] = match.group(2)
        try:
            ex = float(match.group(1).replace(",", ""))
            nw = float(match.group(2).replace(",", ""))
            if nw > ex:
                res["proposed_increase"] = str(nw - ex)
                res["action"] = "BUY on dip"
        except: pass
    return res


def extract_pdf_text_sync(content: bytes) -> str:
    text = ""
    try:
        with pdfplumber.open(io.BytesIO(content)) as pdf:
            for i in range(min(5, len(pdf.pages))):
                page_text = pdf.pages[i].extract_text()
                if page_text: text += page_text + "\n"
    except Exception as e:
        print(f"pdfplumber extraction error: {e}")
    return text

async def classify_and_extract_async(announcements: list) -> list:
    """Scan subjects and PDFs, then use Groq to extract auth capital fields."""
    
    # Identify candidates based on subject
    candidates = []
    for ann in announcements:
        subj = ann.get("subject", "").lower()
        if any(kw in subj for kw in ["authori", "capital", "outcome of board meeting", "board meeting", "general meeting", "egm", "agm", "postal ballot", "scrutinizer report", "voting results"]):
            candidates.append(ann)
        else:
            ann["is_auth_capital"] = False
            ann.update({k: "" for k in ["board_approval", "dobm", "existing_auth_cap", "new_auth_cap", "proposed_increase", "remark_positive", "remark_negative", "action"]})
            
    # Fetch PDFs for candidates concurrently
    semaphore = asyncio.Semaphore(15)
    async with httpx.AsyncClient(verify=False, follow_redirects=True) as client:
        # Warm up if any NSE candidates
        if any("nseindia.com" in (a.get("link") or "") for a in candidates):
            try: await client.get("https://www.nseindia.com", headers=NSE_HEADERS, timeout=10)
            except: pass

        async def fetch_pdf(ann):
            link = ann.get("link")
            if not link: return
            
            # Use headers but allow PDF content type
            h = BSE_HEADERS.copy() if "bseindia.com" in link else NSE_HEADERS.copy()
            h["Accept"] = "application/pdf, */*"
            
            # BSE Path Fallback: Try AttachHis then AttachLive
            links_to_try = [link]
            if "bseindia.com" in link and "AttachHis" in link:
                links_to_try.append(link.replace("AttachHis", "AttachLive"))
            
            async with semaphore:
                for target_link in links_to_try:
                    for attempt in range(3):
                        try:
                            r = await client.get(target_link, headers=h, timeout=30)
                            if r.status_code == 200:
                                text = await asyncio.to_thread(extract_pdf_text_sync, r.content)
                                ann["_pdf_text"] = text
                                return # Success, exit function
                            elif r.status_code == 404:
                                break # 404, try next link in links_to_try
                            else:
                                print(f"PDF download failed for {ann.get('symbol')} at {target_link}: {r.status_code}")
                                break # Other status, don't retry, try next link
                        except Exception as e:
                            print(f"PDF error for {ann.get('symbol')} at {target_link} (attempt {attempt+1}): {e}")
                            await asyncio.sleep(2)
        
        await asyncio.gather(*(fetch_pdf(a) for a in candidates))

    # Refine candidates based on PDF text
    deep_targets = []
    for ann in candidates:
        subj = ann.get("subject", "").lower()
        raw_text = ann.get("_pdf_text", "")
        text = raw_text.lower()
        # Normalize text to catch spaced-out words like "A u t h o r i s e d"
        norm = re.sub(r"\s+", "", text)
        
        is_candidate = False
        # If subject explicitly mentions it
        if any(kw in subj for kw in ["authoris", "authoriz", "auth capital", "alteration of capital"]):
            is_candidate = True
        # If PDF text contains both 'authori' AND 'capital' in close proximity or normalized
        elif ("authori" in text and "capital" in text) or ("authori" in norm and "capital" in norm):
            # Guard against "authority" false positives if possible, though LLM will filter anyway
            is_candidate = True
            
        if is_candidate:
            deep_targets.append(ann)
        else:
            ann["is_auth_capital"] = False
            ann.update({k: "" for k in ["board_approval", "dobm", "existing_auth_cap", "new_auth_cap", "proposed_increase", "remark_positive", "remark_negative", "action"]})

    # Query Groq for deep_targets
    if not deep_targets or not GROQ_API_KEY:
        for ann in deep_targets:
            subj = ann.get("subject", "").lower()
            text = ann.get("_pdf_text", "").lower()
            norm = re.sub(r"\s+", "", text)
            
            is_auth = False
            if any(kw in subj for kw in ["authoris", "authoriz", "auth capital", "alteration of capital"]):
                is_auth = True
            elif ("authori" in text and "capital" in text) or ("authori" in norm and "capital" in norm):
                is_auth = True
                
            ann["is_auth_capital"] = is_auth
            fallback_data = extract_capital_fallback(text)
            ann.update(fallback_data)
        return announcements
    
    client = AsyncGroq(api_key=GROQ_API_KEY)
    SYSTEM_PROMPT = """You are a senior Indian equity market analyst. Extract structural data from the announcement.
You MUST respond with ONLY a valid JSON object — no explanation, no markdown fences."""

    for ann in deep_targets:
        text = ann.get("_pdf_text", "")
        text = text[:4000] # truncate to save tokens
        
        user_prompt = f"""Analyze this corporate announcement.
Subject: {ann.get("subject")}
PDF Content:
{text}

Return a JSON object with:
"is_auth_capital": boolean (true if about increase in Authorised Equity Capital),
"board_approval": string (DD-MM-YYYY),
"dobm": string (DD-MM-YYYY),
"existing_auth_cap": string (e.g. "10,00,00,000"),
"new_auth_cap": string,
"proposed_increase": string,
"remark_positive": string (1 line),
"remark_negative": string (1 line),
"action": string ("BUY on dip", "ACCUMULATE", "WATCH", "NEUTRAL", "AVOID")

Reply ONLY with the JSON object."""
        try:
            response = await client.chat.completions.create(
                model="llama-3.1-8b-instant",
                messages=[
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user",   "content": user_prompt},
                ],
                response_format={"type": "json_object"},
                temperature=0.1,
                max_tokens=1000,
            )
            raw = response.choices[0].message.content.strip()
            raw = re.sub(r"^```json\s*", "", raw)
            raw = re.sub(r"^```\s*",     "", raw)
            raw = re.sub(r"\s*```$",     "", raw)
            res = json.loads(raw)
            
            ann.update({
                "is_auth_capital":   res.get("is_auth_capital",   False),
                "board_approval":    res.get("board_approval",    ""),
                "dobm":              res.get("dobm",              ""),
                "existing_auth_cap": res.get("existing_auth_cap", ""),
                "new_auth_cap":      res.get("new_auth_cap",      ""),
                "proposed_increase": res.get("proposed_increase", ""),
                "remark_positive":   res.get("remark_positive",   ""),
                "remark_negative":   res.get("remark_negative",   ""),
                "action":            res.get("action",            "NEUTRAL"),
            })
        except Exception as e:
            print(f"Groq error for {ann.get('symbol')}: {e}")
            subj = ann.get("subject", "").lower()
            text = ann.get("_pdf_text", "").lower()
            norm = re.sub(r"\s+", "", text)
            
            is_auth = False
            if any(kw in subj for kw in ["authoris", "authoriz", "auth capital", "alteration of capital"]):
                is_auth = True
            elif ("authori" in text and "capital" in text) or ("authori" in norm and "capital" in norm):
                is_auth = True
                
            ann["is_auth_capital"] = is_auth
            fallback_data = extract_capital_fallback(text)
            ann.update(fallback_data)

    return announcements

    client = GroqClient(api_key=GROQ_API_KEY)

    SYSTEM_PROMPT = """You are a senior Indian equity market analyst.
You will receive a list of NSE/BSE corporate announcements in JSON.
Your job is to analyze each one and return structured data.
You MUST respond with ONLY a valid JSON array — no explanation, no markdown fences, no extra text."""

    batch_size = 10
    skip_groq = False
    for i in range(0, len(announcements), batch_size):
        batch = announcements[i:i + batch_size]
        
        if skip_groq:
            # Automatic fallback if rate limit was hit in previous batch
            for ann in batch:
                subj = ann.get("subject", "").lower()
                ann["is_auth_capital"]   = "authoris" in subj or "authoriz" in subj
                ann.update({k: "" for k in ["board_approval", "dobm", "existing_auth_cap", "new_auth_cap", "proposed_increase", "remark_positive", "remark_negative"]})
                ann["action"] = "Review" if ann["is_auth_capital"] else "NEUTRAL"
            continue

        batch_text = json.dumps([
            {
                "idx":     j,
                "company": a.get("company", ""),
                "subject": a.get("subject", ""),
            }
            for j, a in enumerate(batch)
        ], indent=2)

        user_prompt = f"""Analyze these corporate announcements and return a JSON array.

For EACH announcement:
1. is_auth_capital: true if announcement is about increase in Authorised/Authorized Equity Capital, else false
2. board_approval: Board approval date in DD-MM-YYYY if found in subject, else ""
3. dobm: Date of Board Meeting in DD-MM-YYYY if found, else ""
4. existing_auth_cap: Existing authorised equity capital amount (e.g. "10,00,00,000") if found, else ""
5. new_auth_cap: New/proposed authorised equity capital amount if found, else ""
6. proposed_increase: Difference (new - existing) if both found, else ""
7. remark_positive: 1 line — positive angle (expansion, fundraising flexibility, growth signal). Only if is_auth_capital=true, else ""
8. remark_negative: 1 line — risk angle (dilution risk, overleveraging). Only if is_auth_capital=true, else ""
9. action: One of — "BUY on dip", "ACCUMULATE", "WATCH", "NEUTRAL", "AVOID". Only if is_auth_capital=true, else "NEUTRAL"

Announcements:
{batch_text}

Reply ONLY with a JSON object containing an "announcements" array, like this:
{{"announcements": [{{"idx":0,"is_auth_capital":true,"board_approval":"","dobm":"","existing_auth_cap":"","new_auth_cap":"","proposed_increase":"","remark_positive":"","remark_negative":"","action":""}}]}}"""

        try:
            response = client.chat.completions.create(
                model="llama-3.1-8b-instant",
                messages=[
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user",   "content": user_prompt},
                ],
                response_format={"type": "json_object"},
                temperature=0.1,
                max_tokens=3000,
            )
            raw = response.choices[0].message.content.strip()
            # Strip accidental markdown fences
            raw = re.sub(r"^```json\s*", "", raw)
            raw = re.sub(r"^```\s*",     "", raw)
            raw = re.sub(r"\s*```$",     "", raw)

            results = json.loads(raw)
            if "announcements" in results:
                results = results["announcements"]
                
            for r in results:
                idx = r.get("idx", 0)
                if idx < len(batch):
                    batch[idx].update({
                        "is_auth_capital":   r.get("is_auth_capital",   False),
                        "board_approval":    r.get("board_approval",    ""),
                        "dobm":              r.get("dobm",              ""),
                        "existing_auth_cap": r.get("existing_auth_cap", ""),
                        "new_auth_cap":      r.get("new_auth_cap",      ""),
                        "proposed_increase": r.get("proposed_increase", ""),
                        "remark_positive":   r.get("remark_positive",   ""),
                        "remark_negative":   r.get("remark_negative",   ""),
                        "action":            r.get("action",            "NEUTRAL"),
                    })

        except Exception as e:
            print(f"Groq classification error (batch {i}): {e}")
            if "rate_limit_exceeded" in str(e).lower() or "429" in str(e):
                print("Rate limit reached. Skipping Groq for remaining batches and using keyword fallback.")
                skip_groq = True
            
            # Keyword fallback for this batch
            for ann in batch:
                subj = ann.get("subject", "").lower()
                ann["is_auth_capital"]   = "authoris" in subj or "authoriz" in subj
                ann.setdefault("board_approval",    "")
                ann.setdefault("dobm",              "")
                ann.setdefault("existing_auth_cap", "")
                ann.setdefault("new_auth_cap",      "")
                ann.setdefault("proposed_increase", "")
                ann.setdefault("remark_positive",   "")
                ann.setdefault("remark_negative",   "")
                ann.setdefault("action",            "")

    return announcements


# ─── EXCEL GENERATOR ─────────────────────────────────────────────────────────

def style_header_cell(cell, bg_color: str, font_color: str = "FFFFFF"):
    cell.font = Font(bold=True, color=font_color, size=10, name="Calibri")
    cell.fill = PatternFill("solid", fgColor=bg_color)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin = Side(style="thin", color="CCCCCC")
    cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)


def style_data_cell(cell, wrap=False):
    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=wrap)
    thin = Side(style="thin", color="E0E0E0")
    cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
    cell.font = Font(size=9, name="Calibri")


COLUMNS = [
    ("Sr.no",               8),
    ("Date of Entry",       13),
    ("Name of the Company", 28),
    ("Board Approval",      14),
    ("DOBM",                12),
    ("Exst Auth Eq Cap (INR)", 18),
    ("New Auth Eq Cap (INR)",  18),
    ("Proposed Increase (INR)",18),
    ("CMP",                  9),
    ("M Cap (in Cr)",        13),
    ("Sector",               18),
    ("Remark Positive",      28),
    ("Remark Negative",      28),
    ("Action",               14),
    ("Link",                 35),
]


def write_sheet(ws, rows: list, header_color: str, title: str):
    # Title row
    ws.merge_cells(f"A1:{get_column_letter(len(COLUMNS))}1")
    title_cell = ws["A1"]
    title_cell.value = title
    title_cell.font = Font(bold=True, size=13, color="FFFFFF", name="Calibri")
    title_cell.fill = PatternFill("solid", fgColor="1A1A2E")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
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
            # Zebra striping
            if row_idx % 2 == 0:
                cell.fill = PatternFill("solid", fgColor="F8F9FF")
            # Color-code action column
            if col_idx == 14 and value:
                v = str(value).upper()
                colors = {"BUY": "C6EFCE", "ACCUMULATE": "C6EFCE",
                          "AVOID": "FFC7CE", "NEUTRAL": "FFEB9C", "WATCH": "FFEB9C"}
                for key, clr in colors.items():
                    if key in v:
                        cell.fill = PatternFill("solid", fgColor=clr)
                        break

    # Freeze panes
    ws.freeze_panes = "A3"


def generate_excel(auth_anns: list, other_anns: list, date_str: str) -> str:
    wb = openpyxl.Workbook()

    ws_auth = wb.active
    ws_auth.title = "Auth Capital"
    write_sheet(ws_auth, auth_anns, "1565C8",
                f"Authorised Capital Increase Announcements — {date_str}")

    ws_other = wb.create_sheet("Other Announcements")
    write_sheet(ws_other, other_anns, "C85215",
                f"Other Announcements — {date_str}")

    filename = f"announcements_{date_str.replace(' ', '_').replace('-', '')}.xlsx"
    filepath = os.path.join(OUTPUT_DIR, filename)
    wb.save(filepath)
    return filename


# ─── MAIN ENDPOINT ───────────────────────────────────────────────────────────

fetch_status = {"status": "idle", "message": "", "progress": 0, "filename": ""}

# Store the last fetch date range so /announcements can filter correctly
_last_fetch_range = {"from_date": None, "to_date": None}


@app.get("/status")
async def get_status():
    # Return a lightweight status object — never embed announcements here.
    # The frontend should call GET /announcements separately once status == 'done'.
    safe = {
        "status":      fetch_status.get("status"),
        "message":     fetch_status.get("message"),
        "progress":    fetch_status.get("progress"),
        "filename":    fetch_status.get("filename"),
        "auth_count":  fetch_status.get("auth_count"),
        "other_count": fetch_status.get("other_count"),
        "total":       fetch_status.get("total"),
    }
    return safe


@app.post("/fetch")
async def trigger_fetch(req: FetchRequest, background_tasks: BackgroundTasks, api_key: str = Depends(verify_api_key)):
    if fetch_status["status"] == "running":
        raise HTTPException(status_code=409, detail="Fetch already in progress")
    today = date.today().strftime("%d-%m-%Y")
    from_date = req.from_date or today
    to_date   = req.to_date   or today
    background_tasks.add_task(run_full_pipeline, from_date, to_date)
    return {"message": "Fetch started", "from_date": from_date, "to_date": to_date}


async def run_full_pipeline(from_date: str, to_date: str):
    global _last_fetch_range
    _last_fetch_range = {"from_date": from_date, "to_date": to_date}
    global fetch_status
    try:
        fetch_status = {"status": "running", "message": "Fetching NSE announcements...", "progress": 10, "filename": ""}

        nse_anns = await fetch_nse_announcements(from_date, to_date)
        print(f"NSE fetched: {len(nse_anns)} items")
        fetch_status["message"] = f"NSE done ({len(nse_anns)} announcements). Fetching BSE..."
        fetch_status["progress"] = 25

        bse_anns = await fetch_bse_announcements(from_date, to_date)
        print(f"BSE fetched: {len(bse_anns)} items")
        fetch_status["message"] = f"BSE done ({len(bse_anns)} announcements). Running AI classification..."
        fetch_status["progress"] = 45

        all_anns = nse_anns + bse_anns
        print(f"Total announcements: {len(all_anns)}")

        # AI Classification (async)
        all_anns = await classify_and_extract_async(all_anns)
        fetch_status["message"] = "AI done. Fetching stock prices..."
        fetch_status["progress"] = 65

        # Fetch stock data for each announcement with concurrency control
        semaphore = asyncio.Semaphore(5)
        async with httpx.AsyncClient(timeout=15, follow_redirects=True) as client:
            try:
                # One warm-up for the whole batch
                fetch_status["message"] = "Warming up NSE session..."
                await client.get("https://www.nseindia.com", headers=NSE_HEADERS, timeout=10)
            except Exception as e:
                print(f"NSE Warm-up error: {e}")

            total_stocks = len(all_anns)
            for idx, ann in enumerate(all_anns):
                if idx % 10 == 0:
                    fetch_status["message"] = f"Fetching stock data ({idx}/{total_stocks})..."
                    # Progress from 45% to 85%
                    fetch_status["progress"] = 45 + int((idx / total_stocks) * 40)
                
                # Fetch stock data one by one (throttled by semaphore anyway)
                # Actually gather is faster, let's keep gather but with chunks for progress
                pass

            # Refined gather with progress updates
            chunk_size = 20
            for i in range(0, len(all_anns), chunk_size):
                chunk = all_anns[i:i + chunk_size]
                fetch_status["message"] = f"Fetching stock data ({i}/{total_stocks})..."
                fetch_status["progress"] = 45 + int((i / total_stocks) * 40)
                
                tasks = [fetch_stock_data(client, a["symbol"], a["exchange"], semaphore) for a in chunk]
                stock_results = await asyncio.gather(*tasks, return_exceptions=True)
                for ann, stock in zip(chunk, stock_results):
                    ann["_stock"] = stock if isinstance(stock, dict) else {}

        fetch_status["message"] = "Generating Excel files..."
        fetch_status["progress"] = 85

        auth_anns  = [a for a in all_anns if a.get("is_auth_capital")]
        other_anns = [a for a in all_anns if not a.get("is_auth_capital")]

        date_label = f"{from_date} to {to_date}" if from_date != to_date else from_date
        loop = asyncio.get_event_loop()
        filename = await loop.run_in_executor(None, generate_excel, auth_anns, other_anns, date_label)

        fetch_status["message"] = "Saving to database..."
        save_to_db(all_anns)

        # Store lightweight counts — DO NOT embed announcements in fetch_status.
        # The frontend fetches them via GET /announcements after status == 'done'.
        fetch_status = {
            "status":      "done",
            "message":     f"Done! {len(auth_anns)} auth capital + {len(other_anns)} other announcements.",
            "progress":    100,
            "filename":    filename,
            "auth_count":  len(auth_anns),
            "other_count": len(other_anns),
            "total":       len(auth_anns) + len(other_anns),
        }

    except Exception as e:
        fetch_status = {"status": "error", "message": str(e), "progress": 0, "filename": ""}


@app.get("/download/{filename}")
async def download_file(filename: str):
    filepath = os.path.join(OUTPUT_DIR, filename)
    if not os.path.exists(filepath):
        raise HTTPException(status_code=404, detail="File not found")
    return FileResponse(filepath, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        filename=filename)


@app.get("/announcements")
async def get_announcements():
    """Return all announcements from the DB for the last fetched date range."""
    from_date = _last_fetch_range.get("from_date")
    to_date   = _last_fetch_range.get("to_date")
    anns = get_from_db(from_date, to_date)
    return {"data": anns, "total": len(anns)}


@app.get("/health")
async def health():
    return {"status": "ok", "api_key_set": bool(GROQ_API_KEY)}

# Serve Frontend
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FRONTEND_DIR = os.path.join(os.path.dirname(BASE_DIR), "frontend")
if not os.path.exists(FRONTEND_DIR):
    FRONTEND_DIR = os.path.join(BASE_DIR, "frontend")

@app.get("/")
async def serve_frontend():
    index_path = os.path.join(FRONTEND_DIR, "index.html")
    if os.path.exists(index_path):
        return FileResponse(index_path)
    return {"error": "Frontend not found"}
