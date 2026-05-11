# Diagnostic: Tests the full classification pipeline step-by-step.
# Run from the nse-bse-tracker directory:
#   python diagnose_classifier.py

import asyncio, json, os, re, sys
from datetime import datetime, timedelta
import httpx
from dotenv import load_dotenv

load_dotenv("backend/.env")
GROQ_API_KEY = os.getenv("GROQ_API_KEY", "")

BSE_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Accept": "application/json",
    "Referer": "https://www.bseindia.com/",
}

AUTH_CAPITAL_KEYWORDS = [
    "authorised capital", "authorized capital",
    "increase in authorised", "increase in authorized",
    "alteration of capital clause", "increase of capital clause",
    "memorandum of association", "moa alteration",
    "increase in auth", "raising of authorised",
    "change in authorised capital", "change in authorized capital",
    "alteration in capital clause", "sub-division", "subdivision",
    "increase in the authorised", "increase in the authorized",
]

AUTH_CAPITAL_BLACKLIST = [
    "financial results", "dividend", "buyback", "merger", "amalgamation",
    "acquisition", "change in management", "resignation",
    "loss of share certificate", "trading window", "credit rating",
    "press release", "investor presentation", "analyst meet",
    "debentures", "rights issue", "esop", "bonus issue",
]

BSE_AUTH_SUBCAT_KEYWORDS = [
    "alteration of capital", "increase in authorised capital",
    "increase in authorized capital", "sub-division of shares",
    "subdivision of shares", "capital change",
]


def safe_subject(ann):
    """Always return subject as a string."""
    subj = ann.get("subject", "")
    if isinstance(subj, dict):
        return str(list(subj.values())[0]) if subj else ""
    return str(subj) if subj else ""


def is_auth_capital_keyword_match(subject: str) -> bool:
    subj = subject.lower().strip()
    if any(bl in subj for bl in AUTH_CAPITAL_BLACKLIST):
        return False
    return any(kw in subj for kw in AUTH_CAPITAL_KEYWORDS)


def is_exchange_auth_category(ann: dict) -> bool:
    raw = ann.get("raw", {})
    exchange = ann.get("exchange", "")
    if exchange == "BSE":
        subcat_raw = (
            raw.get("SUBCATEGORYNAME") or raw.get("SUBCATNAME")
            or raw.get("SubCategoryName") or raw.get("subcategoryname") or ""
        )
        subcat = str(subcat_raw).strip().lower()
        return any(kw in subcat for kw in BSE_AUTH_SUBCAT_KEYWORDS)
    return False


async def fetch_bse_sample(days_back: int = 7) -> list:
    announcements = []
    end_dt   = datetime.now()
    start_dt = end_dt - timedelta(days=days_back)

    async with httpx.AsyncClient(timeout=30, follow_redirects=True) as client:
        current_dt = start_dt
        while current_dt <= end_dt:
            d_str = current_dt.strftime("%Y%m%d")
            pageno = 1
            while True:
                url = (
                    f"https://api.bseindia.com/BseIndiaAPI/api/AnnSubCategoryGetData/w"
                    f"?pageno={pageno}&strCat=-1&strPrevDate={d_str}&strScrip=&strSearch=P"
                    f"&strToDate={d_str}&strType=C&subcategory=-1"
                )
                try:
                    resp = await client.get(url, headers=BSE_HEADERS)
                    if resp.status_code == 200:
                        data  = resp.json()
                        items = data.get("Table", data if isinstance(data, list) else [])
                        if not items:
                            break
                        for item in items:
                            scrip  = item.get("SCRIP_CD", "")
                            attach = item.get("ATTACHMENTNAME", "")
                            announcements.append({
                                "exchange": "BSE",
                                "company":  item.get("SLONGNAME", ""),
                                "symbol":   str(scrip),
                                "subject":  item.get("HEADLINE", ""),
                                "date":     item.get("NEWS_DT", ""),
                                "link":     f"https://www.bseindia.com/xml-data/corpfiling/AttachHis/{attach}" if attach else "",
                                "raw":      item,
                            })
                        if len(items) < 50:
                            break
                        pageno += 1
                    else:
                        break
                except Exception as e:
                    print(f"  Fetch error {d_str} page {pageno}: {e}")
                    break
            current_dt += timedelta(days=1)

    return announcements


def test_groq_directly():
    print("\n" + "="*60)
    print("STEP 1: Testing Groq API key directly")
    print("="*60)

    if not GROQ_API_KEY:
        print("  FAIL: GROQ_API_KEY is empty in .env!")
        return False

    print(f"  Key prefix: {GROQ_API_KEY[:12]}...")

    try:
        from groq import Groq as GroqClient
        client = GroqClient(api_key=GROQ_API_KEY)
        response = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[{"role": "user", "content": 'Reply with exactly: {"ok": true}'}],
            temperature=0.0,
            max_tokens=20,
        )
        text = response.choices[0].message.content.strip()
        print(f"  OK: Groq responded: {text}")
        return True
    except Exception as e:
        print(f"  FAIL: Groq API error: {e}")
        return False


async def main():
    print("\n" + "="*60)
    print("NSE/BSE Classifier Diagnostic")
    print("="*60)

    groq_ok = test_groq_directly()

    # ── Fetch real BSE data ──────────────────────────────────────
    print("\n" + "="*60)
    print("STEP 2: Fetching last 7 days of BSE announcements")
    print("="*60)
    anns = await fetch_bse_sample(days_back=7)
    print(f"  Total BSE announcements fetched: {len(anns)}")

    if not anns:
        print("  FAIL: No announcements fetched")
        return

    # ── Audit SUBCATNAME values ──────────────────────────────────
    print("\n" + "="*60)
    print("STEP 3: Auditing SUBCATNAME field values (unique sample)")
    print("="*60)
    subcat_values = {}
    for ann in anns:
        raw = ann.get("raw", {})
        val = (
            raw.get("SUBCATEGORYNAME") or raw.get("SUBCATNAME")
            or raw.get("SubCategoryName") or ""
        )
        val = str(val).strip()
        subcat_values[val] = subcat_values.get(val, 0) + 1

    # Print all unique SUBCATNAME values sorted by count
    sorted_subcats = sorted(subcat_values.items(), key=lambda x: -x[1])
    print(f"  Found {len(sorted_subcats)} unique SUBCATNAME values:")
    for val, count in sorted_subcats[:40]:
        marker = " <-- WOULD MATCH" if any(kw in val.lower() for kw in BSE_AUTH_SUBCAT_KEYWORDS) else ""
        print(f"    [{count:4d}x] '{val}'{marker}")

    # ── Keyword + exchange checks ────────────────────────────────
    print("\n" + "="*60)
    print("STEP 4: Keyword and exchange category classification")
    print("="*60)

    keyword_hits  = []
    exchange_hits = []
    blacklist_kills = []

    for ann in anns:
        subj = safe_subject(ann)
        ann["subject"] = subj  # normalize so Groq call won't crash

        kw_hit = is_auth_capital_keyword_match(subj)
        ex_hit = is_exchange_auth_category(ann)

        wl_hit = any(kw in subj.lower() for kw in AUTH_CAPITAL_KEYWORDS)
        bl_hit = any(bl in subj.lower() for bl in AUTH_CAPITAL_BLACKLIST)

        if ex_hit:
            exchange_hits.append(ann)
        if kw_hit:
            keyword_hits.append(ann)
        if wl_hit and bl_hit and not kw_hit:
            blacklist_kills.append({
                "company": ann["company"],
                "subject": subj,
                "bl_match": [bl for bl in AUTH_CAPITAL_BLACKLIST if bl in subj.lower()],
                "wl_match": [kw for kw in AUTH_CAPITAL_KEYWORDS  if kw in subj.lower()],
            })

    print(f"  Exchange-category hits : {len(exchange_hits)}")
    print(f"  Keyword hits           : {len(keyword_hits)}")

    if blacklist_kills:
        print(f"\n  WARNING: BLACKLIST KILLED {len(blacklist_kills)} auth-capital-like subjects:")
        for item in blacklist_kills:
            print(f"    Company : {item['company']}")
            print(f"    Subject : {item['subject'][:90]}")
            print(f"    WL match: {item['wl_match']}")
            print(f"    BL match: {item['bl_match']}")
            print()

    if exchange_hits:
        print("\n  Exchange-category hits:")
        for ann in exchange_hits[:10]:
            raw    = ann.get("raw", {})
            subcat = raw.get("SUBCATEGORYNAME") or raw.get("SUBCATNAME") or "?"
            print(f"    [{ann['date'][:10]}] {ann['company']}: {ann['subject'][:80]}")
            print(f"      subcat='{subcat}'")

    if keyword_hits:
        print("\n  Keyword hits:")
        for ann in keyword_hits:
            print(f"    [{ann['date'][:10]}] {ann['company']}: {ann['subject'][:90]}")

    # ── Test Groq on actual candidates ──────────────────────────
    candidates = exchange_hits + keyword_hits
    # deduplicate by subject
    seen = set()
    unique_candidates = []
    for a in candidates:
        key = a.get("company","") + a.get("subject","")
        if key not in seen:
            seen.add(key)
            unique_candidates.append(a)

    print("\n" + "="*60)
    print(f"STEP 5: Sending {len(unique_candidates)} candidates to Groq")
    print("="*60)

    if not unique_candidates:
        print("  FAIL: No candidates -- Groq receives 0 items, so 0 API calls.")
        print("  Searching ALL subjects for partial capital-related words...")
        found = []
        for ann in anns:
            subj = safe_subject(ann).lower()
            if any(w in subj for w in ["authoris", "authoriz", "alteration", "capital"]):
                found.append(ann)
        print(f"  Found {len(found)} announcements with capital-related words:")
        for ann in found[:30]:
            print(f"    [{ann['date'][:10]}] {ann['company']}: {safe_subject(ann)[:90]}")
        return

    if not groq_ok:
        print("  SKIP: Groq key failed above")
        return

    from groq import Groq as GroqClient
    client = GroqClient(api_key=GROQ_API_KEY)
    batch  = unique_candidates[:5]
    batch_text = json.dumps(
        [{"idx": j, "company": a["company"], "subject": a["subject"]} for j, a in enumerate(batch)],
        indent=2
    )
    prompt = f"""These announcements are pre-filtered as LIKELY about Increase in Authorised Equity Capital.
Confirm each and set is_auth_capital=true/false.
Reply ONLY with a JSON array: [{{"idx":0,"is_auth_capital":true}}]

Announcements:
{batch_text}"""

    try:
        response = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.0,
            max_tokens=500,
        )
        raw = response.choices[0].message.content.strip()
        print(f"  Groq AI response:\n{raw}")
        print("\n  SUCCESS: Groq is working and classifying correctly.")
    except Exception as e:
        print(f"  FAIL: Groq call failed: {e}")


if __name__ == "__main__":
    asyncio.run(main())
