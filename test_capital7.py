import asyncio
from backend.main import fetch_nse_announcements, fetch_bse_announcements

async def main():
    date_str = "07-05-2026"
    nse = await fetch_nse_announcements(date_str, date_str)
    bse = await fetch_bse_announcements(date_str, date_str)
    
    keywords = ["authorised", "authorized", "capital", "alteration"]
    print("\n--- Potential Auth Capital on 07-05-2026 ---")
    for a in nse + bse:
        subj = a.get("subject", "").lower()
        subcat = (a.get("raw", {}).get("SUBCATNAME") or "").lower()
        if "capital" in subcat or "alter" in subcat:
            print(f"[BSE CAT MATCH] {a['company']} - {a['subject']} (Cat: {subcat})")
        if "increase in auth" in subj or "authorised capital" in subj or "authorized capital" in subj:
            print(f"[SUBJ MATCH] {a['company']} - {a['subject']} (Cat: {subcat})")

if __name__ == "__main__":
    asyncio.run(main())
