import asyncio
from backend.main import fetch_nse_announcements, fetch_bse_announcements

async def main():
    date_str = "09-05-2026"
    nse = await fetch_nse_announcements(date_str, date_str)
    bse = await fetch_bse_announcements(date_str, date_str)
    
    keywords = ["authorised", "authorized", "capital", "alteration"]
    print("\n--- Potential Auth Capital Announcements (ignoring blacklist) ---")
    for a in nse + bse:
        subj = a.get("subject", "").lower()
        if "capital" in subj or "auth" in subj:
            print(f"[{a['exchange']}] {a['company']} - {a['subject']}")

if __name__ == "__main__":
    asyncio.run(main())
