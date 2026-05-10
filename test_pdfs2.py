import asyncio
from backend.main import fetch_nse_announcements, fetch_bse_announcements, download_pdf_text

async def main():
    date_str = "09-05-2026"
    nse = await fetch_nse_announcements(date_str, date_str)
    bse = await fetch_bse_announcements(date_str, date_str)
    
    print(f"Total: {len(nse) + len(bse)}")
    
    tasks = []
    sem = asyncio.Semaphore(10)
    
    async def check_pdf(a):
        async with sem:
            url = a.get("link", "")
            if a.get("exchange") == "NSE":
                url = a.get("raw", {}).get("attchmntFile", "")
            if not url or "pdf" not in url.lower(): return
            text = await download_pdf_text(url, a.get("exchange", "NSE"))
            if "authorised share capital" in text.lower() or "authorized share capital" in text.lower() or "authorised capital" in text.lower():
                print(f"[FOUND IN PDF] {a['exchange']} - {a['company']} - {a['subject']}")
                print(f"URL: {url}")
                
    for a in nse + bse:
        tasks.append(check_pdf(a))
        
    await asyncio.gather(*tasks)

if __name__ == "__main__":
    asyncio.run(main())
