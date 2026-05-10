import asyncio
from backend.main import fetch_nse_announcements, fetch_bse_announcements, classify_with_ai

async def main():
    date_str = "09-05-2026"
    nse = await fetch_nse_announcements(date_str, date_str)
    bse = await fetch_bse_announcements(date_str, date_str)
    all_anns = nse + bse
    print(f"Total: {len(all_anns)}")
    
    # Run classification
    classified = classify_with_ai(all_anns)
    
    auth_anns = [a for a in classified if a.get("is_auth_capital")]
    print(f"Auth Capital count: {len(auth_anns)}")
    
if __name__ == "__main__":
    asyncio.run(main())
