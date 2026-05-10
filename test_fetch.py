import asyncio
from backend.main import fetch_nse_announcements, fetch_bse_announcements, is_exchange_auth_category

async def main():
    date_str = "08-05-2026"
    date_to = "09-05-2026"
    print("Fetching NSE...")
    nse = await fetch_nse_announcements(date_str, date_to)
    print(f"NSE count: {len(nse)}")
    if nse:
        print("NSE keys:", nse[0]['raw'].keys())
        # Print a few raw items
        for i in range(min(5, len(nse))):
            print(nse[i]['raw'])
    
    print("Fetching BSE...")
    bse = await fetch_bse_announcements(date_str, date_to)
    print(f"BSE count: {len(bse)}")
    if bse:
        print("BSE keys:", bse[0]['raw'].keys())
        for i in range(min(5, len(bse))):
            print(bse[i]['raw'])

if __name__ == "__main__":
    asyncio.run(main())
