import asyncio
from backend.main import fetch_nse_announcements, fetch_bse_announcements

async def main():
    date_str = "08-05-2026"
    
    print("Fetching BSE...")
    bse = await fetch_bse_announcements(date_str, date_str)
    print(f"BSE count: {len(bse)}")
    if bse:
        print("BSE keys:", bse[0]['raw'].keys())
        for i in range(min(3, len(bse))):
            print(bse[i]['raw'])

if __name__ == "__main__":
    asyncio.run(main())
