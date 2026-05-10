import asyncio
from backend.main import fetch_bse_announcements

async def main():
    date_str = "09-05-2026"
    bse = await fetch_bse_announcements(date_str, date_str)
    subcats = set()
    for b in bse:
        raw = b.get("raw", {})
        subcat = str(raw.get("SUBCATNAME") or raw.get("SUBCATEGORYNAME") or "").strip().lower()
        subcats.add(subcat)
    print("Unique BSE subcategories on 09-05-2026:")
    for s in sorted(list(subcats)):
        print(f"- {s}")

if __name__ == "__main__":
    asyncio.run(main())
