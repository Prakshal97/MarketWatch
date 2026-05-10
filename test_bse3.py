import asyncio
import httpx

BSE_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Accept": "application/json",
    "Referer": "https://www.bseindia.com/",
}

async def test_bse():
    url = (
        "https://api.bseindia.com/BseIndiaAPI/api/AnnSubCategoryGetData/w"
        "?pageno=1&strCat=-1&strPrevDate=20260509&strScrip=&strSearch=P"
        "&strToDate=20260509&strType=C&subcategory=-1"
    )
    async with httpx.AsyncClient(timeout=30) as client:
        resp = await client.get(url, headers=BSE_HEADERS)
        print("Status:", resp.status_code)
        if resp.status_code == 200:
            print("Content length:", len(resp.text))
            print("Content:", resp.text[:500])
            
asyncio.run(test_bse())
