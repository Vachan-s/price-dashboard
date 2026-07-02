import asyncio
from playwright.async_api import async_playwright

_USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"


async def scrape_myntra_price_with_page(page, url: str) -> dict:
    await page.set_extra_http_headers({"User-Agent": _USER_AGENT})
    try:
        await page.goto(url, wait_until="domcontentloaded", timeout=30000)
        await page.wait_for_timeout(1500)

        price_el = await page.query_selector("span.pdp-price strong")
        price = await price_el.inner_text() if price_el else "Not found"

        name_el = await page.query_selector("h1.pdp-name")
        name = await name_el.inner_text() if name_el else "Not found"

        return {"url": url, "name": name, "price": price, "status": "success"}

    except Exception as e:
        return {"url": url, "name": None, "price": None, "status": f"error: {str(e)}"}


async def scrape_myntra_price(url: str) -> dict:
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page()
        try:
            return await scrape_myntra_price_with_page(page, url)
        finally:
            await browser.close()


if __name__ == "__main__":
    url = "https://www.myntra.com/sports-shoes/ten+x+you/ten-x-you-unisex-switch-og-20-lightweight-cricket-shoes/40099353/buy"
    result = asyncio.run(scrape_myntra_price(url))
    print(result)
