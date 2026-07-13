import asyncio
from playwright.async_api import async_playwright

_USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"

# TenXYou swaps its "Buy Now" / "Add to Kit" buttons for a "Notify Me" button
# when a product is out of stock, rather than showing literal "sold out" /
# "out of stock" text anywhere on the page (confirmed against a real OOS
# listing) — checked alongside the literal phrases as a safety net.
_OOS_BUTTON_MARKERS = ["notify me", "sold out", "out of stock"]


async def _is_tenxyou_oos(page) -> bool:
    """Check the buy-button area (not the whole page) for an out-of-stock
    indicator, to avoid false positives from unrelated page text."""
    buttons = await page.query_selector_all("button")
    for b in buttons:
        text = (await b.inner_text()).strip().lower()
        if any(marker in text for marker in _OOS_BUTTON_MARKERS):
            return True
    return False


async def scrape_tenxyou_price_with_page(page, url: str) -> dict:
    await page.set_extra_http_headers({"User-Agent": _USER_AGENT})
    try:
        await page.goto(url, wait_until="domcontentloaded", timeout=60000)

        try:
            await page.wait_for_selector("span.price", timeout=15000)
        except Exception:
            return {"url": url, "name": None, "price": "Not found", "status": "error"}

        price_el = await page.query_selector("span.price")
        price = await price_el.inner_text() if price_el else "Not found"

        name_el = await page.query_selector("h1")
        name = await name_el.inner_text() if name_el else "Not found"

        if await _is_tenxyou_oos(page):
            return {"url": url, "name": name, "price": None, "status": "unavailable"}

        return {"url": url, "name": name, "price": price, "status": "success"}

    except Exception as e:
        return {"url": url, "name": None, "price": None, "status": f"error: {str(e)}"}


async def scrape_tenxyou_price(url: str) -> dict:
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page()
        try:
            return await scrape_tenxyou_price_with_page(page, url)
        finally:
            await browser.close()


if __name__ == "__main__":
    url = "https://tenxyou.com/product/unisex-humvee-clogs-ranger-green"
    result = asyncio.run(scrape_tenxyou_price(url))
    print(result)
