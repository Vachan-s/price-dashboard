import asyncio
from playwright.async_api import async_playwright

_USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"

# Myntra shows a "SOLD OUT" button in place of "ADD TO BAG" when a product is
# out of stock. Checked against the buy-button element (and its immediate
# container, in case the label sits next to rather than inside the button),
# not the whole page, to avoid false positives from unrelated page text.
_BUY_BUTTON_SELECTORS = [".pdp-add-to-bag", '[class*="add-to-bag"]']
_OOS_MARKERS = ["sold out", "out of stock"]


async def _is_myntra_oos(page) -> bool:
    """Check the buy-button area (not the whole page) for a sold-out / out-of-
    stock indicator, to avoid false positives from unrelated page text."""
    for sel in _BUY_BUTTON_SELECTORS:
        el = await page.query_selector(sel)
        if not el:
            continue
        text = (await el.inner_text()).strip().lower()
        if any(marker in text for marker in _OOS_MARKERS):
            return True
        try:
            parent_handle = await el.evaluate_handle("el => el.parentElement")
            parent = parent_handle.as_element()
            if parent:
                parent_text = (await parent.inner_text()).strip().lower()
                if any(marker in parent_text for marker in _OOS_MARKERS):
                    return True
        except Exception:
            pass
    return False


async def scrape_myntra_price_with_page(page, url: str) -> dict:
    await page.set_extra_http_headers({"User-Agent": _USER_AGENT})
    try:
        await page.goto(url, wait_until="domcontentloaded", timeout=60000)
        await page.wait_for_timeout(3000)

        price_el = await page.query_selector("span.pdp-price strong")
        price = await price_el.inner_text() if price_el else "Not found"

        name_el = await page.query_selector("h1.pdp-name")
        name = await name_el.inner_text() if name_el else "Not found"

        if await _is_myntra_oos(page):
            return {"url": url, "name": name, "price": None, "status": "unavailable"}

        return {"url": url, "name": name, "price": price, "status": "success"}

    except Exception as e:
        return {"url": url, "name": None, "price": None, "status": f"error: {str(e)}"}


async def scrape_myntra_price(url: str) -> dict:
    async with async_playwright() as p:
        # headless=False: Myntra hangs/times out on headless Chrome navigations
        # (confirmed — headed mode succeeds in <1s, headless hangs to timeout).
        browser = await p.chromium.launch(headless=False, args=["--disable-http2"])
        page = await browser.new_page()
        try:
            return await scrape_myntra_price_with_page(page, url)
        finally:
            await browser.close()


if __name__ == "__main__":
    url = "https://www.myntra.com/sports-shoes/ten+x+you/ten-x-you-unisex-aeonic-recovery-running-shoes-with-cushioned-arch-support/39048046/buy"
    result = asyncio.run(scrape_myntra_price(url))
    print(result)
