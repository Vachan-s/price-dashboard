import asyncio
import json
from playwright.async_api import async_playwright

SEARCH_URLS = [
    "https://www.ajio.com/search/?text=ten+x+you&gender=Men",
    "https://www.ajio.com/search/?text=ten+x+you&gender=Women",
]
RESULTS_PATH = "data/ajio_search_results.json"

MAX_SCROLLS = 30

LOAD_MORE_SELECTORS = [
    "button.btn-load-more",
    '[class*="load-more"] button',
    '[class*="loadMore"]',
]

SCROLL_CONTAINER_SELECTORS = [
    '[class*="rilrtl-products-list"]',
    '[class*="search-cntr"]',
    "main",
]


async def _click_load_more(page) -> bool:
    """Click a 'Load More' button if one is present. Returns True if clicked."""
    for sel in LOAD_MORE_SELECTORS:
        btn = await page.query_selector(sel)
        if btn:
            try:
                await btn.click()
                return True
            except Exception:
                continue
    return False


async def _get_scroll_container(page):
    """Find the actual scrollable product container; None means scroll the window."""
    for sel in SCROLL_CONTAINER_SELECTORS:
        el = await page.query_selector(sel)
        if el:
            return el
    return None


async def _scroll_step(page, container) -> None:
    """Scroll one viewport height — the product container if found, else the window."""
    if container:
        await container.evaluate("el => el.scrollBy(0, window.innerHeight)")
    else:
        await page.evaluate("window.scrollBy(0, window.innerHeight)")


async def get_text(element, selectors):
    for sel in selectors:
        el = await element.query_selector(sel)
        if el:
            text = (await el.inner_text()).strip()
            if text:
                return text
    return "Not found"


async def _scrape_one_url(page, search_url: str) -> list:
    products = []

    try:
        await page.goto(search_url, wait_until="domcontentloaded", timeout=30000)

        item_selector = ".item"
        try:
            await page.wait_for_selector(item_selector, timeout=15000)
        except Exception:
            item_selector = '[class*="rilrtl-products-list"] li'
            await page.wait_for_selector(item_selector, timeout=15000)

        scroll_container = await _get_scroll_container(page)

        stale_streak = 0
        previous_count = len(await page.query_selector_all(item_selector))
        for i in range(MAX_SCROLLS):
            await _scroll_step(page, scroll_container)
            await page.wait_for_timeout(1000)
            await _click_load_more(page)

            current_count = len(await page.query_selector_all(item_selector))
            print(f"[SCROLL {i + 1}/{MAX_SCROLLS}] product count: {current_count}", flush=True)

            if current_count > previous_count:
                stale_streak = 0
            else:
                stale_streak += 1
                if stale_streak >= 3:
                    break
            previous_count = current_count

        items = await page.query_selector_all(item_selector)

        for item in items:
            anchor = await item.query_selector("a")
            href = await anchor.get_attribute("href") if anchor else await item.get_attribute("href")

            brand = await get_text(item, [".brand", ".nameCls"])
            description = await get_text(item, [".desc", ".nameCls"])
            price = await get_text(item, [".prod-sp", ".price strong"])

            if href:
                url = href if href.startswith("http") else f"https://www.ajio.com{href}"
            else:
                url = "Not found"

            name = f"{brand} {description}".strip()

            product = {
                "name": name,
                "price": price,
                "url": url,
            }
            products.append(product)
            print(f"{name} | {price} | {url}")

    except Exception as e:
        print("Error:", e)

    return products


async def scrape_ajio_search(search_urls: list = SEARCH_URLS) -> list:
    all_products = []

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)

        for search_url in search_urls:
            page = await browser.new_page()
            try:
                all_products.extend(await _scrape_one_url(page, search_url))
            finally:
                await page.close()

        await browser.close()

    seen = set()
    products = []
    for p in all_products:
        if p["url"] not in seen:
            seen.add(p["url"])
            products.append(p)

    with open(RESULTS_PATH, "w", encoding="utf-8") as f:
        json.dump(products, f, indent=2, ensure_ascii=False)

    print(f"\nScraped {len(all_products)} raw, {len(products)} unique products. Saved to {RESULTS_PATH}")

    return products


if __name__ == "__main__":
    asyncio.run(scrape_ajio_search())
