import asyncio
import json
from playwright.async_api import async_playwright

SEARCH_URL = "https://www.myntra.com/ten+x+you"
RESULTS_PATH = "data/myntra_search_results.json"
ITEM_SELECTOR = "li.product-base"


async def get_text(element, selectors):
    for sel in selectors:
        el = await element.query_selector(sel)
        if el:
            text = (await el.inner_text()).strip()
            if text:
                return text
    return "Not found"


async def scrape_myntra_search(search_url: str = SEARCH_URL) -> list:
    products = []

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        page = await browser.new_page()

        try:
            await page.goto(search_url, wait_until="domcontentloaded", timeout=30000)
            await page.wait_for_selector(ITEM_SELECTOR, timeout=15000)

            previous_count = -1
            for _ in range(15):
                items = await page.query_selector_all(ITEM_SELECTOR)
                current_count = len(items)
                if current_count == previous_count:
                    break
                previous_count = current_count
                await page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
                await page.wait_for_timeout(3000)

            items = await page.query_selector_all(ITEM_SELECTOR)

            for item in items:
                anchor = await item.query_selector("a")
                href = await anchor.get_attribute("href") if anchor else await item.get_attribute("href")

                brand = await get_text(item, [".product-brand"])
                product_name = await get_text(item, [".product-product"])
                price = await get_text(item, [".product-discountedPrice", ".product-price span"])

                if href:
                    url = href if href.startswith("http") else f"https://www.myntra.com/{href.lstrip('/')}"
                else:
                    url = "Not found"

                name = f"{brand} {product_name}".strip()

                product = {
                    "name": name,
                    "price": price,
                    "url": url,
                }
                products.append(product)
                print(f"{name} | {price} | {url}")

        except Exception as e:
            print("Error:", e)

        finally:
            await browser.close()

    with open(RESULTS_PATH, "w", encoding="utf-8") as f:
        json.dump(products, f, indent=2, ensure_ascii=False)

    print(f"\nScraped {len(products)} products. Saved to {RESULTS_PATH}")

    return products


if __name__ == "__main__":
    asyncio.run(scrape_myntra_search())
