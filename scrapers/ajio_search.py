import asyncio
import json
from playwright.async_api import async_playwright

SEARCH_URL = "https://www.ajio.com/search/?text=ten+x+you"
RESULTS_PATH = "data/ajio_search_results.json"


async def get_text(element, selectors):
    for sel in selectors:
        el = await element.query_selector(sel)
        if el:
            text = (await el.inner_text()).strip()
            if text:
                return text
    return "Not found"


async def scrape_ajio_search(search_url: str = SEARCH_URL) -> list:
    products = []

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        page = await browser.new_page()

        try:
            await page.goto(search_url, wait_until="domcontentloaded", timeout=30000)

            item_selector = ".item"
            try:
                await page.wait_for_selector(item_selector, timeout=15000)
            except Exception:
                item_selector = '[class*="rilrtl-products-list"] li'
                await page.wait_for_selector(item_selector, timeout=15000)

            previous_count = -1
            for _ in range(15):
                items = await page.query_selector_all(item_selector)
                current_count = len(items)
                if current_count == previous_count:
                    break
                previous_count = current_count
                await page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
                await page.wait_for_timeout(3000)

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

        finally:
            await browser.close()

    with open(RESULTS_PATH, "w", encoding="utf-8") as f:
        json.dump(products, f, indent=2, ensure_ascii=False)

    print(f"\nScraped {len(products)} products. Saved to {RESULTS_PATH}")

    return products


if __name__ == "__main__":
    asyncio.run(scrape_ajio_search())
