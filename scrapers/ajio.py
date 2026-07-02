import asyncio
from playwright.async_api import async_playwright
from playwright_stealth import Stealth

async def scrape_ajio_price(url: str) -> dict:
    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=True,
            channel="chrome",
            args=["--disable-blink-features=AutomationControlled"]
        )
        page = await browser.new_page(viewport={"width": 1366, "height": 768})

        Stealth().apply_stealth_sync(page)

        await page.evaluate("() => delete navigator.__proto__.webdriver")

        await page.set_extra_http_headers({
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            "Accept-Language": "en-US,en;q=0.9",
            "Accept-Encoding": "gzip, deflate, br"
        })

        try:
            await page.goto(url, wait_until="domcontentloaded", timeout=30000)
            await page.wait_for_timeout(1500)

            price_el = await page.query_selector("div.prod-sp")
            price = await price_el.inner_text() if price_el else "Not found"

            name_el = await page.query_selector("h1.prod-name")
            name = await name_el.inner_text() if name_el else "Not found"

            return {
                "url": url,
                "name": name,
                "price": price,
                "status": "success"
            }

        except Exception as e:
            return {
                "url": url,
                "name": None,
                "price": None,
                "status": f"error: {str(e)}"
            }

        finally:
            await browser.close()


if __name__ == "__main__":
    url = "https://www.ajio.com/ten-x-you-men-all-rounder-cricket-shoes/p/469817211_limegreen"
    result = asyncio.run(scrape_ajio_price(url))
    print(result)
