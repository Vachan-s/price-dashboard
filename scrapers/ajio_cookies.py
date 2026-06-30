import asyncio
from playwright.async_api import async_playwright

URL_HOME    = "https://www.ajio.com"
URL_PRODUCT = "https://www.ajio.com/ten-x-you-men-all-rounder-cricket-shoes/p/469817211_limegreen"


async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=False,
            args=[
                "--disable-blink-features=AutomationControlled",
                "--no-sandbox",
                "--disable-dev-shm-usage",
            ],
        )

        context = await browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/149.0.0.0 Safari/537.36"
            ),
            viewport={"width": 1920, "height": 1080},
            locale="en-IN",
            timezone_id="Asia/Kolkata",
        )

        page = await context.new_page()

        await page.evaluate(
            "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
        )

        try:
            await page.goto(URL_HOME, wait_until="domcontentloaded")
            await page.wait_for_timeout(4000)

            await page.goto(URL_PRODUCT, wait_until="domcontentloaded")
            await page.wait_for_timeout(6000)

            price_el = await page.query_selector("div.prod-sp")
            name_el  = await page.query_selector("h1.prod-name")

            price = await price_el.inner_text() if price_el else "Not found"
            name  = await name_el.inner_text()  if name_el  else "Not found"

            print("Name: ", name)
            print("Price:", price)

        except Exception as e:
            print("Error:", e)

        finally:
            await browser.close()


if __name__ == "__main__":
    asyncio.run(main())
