import asyncio
from playwright.async_api import async_playwright

async def test():
    urls = [
        "https://www.ajio.com/ten-x-you-men-relaxed-fit-running-shorts/p/469817545_black",
        "https://www.ajio.com/ten-x-you-men-regular-fit-shorts/p/469817548_black",
        "https://www.ajio.com/ten-x-you-women-relaxed-fit-running-shorts/p/469817551_black",
    ]
    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=False,
            args=["--disable-blink-features=AutomationControlled"]
        )
        context = await browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
            viewport={"width": 1920, "height": 1080}
        )
        for url in urls:
            page = await context.new_page()
            await page.goto(url, wait_until="domcontentloaded", timeout=30000)
            await page.wait_for_timeout(3000)
            content = await page.content()
            if "access denied" in content.lower():
                print(f"BLOCKED: {url}")
            else:
                el = await page.query_selector("div.prod-sp")
                price = await el.inner_text() if el else "Not found"
                print(f"PRICE: {price} | {url}")
            await page.close()
        await browser.close()

asyncio.run(test())
