import asyncio
import json
import os
import sys
import xml.etree.ElementTree as ET

import requests
from playwright.async_api import async_playwright

# Ensure project root is on sys.path so imports work from any CWD
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

from scrapers.tenxyou import scrape_tenxyou_price_with_page

SITEMAP_URL  = "https://tenxyou.com/sitemap.xml"
RESULTS_PATH = os.path.join(_ROOT, "data", "tenxyou_search_results.json")


def _extract_locs(xml_bytes: bytes) -> list[str]:
    root = ET.fromstring(xml_bytes)
    return [
        elem.text.strip()
        for elem in root.iter()
        if (elem.tag.endswith("}loc") or elem.tag == "loc") and elem.text
    ]


def get_product_urls() -> list[str]:
    """Fetch sitemap.xml and return all /product/ URLs.
    Follows one level of sitemap-index indirection if needed."""
    resp = requests.get(SITEMAP_URL, timeout=30)
    resp.raise_for_status()
    all_locs = _extract_locs(resp.content)

    product_urls = [u for u in all_locs if "/product/" in u]

    # If the main sitemap is a sitemap index, fetch sub-sitemaps
    if not product_urls:
        sub_sitemaps = [u for u in all_locs if "sitemap" in u.lower()]
        for sub_url in sub_sitemaps:
            try:
                sub = requests.get(sub_url, timeout=30)
                sub.raise_for_status()
                product_urls.extend(u for u in _extract_locs(sub.content) if "/product/" in u)
            except Exception as e:
                print(f"Warning: could not fetch sub-sitemap {sub_url}: {e}")

    # Deduplicate while preserving order
    return list(dict.fromkeys(product_urls))


def _product_key(url: str) -> str:
    """Collapse a product URL to a base-product key by dropping the trailing
    color words from its slug (colors are 1-2 hyphenated words, e.g.
    'royal-black', 'vermillion') — used to group color variants of the same
    product so we only scrape one representative per product."""
    slug = url.split("?")[0].rstrip("/").split("/")[-1]
    parts = slug.split("-")
    if len(parts) > 2:
        parts = parts[:-2]
    elif len(parts) > 1:
        parts = parts[:-1]
    return "-".join(parts)


def dedupe_by_product(urls: list) -> list:
    """Keep only the first URL seen for each unique product key."""
    seen = set()
    deduped = []
    for url in urls:
        key = _product_key(url)
        if key not in seen:
            seen.add(key)
            deduped.append(url)
    return deduped


async def scrape_tenxyou_search() -> list:
    urls = get_product_urls()
    print(f"Found {len(urls)} product URLs in sitemap\n")

    deduped_urls = dedupe_by_product(urls)
    print(f"{len(urls)} URLs → {len(deduped_urls)} unique products after deduplication\n")

    products = []
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        try:
            for url in deduped_urls:
                page = await browser.new_page()
                try:
                    result = await scrape_tenxyou_price_with_page(page, url)
                    product = {
                        "name":  result["name"],
                        "price": result["price"],
                        "url":   url,
                    }
                    products.append(product)
                    print(f"{result['name']} | {result['price']} | {url}")
                except Exception as e:
                    print(f"Error scraping {url}: {e}")
                finally:
                    await page.close()
        finally:
            await browser.close()

    with open(RESULTS_PATH, "w", encoding="utf-8") as f:
        json.dump(products, f, indent=2, ensure_ascii=False)

    print(f"\nScraped {len(products)} products. Saved to {RESULTS_PATH}")
    return products


if __name__ == "__main__":
    asyncio.run(scrape_tenxyou_search())
