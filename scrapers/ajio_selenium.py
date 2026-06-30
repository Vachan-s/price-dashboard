import time
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By

url = "https://www.ajio.com/ten-x-you-men-all-rounder-cricket-shoes/p/469817211_limegreen"

driver = uc.Chrome(headless=False)

try:
    driver.get(url)
    time.sleep(5)

    price_el = driver.find_element(By.CSS_SELECTOR, "div.prod-sp")
    name_el  = driver.find_element(By.CSS_SELECTOR, "h1.prod-name")

    print("Name: ", name_el.text)
    print("Price:", price_el.text)

except Exception as e:
    print("Error:", e)

finally:
    driver.quit()
