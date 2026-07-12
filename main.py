from include.functions import setup_chrome_webdriver, extract_best_sellers, export_to_excel
import logging
import time

_logger = logging.getLogger(__name__)

MAX_ITEMS = 12
TARGET_URL = "https://www.amazon.com.br"
if __name__ == "__main__":
    driver = setup_chrome_webdriver(headless=False)
    try:
        results = extract_best_sellers(driver, TARGET_URL, MAX_ITEMS=10)

        if results:
            _logger.info(f"Found {len(results)} best sellers")
            export_to_excel(results, filename="Top_Mais_Vendidos_Amazon.xlsx")
        else:
            _logger.warning("No items were extracted.")
    finally:
        time.sleep(3)
        driver.quit()