import concurrent.futures
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, ElementClickInterceptedException, StaleElementReferenceException
import pandas as pd
from datetime import datetime
import os
import time

file_path = r"D:\Projects\NSC Pre Open Scraper\Book1_parallel.xlsx"
url = "https://www.nseindia.com/market-data/pre-open-market-cm-and-emerge-market"


def get_all_stock_row_indices(driver):
    # Get all rows containing a stock symbol anchor
    rows = driver.find_elements(By.CSS_SELECTOR, "tr")
    valid_indices = [i for i, row in enumerate(rows) if row.find_elements(By.CSS_SELECTOR, "a.symbol-word-break")]
    return valid_indices


def scrape_batch(row_indices):
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--disable-extensions")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64 x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/92.0.451.100 Safari/537.36"
    )

    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 10)
    driver.get(url)

    # Select 'Securities in F&O'
    wait.until(EC.presence_of_element_located((By.ID, "sel-Pre-Open-Market")))
    select_element = driver.find_element(By.ID, 'sel-Pre-Open-Market')
    select_category = Select(select_element)
    select_category.select_by_visible_text('Securities in F&O')

    # Short delay to allow table update
    time.sleep(2)

    records = []

    for idx in row_indices:
        try:
            rows = driver.find_elements(By.CSS_SELECTOR, "tr")
            if idx >= len(rows):
                break
            row = rows[idx]

            symbol_anchors = row.find_elements(By.CSS_SELECTOR, "a.symbol-word-break")
            if not symbol_anchors:
                continue
            stock_name = symbol_anchors[0].text.strip()

            plus_td_elements = row.find_elements(By.CSS_SELECTOR, "td.togglecpm.plus")
            expanded = False

            if plus_td_elements:
                button = None
                try:
                    button = plus_td_elements[0].find_element(By.CSS_SELECTOR, "button.tbtn[aria-label*='Pre Book']")
                except NoSuchElementException:
                    try:
                        button = plus_td_elements[0].find_element(By.TAG_NAME, "button")
                    except NoSuchElementException:
                        button = None

                if button:
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", button)
                    wait.until(EC.element_to_be_clickable(button))
                    try:
                        button.click()
                    except ElementClickInterceptedException:
                        driver.execute_script("arguments[0].click()", button)
                    expanded = True
                    wait.until(lambda d: len(d.find_elements(By.CSS_SELECTOR, "tr")) > len(rows))

            rows = driver.find_elements(By.CSS_SELECTOR, "tr")
            next_row = rows[idx + 1] if expanded and idx + 1 < len(rows) else None

            aggressive_buyer = aggressive_seller = market_buyer = market_seller = ""

            if next_row:
                tds = next_row.find_elements(By.CSS_SELECTOR, "td.text-center")
                for i, td in enumerate(tds):
                    txt = td.text.strip()
                    if txt == "At the Open (ATO)":
                        if i - 1 >= 0:
                            aggressive_buyer = tds[i - 1].text.replace(",", "").strip()
                        if i + 1 < len(tds):
                            aggressive_seller = tds[i + 1].text.replace(",", "").strip()
                    if txt == "Total":
                        if i - 1 >= 0:
                            market_buyer = tds[i - 1].text.replace(",", "").strip()
                        if i + 1 < len(tds):
                            market_seller = tds[i + 1].text.replace(",", "").strip()

            records.append({
                "stock name": stock_name,
                "aggressive buyer": aggressive_buyer,
                "aggressive seller": aggressive_seller,
                "market buyer": market_buyer,
                "market seller": market_seller,
                "date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            })

            if expanded and next_row:
                rows = driver.find_elements(By.CSS_SELECTOR, "tr")
                row = rows[idx]
                plus_td_elements = row.find_elements(By.CSS_SELECTOR, "td.togglecpm.plus")
                close_button = None
                if plus_td_elements:
                    try:
                        close_button = plus_td_elements[0].find_element(By.CSS_SELECTOR, "button.tbtn[aria-label*='Pre Book']")
                    except NoSuchElementException:
                        try:
                            close_button = plus_td_elements[0].find_element(By.TAG_NAME, "button")
                        except NoSuchElementException:
                            close_button = None
                if close_button:
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", close_button)
                    wait.until(EC.element_to_be_clickable(close_button))
                    try:
                        close_button.click()
                    except ElementClickInterceptedException:
                        driver.execute_script("arguments[0].click()", close_button)
                    wait.until(EC.staleness_of(next_row))

        except Exception as e:
            print(f"Error processing row {idx} in thread: {e}")
            continue

    driver.quit()
    return records


def chunkify(lst, n):
    # Split list lst into n chunks
    return [lst[i::n] for i in range(n)]


if __name__ == "__main__":
    options = Options()
    options.add_argument("--headless=new")  # For initial row count, headless to speed up
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64 x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/92.0.451.100 Safari/537.36"
    )

    driver_single = webdriver.Chrome(options=options)
    wait_single = WebDriverWait(driver_single, 10)
    driver_single.get(url)

    wait_single.until(EC.presence_of_element_located((By.ID, "sel-Pre-Open-Market")))
    select_element = driver_single.find_element(By.ID, 'sel-Pre-Open-Market')
    Select(select_element).select_by_visible_text('Securities in F&O')

    # Allow table to load fully
    time.sleep(3)

    row_indices = get_all_stock_row_indices(driver_single)
    driver_single.quit()

    num_workers = 2  # Adjust number of parallel browsers to system capability

    row_chunks = chunkify(row_indices, num_workers)

    all_records = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=num_workers) as executor:
        futures = [executor.submit(scrape_batch, chunk) for chunk in row_chunks]
        for future in concurrent.futures.as_completed(futures):
            result = future.result()
            all_records.extend(result)

    print(f"Total records scraped: {len(all_records)}")

    if all_records:
        os.makedirs(os.path.dirname(file_path), exist_ok=True)
        if os.path.isfile(file_path):
            df_existing = pd.read_excel(file_path)
            df_new = pd.DataFrame(all_records)
            df_combined = pd.concat([df_existing, df_new], ignore_index=True)
            df_combined.to_excel(file_path, index=False)
        else:
            pd.DataFrame(all_records).to_excel(file_path, index=False)

    print("Parallel scraping complete and data saved.")
