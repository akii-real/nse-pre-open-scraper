from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from datetime import datetime
import os
import time


file_path = r"D:\Projects\NSC Pre Open Scraper\Book1.xlsx"
url = "https://www.nseindia.com/market-data/pre-open-market-cm-and-emerge-market"

options = Options()
options.add_argument("--headless=new")   # Run Chrome in headless mode (new mode)
options.add_argument("--disable-gpu")    # Helps with some headless issues on Windows
options.add_argument("--window-size=1920,1080")  # Set window size to standard desktop resolution
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_argument(
    "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64 x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/92.0.451.100 Safari/537.36"
)

driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 20)

driver.get(url)

# Wait for and select category dropdown
wait.until(EC.presence_of_element_located((By.ID, "sel-Pre-Open-Market")))
select_element = driver.find_element(By.ID, 'sel-Pre-Open-Market')
select_category = Select(select_element)
select_category.select_by_visible_text('Securities in F&O')
print("Category set to 'Securities in F&O'")

# Wait for updated table to load stock symbols
wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "a.symbol-word-break")))

records = []
idx = 0

while True:
    rows = driver.find_elements(By.CSS_SELECTOR, "tr")
    print(f"Iteration idx={idx}, total rows={len(rows)}")
    if idx >= len(rows):
        break

    row = rows[idx]
    print(f"Row HTML snippet: {row.get_attribute('innerHTML')[:200]}")  # first 200 chars

    try:
        symbol_anchors = row.find_elements(By.CSS_SELECTOR, "a.symbol-word-break")
        if not symbol_anchors:
            print(f"Row idx={idx} skipped: no stock symbol found.")
            idx += 1
            continue

        stock_name = symbol_anchors[0].text.strip()
        print(f"Processing: {stock_name}")

        plus_td_elements = row.find_elements(By.CSS_SELECTOR, "td.togglecpm.plus")
        expanded = False
        if plus_td_elements:
            try:
                button = plus_td_elements[0].find_element(By.CSS_SELECTOR, "button.tbtn[aria-label*='Pre Book']")
            except:
                button = plus_td_elements[0].find_element(By.TAG_NAME, "button")

            print(f"Found plus button for {stock_name}, clicking...")
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", button)
            # Use explicit wait until button is clickable
            wait.until(EC.element_to_be_clickable(button))
            try:
                button.click()
            except Exception as e:
                print("Click intercepted, trying JS click...", e)
                driver.execute_script("arguments[0].click()")
            expanded = True

            # Wait for expanded row to appear after clicking expand
            wait.until(
                EC.presence_of_element_located(
                    (By.XPATH, f"//tr[following-sibling::tr[{idx + 1}]]")
                )
            )
            # Alternatively, simply wait for next row (expanded row) visibility
            wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "tr:nth-child({})".format(idx + 2))))

        # Refresh rows
        rows = driver.find_elements(By.CSS_SELECTOR, "tr")
        next_row = rows[idx + 1] if expanded and idx + 1 < len(rows) else None

        aggressive_buyer = ""
        aggressive_seller = ""
        market_buyer = ""
        market_seller = ""

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

        print(f"Extracted data for {stock_name} - Aggressive Buyer: {aggressive_buyer}, Aggressive Seller: {aggressive_seller}, Market Buyer: {market_buyer}, Market Seller: {market_seller}")

        records.append(
            {
                "stock name": stock_name,
                "aggressive buyer": aggressive_buyer,
                "aggressive seller": aggressive_seller,
                "market buyer": market_buyer,
                "market seller": market_seller,
                "date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }
        )
        
        # Collapse the expanded row by clicking the plus button again
        if expanded:
            try:
                # Refresh rows and re-fetch current row to avoid stale element reference
                rows = driver.find_elements(By.CSS_SELECTOR, "tr")
                row = rows[idx]  # re-fetch current row

                plus_td_elements = row.find_elements(By.CSS_SELECTOR, "td.togglecpm.plus")
                if plus_td_elements:
                    try:
                        close_button = plus_td_elements[0].find_element(By.CSS_SELECTOR, "button.tbtn[aria-label*='Pre Book']")
                    except:
                        close_button = plus_td_elements[0].find_element(By.TAG_NAME, "button")

                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", close_button)
                    wait.until(EC.element_to_be_clickable(close_button))
                    try:
                        close_button.click()
                        print(f"Collapsed expanded row for {stock_name}")
                    except Exception as e:
                        print("Click intercepted while closing, trying JS click...", e)
                        driver.execute_script("arguments[0].click()")
                        print(f"Collapsed expanded row for {stock_name} (JS click fallback)")
                    # Wait for collapse effect (next row disappears)
                    wait.until(EC.invisibility_of_element(next_row))
            except Exception as e:
                print(f"Error collapsing expanded row for {stock_name}: {e}")

        idx += 1

    except Exception as e:
        print(f"Error processing row {idx}: {e}")
        idx += 1

driver.quit()

print(f"Total records scraped: {len(records)}")
for rec in records:
    print(rec)

if records:
    os.makedirs(os.path.dirname(file_path), exist_ok=True)
    if os.path.isfile(file_path):
        df_exist = pd.read_excel(file_path)
        df = pd.DataFrame(records)
        df_combined = pd.concat([df_exist, df], ignore_index=True)
        df_combined.to_excel(file_path, index=False)
    else:
        pd.DataFrame(records).to_excel(file_path, index=False)

print("Scraping complete and data saved.")
