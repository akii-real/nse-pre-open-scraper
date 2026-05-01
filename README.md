# NSE Pre Open Scraper

A Python + Selenium based scraper for collecting **NSE pre-open market depth signals** for **F&O securities** and saving them into an Excel file for historical analysis.

This project opens the NSE pre-open market page, filters the view to `Securities in F&O`, expands each stock row to read pre-book details, and captures buy/sell-side quantities such as:

- aggressive buyer
- aggressive seller
- market buyer
- market seller

Each run appends timestamped records to an Excel file so you can build a running dataset over multiple days/sessions.

## What This Project Does

- Navigates to NSE pre-open market page.
- Selects `Securities in F&O` from the dropdown.
- Identifies valid stock rows containing a symbol.
- Expands each row to read the detail/pre-book row.
- Extracts:
  - `stock name`
  - `aggressive buyer` (ATO side)
  - `aggressive seller` (ATO side)
  - `market buyer` (Total side)
  - `market seller` (Total side)
  - `date` (current timestamp)
- Runs scraping in parallel using multiple browser instances (`ThreadPoolExecutor`).
- Appends the output to an Excel file if it already exists, otherwise creates a new one.

## Why It Is Useful

Pre-open order book behavior can provide an early indication of demand/supply pressure before regular market trading starts. By storing these values continuously in a structured file, this scraper helps with:

- intraday/pre-open behavior tracking
- signal backtesting with historical snapshots
- manual strategy validation
- rapid screening of F&O names in the opening phase

## Tech Stack

- Python 3
- Selenium WebDriver (Chrome)
- Pandas
- OpenPyXL (used by Pandas for Excel read/write)
- concurrent.futures (built-in Python threading utilities)

## Project Structure

- `NSE Scraper.py` - main scraper script
- `Book1_parallel.xlsx` - sample/output Excel file

## How It Works (Flow)

1. Launches one headless Chrome session to get all target row indices.
2. Splits row indices into chunks (`chunkify`) based on number of workers.
3. Launches parallel workers; each worker:
   - opens NSE page
   - applies F&O filter
   - loops through assigned row indices
   - expands row, extracts values, collapses row
4. Merges all worker records.
5. Writes data to Excel (append mode if file exists).

## Setup

1. Install Python 3.9+.
2. Install Google Chrome.
3. Install dependencies:

```bash
pip install selenium pandas openpyxl
```

4. Ensure matching ChromeDriver is available for your Chrome version.
   - If Selenium Manager is enabled (newer Selenium), driver resolution may be automatic.
   - Otherwise, install ChromeDriver and make sure it is in your PATH.

## Configuration

Update these values in `NSE Scraper.py` as needed:

- `file_path`: destination Excel file path
- `num_workers`: number of parallel browser instances
- `url`: source page (currently NSE pre-open market page)

## Usage

Run the script:

```bash
python "NSE Scraper.py"
```

On successful execution, console output includes:

- total records scraped
- confirmation that parallel scraping completed and data was saved

## Output Format

The Excel file contains rows with these columns:

- `stock name`
- `aggressive buyer`
- `aggressive seller`
- `market buyer`
- `market seller`
- `date`

## Notes and Limitations

- NSE may change page layout/HTML classes; selectors may need updates.
- Heavy parallelism can increase resource usage and may trigger site-side protections.
- Dynamic pages can occasionally cause stale/intercepted click issues; the script already includes fallback handling for common Selenium exceptions.
- Respect NSE terms of use and rate limits when running automated scripts.

## Suggested Improvements

- Add logging instead of print statements.
- Add retry/backoff handling around page load and row expansion.
- Export CSV and database options in addition to Excel.
- Add scheduling (Windows Task Scheduler / cron) for automated periodic collection.
- Add deduplication rules for repeated snapshots.

## Disclaimer

This tool is for educational/research purposes. Always verify extracted market data before taking trading decisions.
