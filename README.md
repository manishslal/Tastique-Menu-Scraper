# Toastique Menu Scraper

This Python script scrapes menu item names and prices from Toastique online ordering pages using Selenium and BeautifulSoup. It reads a list of restaurant locations and URLs from `Menu Links.txt` and saves the combined menu data into an Excel file (`.xlsx`), with each state's menus on a separate sheet.

## Setup

1.  **Install Python:** Ensure you have Python 3 installed.
2.  **Install Libraries:** You need the following libraries. You can install them using pip:
    ```bash
    pip install pandas openpyxl selenium beautifulsoup4 pytz webdriver-manager
    ```
    *(Note: Using `webdriver-manager` can simplify handling the ChromeDriver setup, though your current script uses `Service()` directly which is also fine).*
3.  **ChromeDriver:** This script uses Selenium, which requires a WebDriver compatible with your Chrome browser version. Ensure you have ChromeDriver installed and accessible in your system's PATH, or adjust the script to manage it (e.g., using `webdriver-manager`).
4.  **Input File:** Create a file named `Menu Links.txt` in the same directory as the script. The format should be:
    ```
    State Name
    Location Name 1 - [https://url-for-location-1.com](https://url-for-location-1.com)
    Location Name 2 - [https://url-for-location-2.com](https://url-for-location-2.com)

    Another State Name
    Location Name 3 - [https://url-for-location-3.com](https://url-for-location-3.com)
    ```

## Usage

1.  Make sure `Menu Links.txt` is populated with the correct URLs.
2.  Run the script from your terminal:
    ```bash
    python Menu_Scrape_Combined4.py
    ```
    *(Or `python menu_scraper.py` if you rename the file)*
3.  The script will create a folder named with today's date (MM-DD-YYYY) and save an Excel file inside it (e.g., `Toastique Menu Master File - 02'11PM.xlsx`).

## Notes

* The script uses headless Chrome for scraping.
* It includes basic error handling and retries failed links once.
* Excel sheet names are based on state abbreviations.