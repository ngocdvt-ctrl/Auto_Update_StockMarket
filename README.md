# 📈 Update-StockMarket: VNIndex Scraper

This project is a Python program that uses Selenium to automatically scrape key metrics of the Vietnamese stock market (VNIndex) from the VNDirect price board website. The data is recorded into an Excel file, utilizing a special **trading day adjustment logic** to ensure the data is always assigned to the correct **Trading Date**, even when collected during weekends or before market opening hours.

## 📌 Key Features

* **Primary Index Scraping:** Collects VNIndex, Spread (Points and %), Total Value, Total Volume, and the count of stocks that are Up (`Meigara_Up`), Down (`Meigara_Down`), and Unchanged (`Meigara_Unchanged`).
* **Accurate Trading Date Logic:** Automatically adjusts the data collection date back to the last valid trading day if the script runs before 9:00 AM (opening hour) or on a Saturday/Sunday.
* **Spread Handling:** Identifies the Up/Down trend using the on-page icon and applies a negative sign (-) to the Spread metrics if the market is declining.
* **Duplicate Check:** Prevents recording identical entries when the market is closed or when there is no price movement.

## 🛠️ Installation Requirements

To run the program, you need to install:

1.  **Python 3.x**
2.  **Chrome Browser** and **ChromeDriver** (Selenium will automatically manage ChromeDriver if installed correctly).

### Python Libraries

Install the necessary libraries using the following command:

```bash
pip install selenium pandas openpyxl numpy
