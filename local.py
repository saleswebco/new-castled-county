import os
import time
import pandas as pd
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options


class WillScraper:
    BASE_URL = "https://www3.newcastlede.gov/will/search/"
    YEAR = "2025"
    OUTPUT_XLSX = "wills_2025.xlsx"
    MAX_RETRIES = 5

    def __init__(self, headless=False):
        options = Options()
        if headless:
            options.add_argument("--headless=new")
        options.add_argument("--start-maximized")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")

        self.driver = webdriver.Chrome(options=options)
        self.wait = WebDriverWait(self.driver, 20)
        self.results = []

        # Load previous data if exists
        if os.path.exists(self.OUTPUT_XLSX):
            self.df_existing = pd.read_excel(self.OUTPUT_XLSX)
        else:
            self.df_existing = pd.DataFrame()

    def safe_find(self, xpath, default=""):
        """Try to find element by xpath, else return default"""
        try:
            return self.driver.find_element(By.XPATH, xpath).text.strip()
        except:
            return default

    def search_month(self, month: int):
        """Perform search for a given month"""
        self.driver.get(self.BASE_URL)
        self.wait.until(EC.presence_of_element_located(
            (By.ID, "ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1__TextBoxYear")
        ))

        # Fill year & month
        year_input = self.driver.find_element(By.ID, "ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1__TextBoxYear")
        month_input = self.driver.find_element(By.ID, "ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1__TextBoxMonth")

        year_input.clear()
        year_input.send_keys(self.YEAR)
        month_input.clear()
        month_input.send_keys(str(month))

        # Click Search
        search_btn = self.driver.find_element(By.ID, "ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1__ButtonSearch")
        search_btn.click()

        # Wait for results table (if exists)
        try:
            self.wait.until(EC.presence_of_element_located(
                (By.XPATH, "//table[contains(@class,'grid')]/tbody/tr")
            ))
        except:
            print(f"No results for {self.YEAR}-{month}")
            return

        self.process_results(month)

    def process_results(self, date: datetime):
        """Process results for a given day."""
        print("📊 Processing search results...")
        rows = self.driver.find_elements(By.XPATH, "//table[contains(@class,'grid')]/tbody/tr")

        for row_index in range(1, len(rows)):
            retries, success = 0, False

            while retries < self.MAX_RETRIES and not success:
                try:
                    rows = self.driver.find_elements(By.XPATH, "//table[contains(@class,'grid')]/tbody/tr")
                    row = rows[row_index]
                    cols = row.find_elements(By.TAG_NAME, "td")
                    if not cols:
                        break

                    death_date = cols[4].text.strip()
                    if death_date != date.strftime("%m/%d/%Y"):
                        break

                    last_name = cols[1].text.strip()
                    first_name = cols[2].text.strip()
                    print(f"➡️ Found record: {last_name}, {first_name}, DoD {death_date}")

                    details_link = cols[0].find_element(By.TAG_NAME, "a")
                    self.driver.execute_script("arguments[0].click();", details_link)
                    self.wait.until(
                        EC.presence_of_element_located(
                            (By.XPATH, "//h2[text()='Personal Representatives'] | //h2[text()='Decedent Information']")
                        )
                    )

                    # Prefer Administration date, fallback to Testamentary
                    estate_date = self.safe_find(
                        "//label[contains(text(),'Date Estate Opened (Administration)')]/../following-sibling::td"
                    )
                    if not estate_date:
                        estate_date = self.safe_find(
                            "//label[contains(text(),'Date Estate Opened (Testamentary)')]/../following-sibling::td"
                        )

                    decedent_address = self.safe_find("//label[contains(text(),'Decedent Address')]/../following-sibling::td")

                    # Save record (always save decedent info)
                    self.results.append({
                        "Year": date.year,
                        "Month": date.month,
                        "Last Name": last_name,
                        "First Name": first_name,
                        "Date of Death": death_date,
                        "Personal Representative Name": "",
                        "Personal Representative Address": "",
                        "Date Estate Opened": estate_date,
                        "Decedent Address": decedent_address,
                    })

                    self.driver.back()
                    self.wait.until(
                        EC.presence_of_element_located(
                            (By.XPATH, "//table[contains(@class,'grid')]/tbody/tr")
                        )
                    )
                    time.sleep(1)
                    success = True
                except Exception as e:
                    retries += 1
                    print(f"[Retry {retries}/{self.MAX_RETRIES}] Error on row {row_index} ({date}): {e}")
                    time.sleep(2)
                    if retries == self.MAX_RETRIES:
                        print(f"❌ Skipping row {row_index} on {date}")

    def get_last_scraped_month(self):
        """Check the last month already scraped"""
        if self.df_existing.empty:
            return 0
        return int(self.df_existing["Month"].max())

    def save_results(self, month: int):
        """Save results for a specific month into its own sheet"""
        df_new = pd.DataFrame(self.results)
        if df_new.empty:
            print(f"⚠️ No new records scraped for {self.YEAR}-{month}")
            return

        sheet_name = f"{self.YEAR}-{month:02d}"

        # Write to month-specific sheet
        if os.path.exists(self.OUTPUT_XLSX):
            with pd.ExcelWriter(self.OUTPUT_XLSX, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                df_new.to_excel(writer, sheet_name=sheet_name, index=False)
        else:
            with pd.ExcelWriter(self.OUTPUT_XLSX, engine="openpyxl", mode="w") as writer:
                df_new.to_excel(writer, sheet_name=sheet_name, index=False)

        print(f"✅ Saved {len(df_new)} records to sheet {sheet_name} in {self.OUTPUT_XLSX}")


    def run(self):
        """Main scraper runner"""
        current_month = datetime.now().month

        # Always start from 1 to current month for full workbook
        for month in range(1, current_month + 1):
            print(f"🔎 Searching {self.YEAR}-{month}")
            self.search_month(month)
            self.save_results(month)
            self.results = []  # clear buffer after saving

        self.driver.quit()


if __name__ == "__main__":
    scraper = WillScraper(headless=False)  # set headless=True for background run
    scraper.run()
