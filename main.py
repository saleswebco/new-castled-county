import os
import time
import json
import pandas as pd
from datetime import datetime, timedelta
from google.oauth2 import service_account
from googleapiclient.discovery import build
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


class WillScraper:
    BASE_URL = "https://www3.newcastlede.gov/will/search/"
    YEAR = "2025"
    MAX_RETRIES = 5
    SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID")

    def __init__(self, headless=False):
        options = webdriver.ChromeOptions()
        if headless:
            options.add_argument("--headless=new")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")

        self.driver = webdriver.Remote(
            command_executor=os.environ.get("SELENIUM_REMOTE_URL"),
            options=options
        )
        self.wait = WebDriverWait(self.driver, 20)
        self.results = []

    def safe_find(self, xpath, default=""):
        """Try to find element by xpath, else return default"""
        try:
            return self.driver.find_element(By.XPATH, xpath).text.strip()
        except:
            return default

    def get_last_scraped_date(self):
        """Get the last 'Date of Death' from the most recent month's sheet in Google Sheets."""
        service = build('sheets', 'v4', credentials=self.get_google_credentials())
        now = datetime.now()
        last_month = now.month - 1 if now.month > 1 else 12
        last_year = now.year if now.month > 1 else now.year - 1
        last_month_sheet = f"{last_year}-{last_month:02d}"   # <-- FIXED zero-padded

        try:
            # Column E is Date of Death
            result = service.spreadsheets().values().get(
                spreadsheetId=self.SPREADSHEET_ID,
                range=f"'{last_month_sheet}'!E:E"  # <-- Date of Death column
            ).execute()
            values = result.get('values', [])

            if values:
                # Skip header row if present
                date_values = [row[0] for row in values if row and row[0] != "Date of Death"]

                if date_values:
                    last_date_str = date_values[-1]  # last non-header date
                    return datetime.strptime(last_date_str, "%m/%d/%Y")
        except Exception as e:
            print(f"Error retrieving last date from sheet '{last_month_sheet}': {e}")
        return None



    def create_sheet_if_missing(self, service, sheet_name):
        """Create a sheet if it doesn't already exist."""
        existing_sheets = service.spreadsheets().get(spreadsheetId=self.SPREADSHEET_ID).execute()
        sheet_titles = [sheet['properties']['title'] for sheet in existing_sheets.get('sheets', [])]

        if sheet_name not in sheet_titles:
            requests = [{"addSheet": {"properties": {"title": sheet_name}}}]
            body = {"requests": requests}
            service.spreadsheets().batchUpdate(spreadsheetId=self.SPREADSHEET_ID, body=body).execute()
            print(f"✓ Created sheet: {sheet_name}")

    def search_month(self, month: int, start_date: datetime = None):
        """Perform search for a given month starting from a specific date."""
        self.driver.get(self.BASE_URL)
        self.wait.until(EC.presence_of_element_located(
            (By.ID, "ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1__TextBoxYear")
        ))

        # Fill year & month
        self.driver.find_element(By.ID, "ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1__TextBoxYear").clear()
        self.driver.find_element(By.ID, "ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1__TextBoxYear").send_keys(self.YEAR)

        self.driver.find_element(By.ID, "ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1__TextBoxMonth").clear()
        self.driver.find_element(By.ID, "ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1__TextBoxMonth").send_keys(str(month))

        # If we have a start_date in this month → fill day
        if start_date and start_date.year == int(self.YEAR) and start_date.month == month:
            self.driver.find_element(By.ID, "ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1__TextBoxDay").clear()
            self.driver.find_element(By.ID, "ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1__TextBoxDay").send_keys(str(start_date.day))

        # Click Search
        self.driver.find_element(By.ID, "ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1__ButtonSearch").click()

    def process_results(self, month: int, start_date: str):
        """Process all rows in search results for a given month."""
        rows = self.driver.find_elements(By.XPATH, "//table[contains(@class,'grid')]/tbody/tr")

        for row_index in range(1, len(rows)):
            retries = 0
            success = False

            while retries < self.MAX_RETRIES and not success:
                try:
                    rows = self.driver.find_elements(By.XPATH, "//table[contains(@class,'grid')]/tbody/tr")
                    row = rows[row_index]
                    cols = row.find_elements(By.TAG_NAME, "td")
                    if not cols:
                        break

                    death_date = cols[4].text.strip()
                    death_date_obj = datetime.strptime(death_date, "%m/%d/%Y")  # Adjust date format as necessary

                    # Skip dates before the start date
                    if start_date and death_date_obj < start_date:
                        break

                    last_name = cols[1].text.strip()
                    first_name = cols[2].text.strip()

                    details_link = cols[0].find_element(By.TAG_NAME, "a")
                    self.driver.execute_script("arguments[0].click();", details_link)

                    self.wait.until(EC.presence_of_element_located(
                        (By.XPATH, "//h2[text()='Personal Representatives']")
                    ))

                    pr_table = self.driver.find_elements(
                        By.XPATH, "//h2[text()='Personal Representatives']/following-sibling::table[1]/tbody/tr"
                    )

                    if len(pr_table) > 1:  # skip header row
                        for rep_row in pr_table[1:]:
                            rep_cols = rep_row.find_elements(By.TAG_NAME, "td")
                            pr_name = rep_cols[0].text.strip()
                            pr_address = " ".join([c.text.strip() for c in rep_cols[1:]])

                            estate_date = self.safe_find("//label[contains(text(),'Date Estate Opened')]/../following-sibling::td")
                            decedent_address = self.safe_find("//label[contains(text(),'Decedent Address')]/../following-sibling::td")

                            self.results.append({
                                "Year": self.YEAR,
                                "Month": month,
                                "Last Name": last_name,
                                "First Name": first_name,
                                "Date of Death": death_date,
                                "Personal Representative Name": pr_name,
                                "Personal Representative Address": pr_address,
                                "Date Estate Opened": estate_date,
                                "Decedent Address": decedent_address,
                            })

                    self.driver.back()
                    self.wait.until(EC.presence_of_element_located(
                        (By.XPATH, "//table[contains(@class,'grid')]/tbody/tr")
                    ))

                    success = True

                except Exception as e:
                    retries += 1
                    print(f"[Retry {retries}/{self.MAX_RETRIES}] Error on row {row_index} (month {month}): {e}")
                    time.sleep(2)
                    if retries == self.MAX_RETRIES:
                        print(f"❌ Skipping row {row_index} in {self.YEAR}-{month} after {self.MAX_RETRIES} retries")

    def save_to_google_sheets(self):
        """Append results to Google Sheets."""
        service = build('sheets', 'v4', credentials=self.get_google_credentials())
        for result in self.results:
            self.append_row(service, result)

    def append_row(self, service, result):
        """Append a single row to Google Sheets."""
        sheet_name = f"{self.YEAR}-{result['Month']:02d}"
        self.create_sheet_if_missing(service, sheet_name)  # Create sheet if missing

        values = [[
            result['Year'],
            result['Month'],
            result['Last Name'],
            result['First Name'],
            result['Date of Death'],
            result['Personal Representative Name'],
            result['Personal Representative Address'],
            result['Date Estate Opened'],
            result['Decedent Address'],
        ]]
        body = {'values': values}
        service.spreadsheets().values().append(
            spreadsheetId=self.SPREADSHEET_ID,
            range=f"'{sheet_name}'!A1",
            valueInputOption="RAW",
            body=body
        ).execute()

    def get_google_credentials(self):
        """Load Google credentials from environment variable."""
        creds_raw = os.environ.get("GOOGLE_CREDENTIALS")
        return service_account.Credentials.from_service_account_info(json.loads(creds_raw))

    def run(self):
        last_date = self.get_last_scraped_date()
        start_date = last_date + timedelta(days=1) if last_date else None
        current_month = datetime.now().month
        current_year = datetime.now().year

        if int(self.YEAR) < current_year:
            self.YEAR = str(current_year)

        for month in range(1, current_month + 1):
            print(f"🔎 Searching {self.YEAR}-{month:02d}")
            self.search_month(month, start_date)
            self.save_to_google_sheets()
            self.results = []

        self.driver.quit()


if __name__ == "__main__":
    scraper = WillScraper(headless=True)
    scraper.run()