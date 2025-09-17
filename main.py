# import os
# import time
# import json
# from datetime import datetime, timedelta

# from google.oauth2 import service_account
# from googleapiclient.discovery import build
# from selenium import webdriver
# from selenium.webdriver.common.by import By
# from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.support import expected_conditions as EC


# class WillScraper:
#     BASE_URL = "https://www3.newcastlede.gov/will/search/"
#     YEAR = "2025"
#     MAX_RETRIES = 5
#     SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID")

#     def __init__(self, headless=False):
#         print("🚀 Initializing WebDriver...")
#         options = webdriver.ChromeOptions()
#         if headless:
#             options.add_argument("--headless=new")
#         options.add_argument("--no-sandbox")
#         options.add_argument("--disable-dev-shm-usage")

#         self.driver = webdriver.Remote(
#             command_executor=os.environ.get("SELENIUM_REMOTE_URL"),
#             options=options
#         )
#         self.wait = WebDriverWait(self.driver, 20)
#         self.results = []

#     def safe_find(self, xpath, default=""):
#         """Try to find element by xpath, else return default"""
#         try:
#             return self.driver.find_element(By.XPATH, xpath).text.strip()
#         except:
#             return default

#     def get_last_scraped_date(self):
#         """Get the last 'Date of Death' from the most recent month's sheet in Google Sheets."""
#         print("📅 Checking last scraped date from Google Sheets...")
#         service = build('sheets', 'v4', credentials=self.get_google_credentials())
#         now = datetime.now()
#         last_month = now.month - 1 if now.month > 1 else 12
#         last_year = now.year if now.month > 1 else now.year - 1
#         last_month_sheet = f"{last_year}-{last_month:02d}"

#         try:
#             result = service.spreadsheets().values().get(
#                 spreadsheetId=self.SPREADSHEET_ID,
#                 range=f"'{last_month_sheet}'!E:E"
#             ).execute()
#             values = result.get('values', [])
#             if values:
#                 date_values = [row[0] for row in values if row and row[0] != "Date of Death"]
#                 if date_values:
#                     last_date_str = date_values[-1]
#                     last_date = datetime.strptime(last_date_str, "%m/%d/%Y")
#                     print(f"✅ Last scraped date: {last_date_str}")
#                     return last_date
#         except Exception as e:
#             print(f"⚠️ Could not retrieve last date from sheet '{last_month_sheet}': {e}")
#         return None

#     def create_sheet_if_missing(self, service, sheet_name):
#         """Create a sheet if it doesn't already exist."""
#         print(f"📑 Ensuring sheet exists: {sheet_name}")
#         existing_sheets = service.spreadsheets().get(spreadsheetId=self.SPREADSHEET_ID).execute()
#         sheet_titles = [sheet['properties']['title'] for sheet in existing_sheets.get('sheets', [])]

#         if sheet_name not in sheet_titles:
#             requests = [{"addSheet": {"properties": {"title": sheet_name}}}]
#             body = {"requests": requests}
#             service.spreadsheets().batchUpdate(spreadsheetId=self.SPREADSHEET_ID, body=body).execute()
#             print(f"✅ Created sheet: {sheet_name}")
#         else:
#             print(f"ℹ️ Sheet already exists: {sheet_name}")

#     def search_month(self, month: int, start_date: datetime = None):
#         """Perform search for a given month starting from a specific date."""
#         print(f"🔍 Searching wills for {self.YEAR}-{month:02d}...")
#         self.driver.get(self.BASE_URL)
#         self.wait.until(EC.presence_of_element_located(
#             (By.ID, "ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1__TextBoxYear")
#         ))

#         # Set year
#         year_box = self.driver.find_element(By.ID, "ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1__TextBoxYear")
#         year_box.clear()
#         year_box.send_keys(self.YEAR)

#         # Set month
#         month_box = self.driver.find_element(By.ID, "ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1__TextBoxMonth")
#         month_box.clear()
#         month_box.send_keys(str(month))

#         # Optional: start day if needed
#         if start_date and start_date.year == int(self.YEAR) and start_date.month == month:
#             print(f"⏩ Starting from day {start_date.day}")
#             day_box = self.driver.find_element(By.ID, "ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1__TextBoxDay")
#             day_box.clear()
#             day_box.send_keys(str(start_date.day))

#         # Click Search
#         self.driver.find_element(By.ID, "ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1__ButtonSearch").click()
#         time.sleep(2)

#     def process_results(self, month: int, start_date: str):
#         """Process all rows in search results for a given month."""
#         print("📊 Processing search results...")
#         rows = self.driver.find_elements(By.XPATH, "//table[contains(@class,'grid')]/tbody/tr")

#         for row_index in range(1, len(rows)):
#             retries = 0
#             success = False

#             while retries < self.MAX_RETRIES and not success:
#                 try:
#                     rows = self.driver.find_elements(By.XPATH, "//table[contains(@class,'grid')]/tbody/tr")
#                     row = rows[row_index]
#                     cols = row.find_elements(By.TAG_NAME, "td")
#                     if not cols:
#                         break

#                     death_date = cols[4].text.strip()
#                     death_date_obj = datetime.strptime(death_date, "%m/%d/%Y")

#                     if start_date and death_date_obj < start_date:
#                         print(f"⏭️ Skipping row {row_index} before start date ({death_date})")
#                         break

#                     last_name = cols[1].text.strip()
#                     first_name = cols[2].text.strip()
#                     print(f"➡️ Found record: {last_name}, {first_name}, DoD {death_date}")

#                     details_link = cols[0].find_element(By.TAG_NAME, "a")
#                     self.driver.execute_script("arguments[0].click();", details_link)
#                     self.wait.until(EC.presence_of_element_located(
#                         (By.XPATH, "//h2[text()='Personal Representatives']")
#                     ))

#                     pr_table = self.driver.find_elements(
#                         By.XPATH, "//h2[text()='Personal Representatives']/following-sibling::table[1]/tbody/tr"
#                     )

#                     if len(pr_table) > 1:
#                         for rep_row in pr_table[1:]:
#                             rep_cols = rep_row.find_elements(By.TAG_NAME, "td")
#                             pr_name = rep_cols[0].text.strip()
#                             pr_address = " ".join([c.text.strip() for c in rep_cols[1:]])

#                             estate_date = self.safe_find("//label[contains(text(),'Date Estate Opened')]/../following-sibling::td")
#                             decedent_address = self.safe_find("//label[contains(text(),'Decedent Address')]/../following-sibling::td")

#                             print(f"   👤 PR: {pr_name}, Address: {pr_address}")

#                             self.results.append({
#                                 "Year": self.YEAR,
#                                 "Month": month,
#                                 "Last Name": last_name,
#                                 "First Name": first_name,
#                                 "Date of Death": death_date,
#                                 "Personal Representative Name": pr_name,
#                                 "Personal Representative Address": pr_address,
#                                 "Date Estate Opened": estate_date,
#                                 "Decedent Address": decedent_address,
#                             })

#                     self.driver.back()
#                     self.wait.until(EC.presence_of_element_located(
#                         (By.XPATH, "//table[contains(@class,'grid')]/tbody/tr")
#                     ))
#                     time.sleep(1)

#                     success = True

#                 except Exception as e:
#                     retries += 1
#                     print(f"[Retry {retries}/{self.MAX_RETRIES}] Error on row {row_index} (month {month}): {e}")
#                     time.sleep(2)
#                     if retries == self.MAX_RETRIES:
#                         print(f"❌ Skipping row {row_index} in {self.YEAR}-{month}")

#     def save_to_google_sheets(self):
#         """Append results to Google Sheets."""
#         print("💾 Saving results to Google Sheets...")
#         service = build('sheets', 'v4', credentials=self.get_google_credentials())
#         for result in self.results:
#             self.append_row(service, result)
#         print(f"✅ Saved {len(self.results)} records to Google Sheets")

#     def append_row(self, service, result):
#         """Append a single row to Google Sheets."""
#         sheet_name = f"{self.YEAR}-{result['Month']:02d}"
#         self.create_sheet_if_missing(service, sheet_name)

#         values = [[
#             result['Year'],
#             result['Month'],
#             result['Last Name'],
#             result['First Name'],
#             result['Date of Death'],
#             result['Personal Representative Name'],
#             result['Personal Representative Address'],
#             result['Date Estate Opened'],
#             result['Decedent Address'],
#         ]]
#         body = {'values': values}

#         service.spreadsheets().values().append(
#             spreadsheetId=self.SPREADSHEET_ID,
#             range=f"'{sheet_name}'!A1",
#             valueInputOption="RAW",
#             body=body
#         ).execute()

#     def get_google_credentials(self):
#         """Load Google credentials from environment variable."""
#         creds_raw = os.environ.get("GOOGLE_CREDENTIALS")
#         return service_account.Credentials.from_service_account_info(json.loads(creds_raw))

#     def run(self):
#         print("▶️ Starting Will Scraper...")
#         last_date = self.get_last_scraped_date()
#         start_date = last_date + timedelta(days=1) if last_date else None
#         current_month = datetime.now().month
#         current_year = datetime.now().year

#         if int(self.YEAR) < current_year:
#             self.YEAR = str(current_year)

#         for month in range(1, current_month + 1):
#             print(f"🔎 Running search for {self.YEAR}-{month:02d}")
#             self.search_month(month, start_date)
#             self.process_results(month, start_date)
#             self.save_to_google_sheets()
#             self.results = []
#             time.sleep(2)

#         self.driver.quit()
#         print("🏁 Finished scraping!")


# if __name__ == "__main__":
#     scraper = WillScraper(headless=True)
#     scraper.run()




## ------------------- the above code is get last date and then rerun scraping the whole so if we want it it is comment -------










import os
import time
import json
from datetime import datetime, timedelta
from google.oauth2 import service_account
from googleapiclient.discovery import build
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


class WillScraper:
    BASE_URL = "https://www3.newcastlede.gov/will/search/"
    MAX_RETRIES = 5
    SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID")

    def __init__(self, headless=False):
        print("🚀 Initializing WebDriver...")
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
        try:
            return self.driver.find_element(By.XPATH, xpath).text.strip()
        except:
            return default

    def create_sheet_if_missing(self, service, sheet_name):
        """Ensure the sheet exists and has correct header"""
        existing_sheets = service.spreadsheets().get(
            spreadsheetId=self.SPREADSHEET_ID
        ).execute()
        sheet_titles = [s["properties"]["title"] for s in existing_sheets.get("sheets", [])]

        if sheet_name not in sheet_titles:
            requests = [{"addSheet": {"properties": {"title": sheet_name}}}]
            body = {"requests": requests}
            service.spreadsheets().batchUpdate(
                spreadsheetId=self.SPREADSHEET_ID, body=body
            ).execute()
            print(f"✅ Created sheet: {sheet_name}")

            # Add header row
            header = [[
                "Year", "Month", "Last Name", "First Name", "Date of Death",
                "Personal Representative Name", "Personal Representative Address",
                "Date Estate Opened", "Decedent Address"
            ]]
            service.spreadsheets().values().append(
                spreadsheetId=self.SPREADSHEET_ID,
                range=f"'{sheet_name}'!A1",
                valueInputOption="RAW",
                body={"values": header},
            ).execute()

    def sheet_records_set(self, service, sheet_name):
        """Return set of (Last Name, First Name, Date of Death) already in the sheet"""
        try:
            result = service.spreadsheets().values().get(
                spreadsheetId=self.SPREADSHEET_ID,
                range=f"'{sheet_name}'!C:E"  # Last Name, First Name, Date of Death
            ).execute()
            values = result.get("values", [])
            records = set()
            for row in values[1:]:  # skip header
                if len(row) >= 3:
                    records.add((row[0].strip(), row[1].strip(), row[2].strip()))
            return records
        except:
            return set()

    def search_month(self, year: int, month: int):
        """Search by month/year"""
        print(f"🔍 Searching wills for {year}-{month:02d}...")
        self.driver.get(self.BASE_URL)
        self.wait.until(
            EC.presence_of_element_located(
                (By.ID, "ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1__TextBoxYear")
            )
        )

        year_box = self.driver.find_element(By.ID, "ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1__TextBoxYear")
        month_box = self.driver.find_element(By.ID, "ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1__TextBoxMonth")

        year_box.clear()
        year_box.send_keys(str(year))
        month_box.clear()
        month_box.send_keys(str(month))

        self.driver.find_element(By.ID, "ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1__ButtonSearch").click()
        time.sleep(2)

    def process_results(self, year: int, month: int):
        """Process all rows for a given month"""
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

                    last_name = cols[1].text.strip()
                    first_name = cols[2].text.strip()
                    death_date = cols[4].text.strip()

                    details_link = cols[0].find_element(By.TAG_NAME, "a")
                    self.driver.execute_script("arguments[0].click();", details_link)
                    self.wait.until(
                        EC.presence_of_element_located(
                            (By.XPATH, "//h2[text()='Personal Representatives'] | //h2[text()='Decedent Information']")
                        )
                    )

                    # Estate date logic
                    estate_admin = self.safe_find("//label[contains(text(),'Date Estate Opened (Administration)')]/../following-sibling::td")
                    estate_test = self.safe_find("//label[contains(text(),'Date Estate Opened (Testamentary)')]/../following-sibling::td")
                    estate_date = estate_admin if estate_admin else (estate_test if estate_test else "")

                    decedent_address = self.safe_find("//label[contains(text(),'Decedent Address')]/../following-sibling::td")

                    # PR table
                    pr_table = self.driver.find_elements(
                        By.XPATH, "//h2[text()='Personal Representatives']/following-sibling::table[1]/tbody/tr"
                    )
                    if len(pr_table) > 1:
                        for rep_row in pr_table[1:]:
                            rep_cols = rep_row.find_elements(By.TAG_NAME, "td")
                            pr_name = rep_cols[0].text.strip()
                            pr_address = " ".join([c.text.strip() for c in rep_cols[1:]])

                            self.results.append({
                                "Year": year,
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
                    self.wait.until(
                        EC.presence_of_element_located(
                            (By.XPATH, "//table[contains(@class,'grid')]/tbody/tr")
                        )
                    )
                    time.sleep(1)
                    success = True
                except Exception as e:
                    retries += 1
                    print(f"[Retry {retries}/{self.MAX_RETRIES}] Error on row {row_index} ({year}-{month}): {e}")
                    time.sleep(2)
                    if retries == self.MAX_RETRIES:
                        print(f"❌ Skipping row {row_index} in {year}-{month}")

    def save_to_google_sheets(self, year: int, month: int):
        print("💾 Saving results to Google Sheets...")
        service = build("sheets", "v4", credentials=self.get_google_credentials())
        sheet_name = f"{year}_{datetime(year, month, 1).strftime('%b')}"
        self.create_sheet_if_missing(service, sheet_name)

        existing_records = self.sheet_records_set(service, sheet_name)
        new_rows = []

        for result in self.results:
            key = (result["Last Name"], result["First Name"], result["Date of Death"])
            if key not in existing_records:
                new_rows.append([
                    result["Year"],
                    result["Month"],
                    result["Last Name"],
                    result["First Name"],
                    result["Date of Death"],
                    result["Personal Representative Name"],
                    result["Personal Representative Address"],
                    result["Date Estate Opened"],
                    result["Decedent Address"],
                ])

        if new_rows:
            body = {"values": new_rows}
            service.spreadsheets().values().append(
                spreadsheetId=self.SPREADSHEET_ID,
                range=f"'{sheet_name}'!A1",
                valueInputOption="RAW",
                body=body,
            ).execute()
            print(f"✅ Added {len(new_rows)} new records to {sheet_name}")
        else:
            print("ℹ️ No new records to add.")

    def get_google_credentials(self):
        creds_raw = os.environ.get("GOOGLE_CREDENTIALS")
        return service_account.Credentials.from_service_account_info(json.loads(creds_raw))

    def run(self):
        print("▶️ Starting Will Scraper...")
        today = datetime.now()
        current_year, current_month = today.year, today.month
    
        # Always scrape previous month + current month on first run
        service = build("sheets", "v4", credentials=self.get_google_credentials())
        spreadsheet = service.spreadsheets().get(spreadsheetId=self.SPREADSHEET_ID).execute()
        sheet_titles = [s["properties"]["title"] for s in spreadsheet.get("sheets", [])]
    
        months_to_scrape = []
    
        # If current month sheet not found → scrape last month + current month
        current_sheet = f"{current_year}_{today.strftime('%b')}"
        if current_sheet not in sheet_titles:
            last_month_date = today.replace(day=1) - timedelta(days=1)
            last_sheet = f"{last_month_date.year}_{last_month_date.strftime('%b')}"
            months_to_scrape.append((last_month_date.year, last_month_date.month))
            months_to_scrape.append((current_year, current_month))
        else:
            # Already scraped before → only update current month
            months_to_scrape.append((current_year, current_month))
    
        print(f"📆 Months to scrape: {months_to_scrape}")
    
        for year, month in months_to_scrape:
            self.search_month(year, month)
            self.process_results(year, month)
            self.save_to_google_sheets()
            self.results = []
    
        # Update summary after scraping all data
        self.update_summary(service)
        self.driver.quit()
        print("🏁 Finished scraping!")

if __name__ == "__main__":
    scraper = WillScraper(headless=True)
    scraper.run()


