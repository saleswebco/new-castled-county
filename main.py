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

    def get_last_scraped_date(self):
        print("📅 Checking last scraped date from Google Sheets...")
        service = build('sheets', 'v4', credentials=self.get_google_credentials())
        # Check all sheets to find the maximum date of death entered
        try:
            spreadsheet = service.spreadsheets().get(spreadsheetId=self.SPREADSHEET_ID).execute()
            sheets = spreadsheet.get("sheets", [])
            last_date = None

            for sheet in sheets:
                title = sheet["properties"]["title"]
                try:
                    result = service.spreadsheets().values().get(
                        spreadsheetId=self.SPREADSHEET_ID,
                        range=f"'{title}'!E:E"  # Date of Death column
                    ).execute()
                    values = result.get('values', [])
                    if values:
                        date_values = [row[0] for row in values if row and row[0] != "Date of Death"]
                        for dv in date_values:
                            try:
                                d = datetime.strptime(dv, "%m/%d/%Y")
                                if not last_date or d > last_date:
                                    last_date = d
                            except:
                                pass
                except:
                    continue

            if last_date:
                print(f"✅ Last scraped date: {last_date.strftime('%m/%d/%Y')}")
            else:
                print("⚠️ No previous date found. Starting fresh.")
            return last_date
        except Exception as e:
            print(f"⚠️ Error reading Sheets: {e}")
            return None

    def create_sheet_if_missing(self, service, sheet_name):
        print(f"📑 Ensuring sheet exists: {sheet_name}")
        existing_sheets = service.spreadsheets().get(
            spreadsheetId=self.SPREADSHEET_ID
        ).execute()
        sheet_titles = [sheet["properties"]["title"] for sheet in existing_sheets.get("sheets", [])]

        if sheet_name not in sheet_titles:
            requests = [{"addSheet": {"properties": {"title": sheet_name}}}]
            body = {"requests": requests}
            service.spreadsheets().batchUpdate(
                spreadsheetId=self.SPREADSHEET_ID, body=body
            ).execute()
            print(f"✅ Created sheet: {sheet_name}")
        else:
            print(f"ℹ️ Sheet already exists: {sheet_name}")

    def search_day(self, date: datetime):
        """Search for wills on a specific date."""
        print(f"🔍 Searching wills for {date.strftime('%Y-%m-%d')}...")
        self.driver.get(self.BASE_URL)
        self.wait.until(
            EC.presence_of_element_located(
                (By.ID, "ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1__TextBoxYear")
            )
        )

        # Fill in year, month, day
        year_box = self.driver.find_element(By.ID, "ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1__TextBoxYear")
        year_box.clear()
        year_box.send_keys(str(date.year))

        month_box = self.driver.find_element(By.ID, "ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1__TextBoxMonth")
        month_box.clear()
        month_box.send_keys(str(date.month))

        day_box = self.driver.find_element(By.ID, "ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1__TextBoxDay")
        day_box.clear()
        day_box.send_keys(str(date.day))

        # Search
        self.driver.find_element(By.ID, "ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1__ButtonSearch").click()
        time.sleep(2)

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

    def save_to_google_sheets(self):
        print("💾 Saving results to Google Sheets...")
        service = build("sheets", "v4", credentials=self.get_google_credentials())
        for result in self.results:
            self.append_row(service, result)
        print(f"✅ Saved {len(self.results)} records to Google Sheets")

    def append_row(self, service, result):
        sheet_name = f"{result['Year']}-{result['Month']:02d}"
        self.create_sheet_if_missing(service, sheet_name)

        values = [[
            result["Year"],
            result["Month"],
            result["Last Name"],
            result["First Name"],
            result["Date of Death"],
            result["Personal Representative Name"],
            result["Personal Representative Address"],
            result["Date Estate Opened"],
            result["Decedent Address"],
        ]]
        body = {"values": values}
        service.spreadsheets().values().append(
            spreadsheetId=self.SPREADSHEET_ID,
            range=f"'{sheet_name}'!A1",
            valueInputOption="RAW",
            body=body,
        ).execute()

    def get_google_credentials(self):
        creds_raw = os.environ.get("GOOGLE_CREDENTIALS")
        return service_account.Credentials.from_service_account_info(json.loads(creds_raw))

    def run(self):
        print("▶️ Starting Will Scraper...")
        last_date = self.get_last_scraped_date()
        start_date = last_date + timedelta(days=1) if last_date else None
        today = datetime.now()

        if not start_date:
            print("⚠️ No last date found, starting from today.")
            start_date = today

        print(f"📆 Scraping range: {start_date.strftime('%Y-%m-%d')} ➡ {today.strftime('%Y-%m-%d')}")

        current_date = start_date
        while current_date <= today:
            self.search_day(current_date)
            self.process_results(current_date)
            self.save_to_google_sheets()
            self.results = []
            current_date += timedelta(days=1)

        # Update summary after scraping all data
        service = build("sheets", "v4", credentials=self.get_google_credentials())
        self.update_summary(service)
        self.driver.quit()
        print("🏁 Finished scraping!")

    def update_summary(self, service):
        """Create/update a summary sheet with counts by month."""
        print("📑 Updating Summary sheet...")
        self.create_sheet_if_missing(service, "Summary")

        # Collect all sheet names except "Summary"
        spreadsheet = service.spreadsheets().get(
            spreadsheetId=self.SPREADSHEET_ID
        ).execute()
        sheet_titles = [s["properties"]["title"] for s in spreadsheet.get("sheets", [])]
        data = [["Sheet", "Total Records", "Last Updated"]]

        for title in sheet_titles:
            if title == "Summary":
                continue
            try:
                result = service.spreadsheets().values().get(
                    spreadsheetId=self.SPREADSHEET_ID,
                    range=f"'{title}'!A:A"
                ).execute()
                values = result.get("values", [])
                total = len(values) - 1 if values else 0  # exclude header if present
                data.append([title, total, datetime.now().strftime("%Y-%m-%d %H:%M")])
            except Exception as e:
                print(f"⚠️ Could not read {title}: {e}")

        # Clear and rewrite the Summary sheet
        service.spreadsheets().values().clear(
            spreadsheetId=self.SPREADSHEET_ID,
            range="Summary!A:Z"
        ).execute()
        service.spreadsheets().values().update(
            spreadsheetId=self.SPREADSHEET_ID,
            range="Summary!A1",
            valueInputOption="RAW",
            body={"values": data},
        ).execute()
        print("✅ Summary sheet updated")



if __name__ == "__main__":
    scraper = WillScraper(headless=True)
    scraper.run()


