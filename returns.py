import os
import time
import getpass
import gspread

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# === CONFIG ===
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
CLIENT_SECRETS_FILE = 'credentials_oauth.json'
TOKEN_FILE = 'token.json'
SPREADSHEET_ID = '1H6DOJIBqjyWMfSjtVeG35OwX97PC-Jv21fLBjAdIXb4'
WORKSHEET_NAME = 'Sheet1'
CHROMEDRIVER_PATH = '/Users/WenSin/Downloads/chromedriver-mac-arm64/chromedriver'

# === Prompt for login credentials ===
DPD_USERNAME = input("Enter your DPD username: ")
DPD_PASSWORD = getpass.getpass("Enter your DPD password: ")
AMAZON_EMAIL = input("Enter your Amazon email: ")
AMAZON_PASSWORD = getpass.getpass("Enter your Amazon password: ")


# === Google Sheets Auth ===
def get_gspread_client():
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    if not creds or not creds.valid:
        flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRETS_FILE, SCOPES)
        creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())
    return gspread.authorize(creds)


# === Login Helpers ===
def login_to_dpd(driver):
    driver.get("https://mydeliveries.dpdlocal.co.uk/login")
    time.sleep(2)
    driver.find_element(By.NAME, "username").send_keys(DPD_USERNAME)
    driver.find_element(By.NAME, "password").send_keys(DPD_PASSWORD)
    for _ in range(10):
        login_btn = driver.find_element(By.ID, "submitBtn")
        if login_btn.is_enabled():
            login_btn.click()
            break
        time.sleep(1)
    time.sleep(5)


def login_to_amazon(driver):
    wait = WebDriverWait(driver, 20)
    driver.get("https://sellercentral.amazon.co.uk/gp/returns/list/v2")

    email_input = wait.until(EC.presence_of_element_located((By.ID, "ap_email")))
    email_input.send_keys(AMAZON_EMAIL)
    email_input.send_keys(Keys.RETURN)

    password_input = wait.until(EC.presence_of_element_located((By.ID, "ap_password")))
    password_input.send_keys(AMAZON_PASSWORD)
    password_input.send_keys(Keys.RETURN)

    otp_input = wait.until(EC.presence_of_element_located((By.ID, "auth-mfa-otpcode")))
    otp_code = input("Enter the OTP sent to your device: ")
    otp_input.send_keys(otp_code)
    otp_input.send_keys(Keys.RETURN)

    try:
        uk_account = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//span[contains(text(), 'United Kingdom')]/ancestor::button")
        ))
        uk_account.click()
        print("✅ Selected 'United Kingdom' account.")
    except Exception as e:
        print(f"⚠️ Failed to select UK account: {e}")
        driver.quit()
        return

    try:
        # Find the 'Select account' button by its unique attribute
        select_account_btn = wait.until(EC.element_to_be_clickable((
            By.CSS_SELECTOR, "[data-test='confirm-selection']"
        )))
        
        # Scroll into view and click with JS for reliability
        driver.execute_script("arguments[0].scrollIntoView(true);", select_account_btn)
        time.sleep(1)
        driver.execute_script("arguments[0].click();", select_account_btn)

        print("✅ Clicked 'Select account' button.")

    except Exception as e:
        print(f"⚠️ Failed to click 'Select account': {e}")
        driver.quit()
        return


    wait.until(EC.url_contains("sellercentral.amazon.co.uk"))


# === Abstract Handler Base ===
class CarrierHandler:
    def __init__(self, driver, row_data, row_index, sheet):
        self.driver = driver
        self.row_data = row_data
        self.row_index = row_index
        self.sheet = sheet

    def search(self):
        raise NotImplementedError("Each subclass must implement its search method.")


# === DPD Blue Handler ===
class DPDBlueHandler(CarrierHandler):
    def search(self):
        search_value = self.row_data[1]
        try:
            self.driver.get("https://mydeliveries.dpdlocal.co.uk/search")
            time.sleep(3)
            self.driver.find_element(By.CLASS_NAME, "searchbar-icon").click()
            time.sleep(1)

            search_input = self.driver.find_element(By.NAME, "searchText")
            search_input.clear()
            search_input.send_keys(search_value)
            search_input.send_keys(Keys.RETURN)
            time.sleep(5)

            td_text = self.driver.find_element(By.ID, "0_Senders ref_0").text.strip()
            before_comma, after_comma = (td_text.split(",", 1) + [""])[:2]

            self.sheet.update_cell(self.row_index, 3, before_comma.strip())
            self.sheet.update_cell(self.row_index, 4, after_comma.strip())
            print(f"✅ DPD Blue → Row {self.row_index}: '{before_comma}', '{after_comma}'")
        except Exception as e:
            print(f"❌ DPD Blue failed on row {self.row_index}: {e}")


# === DPD Red Handler ===
class DPDRedHandler(CarrierHandler):
    def search(self):
        tracking_id = self.row_data[1].strip() if len(self.row_data) > 1 else ""
        if not tracking_id:
            print(f"Row {self.row_index}: Blank tracking number.")
            return

        try:
            self.driver.get("https://www.dpd.co.uk/")
            wait = WebDriverWait(self.driver, 10)

            input_elem = wait.until(EC.presence_of_element_located((By.ID, "receiver_reference")))
            input_elem.clear()
            input_elem.send_keys(tracking_id)
            input_elem.send_keys(Keys.RETURN)
            print(f"🔍 DPD Red: Submitted tracking ID '{tracking_id}'")
            time.sleep(3)

            postcode_input = wait.until(EC.presence_of_element_located((By.NAME, "postcode")))
            postcode_input.clear()
            postcode_input.send_keys("UB2 4AB")
            postcode_input.send_keys(Keys.RETURN)
            print("📮 DPD Red: Submitted postcode")
            time.sleep(5)

            parcel_info_btn = wait.until(EC.element_to_be_clickable((
                By.XPATH, "//span[text()='Parcel info']/parent::button"
            )))
            parcel_info_btn.click()
            print("📦 DPD Red: Clicked 'Parcel info' button")
            time.sleep(3)

            consignment_td = wait.until(EC.presence_of_element_located((
                By.XPATH, "(//table)[2]//tbody/tr[1]/td[2]"
            )))
            consignment_text = consignment_td.text.strip()
            parts = [p.strip() for p in consignment_text.split(",")]

            middle_value = parts[1] if len(parts) >= 3 else ""

            self.sheet.update_cell(self.row_index, 8, middle_value)
            print(f"✅ DPD Red → Row {self.row_index} updated with consignment number")
        except Exception as e:
            print(f"❌ DPD Red row {self.row_index} error: {e}")


# === Amazon Handler ===
class AmazonHandler(CarrierHandler):
    def __init__(self, driver, row_data, row_index, sheet, search_mode, login_needed=False):
        super().__init__(driver, row_data, row_index, sheet)
        self.login_needed = login_needed
        self.search_mode = search_mode  # Either "Tracking ID" or "RMA ID"

    def search(self):
        if self.login_needed:
            login_to_amazon(self.driver)

        wait = WebDriverWait(self.driver, 20)

        # === Step 1: Select correct dropdown option ===
        try:
            dropdown_container = wait.until(EC.presence_of_element_located(
                (By.CSS_SELECTOR, "kat-dropdown.search-by-filter-dropdown")
            ))
            dropdown_trigger = dropdown_container.find_element(By.CSS_SELECTOR, ".select-header")
            self.driver.execute_script("arguments[0].click();", dropdown_trigger)

            WebDriverWait(dropdown_container, 10).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".option-row"))
            )

            options = dropdown_container.find_elements(By.CSS_SELECTOR, ".option-row")
            matched_option = None
            for opt in options:
                if opt.get_attribute("data-name") == self.search_mode:
                    matched_option = opt
                    break

            if matched_option:
                self.driver.execute_script("arguments[0].click();", matched_option)
                print(f"✅ Selected '{self.search_mode}' from dropdown.")
            else:
                print(f"❌ Option '{self.search_mode}' not found in dropdown.")

        except Exception as e:
            print(f"❌ Failed to select search mode '{self.search_mode}': {e}")
            return

        # === Step 2: Perform search with value from Column B ===
        try:
            search_value = self.row_data[1]
            search_input = wait.until(EC.presence_of_element_located((
                By.XPATH, "//div[contains(@class, 'search-field') and contains(@class, 'mr-search-box')]//input[@id='katal-id-3']"
            )))
            search_input.clear()
            search_input.send_keys(search_value)
            search_input.send_keys(Keys.RETURN)
            time.sleep(4)

            # === Step 3: Extract details ===
            order_number = self.driver.find_element(
                By.CSS_SELECTOR, ".orderDetailsDiv a[href*='/orders-v3/order/']"
            ).text.strip()

            sku = self.driver.find_element(
                By.XPATH, "//kat-label[@emphasis='Merchant SKU:']//span[@class='text']"
            ).text.strip()

            reason = self.driver.find_element(
                By.CSS_SELECTOR, ".return-reason-value"
            ).text.strip()

            comment = self.driver.find_element(
                By.CSS_SELECTOR, ".customer-comment"
            ).text.strip()

            # === Step 4: Write values to Sheet ===
            self.sheet.update_cell(self.row_index, 3, order_number)  # Column C
            self.sheet.update_cell(self.row_index, 4, sku)           # Column D
            self.sheet.update_cell(self.row_index, 6, reason)        # Column F
            self.sheet.update_cell(self.row_index, 7, comment)       # Column G

            print(f"✅ Amazon → Row {self.row_index} updated with order, SKU, reason, and comment.")

        except Exception as e:
            print(f"❌ Amazon failed for row {self.row_index}: {e}")


# === Main Search Logic ===
def main():
    client = get_gspread_client()
    sheet = client.open_by_key(SPREADSHEET_ID).worksheet(WORKSHEET_NAME)
    data = sheet.get_all_values()[1:]  # skip header
    start_row = 2

    options = Options()
    options.add_argument("--start-maximized")
    service = ChromeService(executable_path=CHROMEDRIVER_PATH)
    driver = webdriver.Chrome(service=service, options=options)

    amazon_logged_in = False

    try:
        login_to_dpd(driver)

        for idx, row in enumerate(data, start=start_row):
            return_type = row[4] if len(row) > 4 else ""
            if not row or not row[1].strip():
                continue

            handler = None

            if return_type == "DPD Blue":
                handler = DPDBlueHandler(driver, row, idx, sheet)
            elif return_type == "DPD Red":
                handler = DPDRedHandler(driver, row, idx, sheet)
                # After DPD Red: Use Column H (index 7) to search on Amazon orders page
                consignment_id = row[7].strip() if len(row) > 7 else ""
                if consignment_id:
                    try:
                        if not amazon_logged_in:
                            login_to_amazon(driver)
                            amazon_logged_in = True

                        # Step 1: Go to orders page
                        driver.get("https://sellercentral.amazon.co.uk/orders-v3")
                        wait = WebDriverWait(driver, 20)

                        # Step 2: Select "Tracking ID" from dropdown
                        dropdown_container = wait.until(EC.presence_of_element_located(
                            (By.CSS_SELECTOR, "kat-dropdown.search-by-filter-dropdown")
                        ))
                        dropdown_trigger = dropdown_container.find_element(By.CSS_SELECTOR, ".select-header")
                        driver.execute_script("arguments[0].click();", dropdown_trigger)
                        WebDriverWait(dropdown_container, 10).until(
                            EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".option-row"))
                        )
                        for opt in dropdown_container.find_elements(By.CSS_SELECTOR, ".option-row"):
                            if opt.get_attribute("data-name") == "Tracking ID":
                                driver.execute_script("arguments[0].click();", opt)
                                print("✅ Selected 'Tracking ID'")
                                break

                        # Step 3: Search with consignment ID
                        search_input = wait.until(EC.presence_of_element_located(
                            (By.CSS_SELECTOR, "input[placeholder='Search']")
                        ))
                        search_input.clear()
                        search_input.send_keys(consignment_id)
                        search_input.send_keys(Keys.RETURN)
                        print(f"🔍 Searched Amazon Orders with consignment: {consignment_id}")
                        time.sleep(4)

                        # Step 4: Extract order number
                        order_link = wait.until(EC.presence_of_element_located(
                            (By.CSS_SELECTOR, "a[href*='/orders-v3/order/']")
                        ))
                        order_number = order_link.text.strip()

                        # Step 5: Extract SKU
                        sku_elem = wait.until(EC.presence_of_element_located(
                            (By.XPATH, "//span[text()='SKU']/following-sibling::b")
                        ))
                        sku = sku_elem.text.strip()

                        # Step 6: Update Google Sheet
                        sheet.update_cell(idx, 3, order_number)  # Column C
                        sheet.update_cell(idx, 4, sku)           # Column D
                        print(f"✅ Amazon post-DPD Red: Row {idx} updated → Order: {order_number}, SKU: {sku}")

                    except Exception as e:
                        print(f"❌ Amazon post-DPD Red search failed on row {idx}: {e}")

            elif return_type.startswith("Amazon"):
                if not amazon_logged_in:
                    login_to_amazon(driver)
                    amazon_logged_in = True

                if return_type == "Amazon RMA":
                    search_mode = "RMA ID"
                elif return_type == "Amazon Tracking":
                    search_mode = "Tracking ID"
                else:
                    print(f"⚠️ Unknown Amazon subtype on row {idx}: '{return_type}'")
                    continue

                handler = AmazonHandler(driver, row, idx, sheet, search_mode, login_needed=False)
                handler.search()

            if handler:
                handler.search()

    finally:
        driver.quit()


if __name__ == "__main__":
    main()
