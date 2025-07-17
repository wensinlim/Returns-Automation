import os
import time
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
import getpass

# ===== CONFIG =====
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
CLIENT_SECRETS_FILE = 'credentials_oauth.json'
TOKEN_FILE = 'token.json'
SPREADSHEET_ID = '1H6DOJIBqjyWMfSjtVeG35OwX97PC-Jv21fLBjAdIXb4'
WORKSHEET_NAME = 'Sheet1'
CHROMEDRIVER_PATH = '/Users/WenSin/Downloads/chromedriver-mac-arm64/chromedriver'

# === Prompt for DPD & Amazon credentials ===
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

# === DPD Login ===
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

# === Amazon Login + Dropdown Setup ===
def login_to_amazon(driver):
    wait = WebDriverWait(driver, 20)
    print("Logging into Amazon...")
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

    # Step 1: Select 'United Kingdom' account
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

    # Step 2: Click 'Select account' button
    try:
        select_account_btn = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//button[contains(text(), 'Select account')]")
        ))
        select_account_btn.click()
        print("✅ Clicked 'Select account' button.")
        time.sleep(3)
    except Exception as e:
        print(f"⚠️ Failed to click 'Select account': {e}")
        driver.quit()
        return

    wait.until(EC.url_contains("sellercentral.amazon.co.uk"))

    # Locate the specific dropdown container (Search By)
    dropdown_container = wait.until(EC.presence_of_element_located(
        (By.CSS_SELECTOR, "kat-dropdown.search-by-filter-dropdown")
    ))

    # Then locate the clickable header *within* that container
    dropdown_trigger = dropdown_container.find_element(By.CSS_SELECTOR, ".select-header")
    driver.execute_script("arguments[0].click();", dropdown_trigger)

    # Wait until option rows are rendered
    WebDriverWait(dropdown_container, 10).until(
        EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".option-row"))
    )

    # Now find the 'Tracking ID' option manually
    options = dropdown_container.find_elements(By.CSS_SELECTOR, ".option-row")
    tracking_option = None
    for opt in options:
        if opt.get_attribute("data-name") == "Tracking ID":
            tracking_option = opt
            break

    if tracking_option:
        driver.execute_script("arguments[0].click();", tracking_option)
        print("✅ Successfully selected 'Tracking ID'")
    else:
        print("❌ 'Tracking ID' option not found.")

# === Main Logic ===
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
        driver.get("https://mydeliveries.dpdlocal.co.uk/search")
        time.sleep(3)

        for idx, row in enumerate(data, start=start_row):
            search_value = row[1] if len(row) > 1 else ""
            return_type = row[4] if len(row) > 4 else ""

            if not search_value.strip():
                print(f"Row {idx}: Blank value in column B — stopping.")
                break

            if return_type.strip().upper() == "DPD":
                try:
                    search_button = driver.find_element(By.CLASS_NAME, "searchbar-icon")
                    search_button.click()
                    time.sleep(1)

                    search_input = driver.find_element(By.NAME, "searchText")
                    search_input.clear()
                    search_input.send_keys(search_value)
                    search_input.send_keys(Keys.RETURN)
                    print(f"Searched for (DPD): {search_value}")
                    time.sleep(5)

                    try:
                        td_element = driver.find_element(By.ID, "0_Senders ref_0")
                        td_text = td_element.text.strip()

                        if "," in td_text:
                            before_comma, after_comma = [s.strip() for s in td_text.split(",", 1)]
                        else:
                            before_comma, after_comma = td_text, ""

                        sheet.update_cell(idx, 3, before_comma)
                        sheet.update_cell(idx, 4, after_comma)
                        print(f" → Wrote to C{idx}: '{before_comma}', D{idx}: '{after_comma}'")

                    except Exception as e:
                        print(f"Could not find <td> for '{search_value}': {e}")

                except Exception as e:
                    print(f"DPD error on row {idx}: {e}")

            elif return_type.strip().upper() == "AMAZON":
                if not amazon_logged_in:
                    login_to_amazon(driver)
                    amazon_logged_in = True
                print(f"Row {idx}: Ready to run Amazon search for '{search_value}'")

                # Wait for search bar input field and enter search_value
                wait = WebDriverWait(driver, 20)
                search_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[placeholder='Search']")))
                search_input.clear()
                search_input.send_keys(search_value)
                search_input.send_keys(Keys.RETURN)
                print(f"🔍 Searched for (Amazon): {search_value}")

                # Wait for search results to load (adjust if too short)
                time.sleep(4)

                try:
                    # --- ORDER NUMBER (Column C) ---
                    order_elem = driver.find_element(By.CSS_SELECTOR, ".orderDetailsDiv a[href*='/orders-v3/order/']")
                    order_number = order_elem.text.strip()

                    # --- ITEM SKU (Column D) ---
                    sku_elem = driver.find_element(By.XPATH, "//kat-label[@emphasis='Merchant SKU:']//span[@class='text']")
                    item_sku = sku_elem.text.strip()

                    # --- RETURN REASON (Column F) ---
                    reason_elem = driver.find_element(By.CSS_SELECTOR, ".return-reason-value")
                    return_reason = reason_elem.text.strip()

                    # --- BUYER COMMENT (Column G) ---
                    comment_elem = driver.find_element(By.CSS_SELECTOR, ".customer-comment")
                    buyer_comment = comment_elem.text.strip()

                    # Update the sheet
                    sheet.update_cell(idx, 3, order_number)     # Column C
                    sheet.update_cell(idx, 4, item_sku)         # Column D
                    sheet.update_cell(idx, 6, return_reason)    # Column F
                    sheet.update_cell(idx, 7, buyer_comment)    # Column G

                    print(f"✅ Row {idx} updated → Order: {order_number}, SKU: {item_sku}, Reason: {return_reason}, Comment: {buyer_comment}")

                except Exception as e:
                    print(f"❌ Error extracting info for row {idx}: {e}")

                try:
                    search_input = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.ID, "katal-id-3"))
                    )
                    search_input.clear()
                    search_input.send_keys(search_value)
                    search_input.send_keys(Keys.RETURN)
                    print(f" → Performed Amazon search for: {search_value}")
                    time.sleep(4)
                except Exception as e:
                    print(f"⚠️ Amazon search failed for '{search_value}': {e}")

            else:
                print(f"Row {idx}: Unknown return type '{return_type}', skipping.")

    finally:
        driver.quit()

if __name__ == "__main__":
    main()
