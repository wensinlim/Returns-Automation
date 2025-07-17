# 📦 Amazon & DPD Returns Automation Script

This project automates the process of retrieving return order details from **Amazon Seller Central** and **DPD Local**, logging the results into a **Google Sheet**.



## ✅ Features

- 🔐 Login automation for Amazon (2FA supported)
- 🔄 Dynamic dropdown selection (e.g. *Tracking ID*)
- 🔎 Search entries from a spreadsheet
- 🧾 Scrape:
  - Order Number
  - Item SKU
  - Return Reason
  - Buyer Comment
- 📤 Write back into the same Google Sheet

---


## 🚀 Getting Started

### 1. Prerequisites

Ensure the following are installed:

- Python 3.9+
- Google Chrome
- ChromeDriver (must match your Chrome version)
- An Amazon Seller Central account
- (Optional) A DPD Local account

### 2. Clone This Repository

bash
git clone https://github.com/your-org/returns-automation.git
cd returns-automation

### 3. Install Dependencies

It's recommended to use a virtual environment:


## 🔐 Google Sheets API Setup

1. Go to [Google Cloud Console](https://console.cloud.google.com/)

2. Create or select a project.

3. Enable the following APIs:
   - Google Sheets API
   - Google Drive API

4. Create **OAuth 2.0 Client ID** credentials:
   - Choose *Desktop app* as the application type.

5. Download the `credentials_oauth.json` file.

6. Place `credentials_oauth.json` in the root folder of this project.

7. Run the script once — this will open a browser window for you to authenticate and generate `token.json` with your access token.


## 🔑 Amazon & DPD Authentication

- When you run the script, you will be prompted to enter your Amazon email and password.
- For Amazon, after entering credentials, you may be prompted to enter a One-Time Password (OTP) if two-factor authentication (2FA) is enabled.
- The script logs into Amazon Seller Central and selects the "United Kingdom" account.
- The DPD login code is present but currently commented out; if you want to use DPD functionality, uncomment and provide your DPD credentials in the script.

---

## 📋 Google Sheet Format

The Google Sheet must have the following columns with headers on the first row:

| Column | Header          | Description                                |
|--------|-----------------|--------------------------------------------|
| B      | Search Value    | The value to search for in Amazon or DPD   |
| C      | Order Number    | (Auto-filled) Order ID from search results |
| D      | Item SKU       | (Auto-filled) SKU from search results       |
| E      | Return Type    | Should be either "Amazon" or "DPD"          |
| F      | Return Reason  | (Auto-filled) Reason for return              |
| G      | Buyer Comment  | (Auto-filled) Comment from buyer             |

- The script processes rows starting from row 2 (after the header).
- Rows with empty search values or unrecognized return types are skipped.


## 🚀 Running the Script

1. Ensure all dependencies are installed (see [Step 2: Install Dependencies](#-step-2-install-dependencies)).

2. Make sure your ChromeDriver version matches your installed Chrome browser.

3. Open a terminal and navigate to the project directory.

4. Run the script with Python 3:

    ```bash
    python returns.py
    ```

5. When prompted, enter your Amazon email and password.

6. If your Amazon account has 2FA enabled, enter the OTP code sent to your device.

7. The script will process the rows in your Google Sheet, perform searches on Amazon (and optionally DPD), and update the sheet with results.

8. Once finished, the browser window will close automatically.

---

**Notes:**
- Ensure your Google Sheets API credentials and tokens are set up correctly before running.
