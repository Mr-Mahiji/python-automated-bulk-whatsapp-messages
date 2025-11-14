# -------------------------------------------------------------
# WhatsApp Bulk Messaging Script (Updated for Selenium 4+)
# -------------------------------------------------------------

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import time


# -------------------------------------------------------------
# SETUP SELENIUM + CHROME DRIVER
# -------------------------------------------------------------
service = Service(ChromeDriverManager().install())
options = webdriver.ChromeOptions()

# Optional: keeps WhatsApp session saved
options.add_argument("--user-data-dir=C:/Users/win/AppData/Local/Google/Chrome/User Data")

driver = webdriver.Chrome(service=service, options=options)
wait = WebDriverWait(driver, 30)


# -------------------------------------------------------------
# OPEN WHATSAPP WEB
# -------------------------------------------------------------
driver.get("https://web.whatsapp.com/")
print("\n🔄 Waiting for WhatsApp Web to load... Scan QR if required...\n")
time.sleep(35)  # enough time for login


# -------------------------------------------------------------
# LOAD EXCEL FILE
# -------------------------------------------------------------
excel_data = pd.read_excel("TESTING.xlsx", sheet_name="Customers")


# -------------------------------------------------------------
# FUNCTION TO SEND MESSAGE
# -------------------------------------------------------------
def send_message(contact_number, message):
    try:
        # Click search box
        search_box_xpath = '//div[@contenteditable="true"][@data-tab="3"]'
        search_box = wait.until(
            EC.presence_of_element_located((By.XPATH, search_box_xpath))
        )
        search_box.click()
        search_box.clear()

        # Type number
        search_box.send_keys(contact_number)
        time.sleep(3)

        # Select contact
        search_box.send_keys(Keys.ENTER)
        time.sleep(2)

        # Message box
        message_box_xpath = '//div[@contenteditable="true"][@data-tab="10"]'
        message_box = wait.until(
            EC.presence_of_element_located((By.XPATH, message_box_xpath))
        )
        message_box.send_keys(message)
        message_box.send_keys(Keys.ENTER)

        print(f"✅ Message sent to: {contact_number}")

    except TimeoutException:
        print(f"❌ ERROR: Contact not found or WhatsApp UI slow → {contact_number}")
    except Exception as e:
        print(f"❌ Unexpected error for {contact_number}: {e}")


# -------------------------------------------------------------
# LOOP THROUGH EXCEL ROWS
# -------------------------------------------------------------
for index, row in excel_data.iterrows():
    name = row["Name"]
    number = str(row["Contact"])
    msg_template = row["Message"]

    # Replace placeholder
    final_message = msg_template.replace("{customer_name}", name)

    send_message(number, final_message)
    time.sleep(1)  # avoid flooding


print("\n🎉 ALL MESSAGES SENT SUCCESSFULLY!")
driver.quit()
