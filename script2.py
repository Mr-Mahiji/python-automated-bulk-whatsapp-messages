# Program to send bulk customized messages through WhatsApp Web
# Author: Updated for Selenium 4+

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import time

# Load Chrome driver
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=Service(ChromeDriverManager(version="142.0.7444.162").install()))
#driver = webdriver.Chrome(service=service)
wait = WebDriverWait(driver, 20)

# Open WhatsApp Web
driver.get("https://web.whatsapp.com/")
print("Scan the QR code if not already logged in...")
time.sleep(15)  # Give time to scan QR code

# Load Excel data
excel_data = pd.read_excel('TESTING.xlsx', sheet_name='Customers')

# Iterate over each contact
for index, row in excel_data.iterrows():
    contact_number = str(row['Contact'])
    customer_name = row['Name']
    message_template = row['Message']

    # Replace placeholder with actual customer name
    message = message_template.replace("{customer_name}", customer_name)

    # Locate WhatsApp search box
    search_box_xpath = '//*[@id="side"]/div[1]/div/label/div/div[2]'
    try:
        search_box = wait.until(EC.presence_of_element_located((By.XPATH, search_box_xpath)))
        search_box.clear()
        search_box.send_keys(contact_number)
        time.sleep(3)  # Wait for search results to load

        # Select the contact
        search_box.send_keys(Keys.ENTER)
        time.sleep(2)

        # Send the message
        actions = ActionChains(driver)
        actions.send_keys(message)
        actions.send_keys(Keys.ENTER)
        actions.perform()

        print(f"Message sent to {customer_name} ({contact_number})")
        time.sleep(1)

    except NoSuchElementException:
        print(f"Contact not found: {contact_number}")
    except Exception as e:
        print(f"Error sending message to {contact_number}: {e}")

# Close browser
driver.quit()
print("All messages sent.")
