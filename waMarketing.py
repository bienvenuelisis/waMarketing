# --------------------------------------------------------------------
#  waMarketing : IndabaX Togo automation tool to send bulk WhatsApp message
# --------------------------------------------------------------------

import pandas as pd
from time import sleep
from urllib.parse import quote
from selenium import webdriver 
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

print("Choose your prefereed browser : \n 1. Google Chrome \n 2. Microsoft Edge \n")
print("Your choice : ", end="")

choice = int(input())

driver = None

if choice == 1:
    driver = webdriver.Chrome(service=webdriver.ChromeService(executable_path='chromedriver.exe'))
elif choice == 2:
    driver = webdriver.Edge(service=webdriver.EdgeService(executable_path='msedgedriver.exe'))
else:
    print("Invalid choice. Please choose 1 or 2.")
    exit()

print("Enter the path of your Excel file (XLS, XLSX) : ", end="")
pathToExcelFile = input()

if pathToExcelFile == "":
    print("Invalid path. Please provide a valid path.")
    exit()
    
if not pathToExcelFile.endswith('.xlsx') and not pathToExcelFile.endswith('.xls'):
    print("Invalid file format. Please provide a valid Excel file.")
    exit()

# Load data from Excel
print("Loading data from Excel : " + pathToExcelFile)
data = pd.read_excel(pathToExcelFile, sheet_name='Sheet0')


# Open WhatsApp Web
driver.get('https://web.whatsapp.com')

input("Press ENTER after logging into WhatsApp Web and when your chats are visible.\n\n\n")

# Loop through the contacts in the Excel file
for index, row in data.iterrows():
    name = row.get("Name")
    message = row.get("Message")
    number = row.get("Number")
    
    # Check if number and message exist
    if pd.isna(number) or pd.isna(message):
        print(f"Skipping row {index + 1}: Missing number or message.")
        continue
    
    # URL encode the message for safe URL usage
    encoded_message = quote(str(message))
    url = f'https://web.whatsapp.com/send?phone={str(number)}&text={encoded_message}'
    
    try:
        # Navigate to the URL
        driver.get(url)
        
        # Wait for the message box to load
        message_box = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, "//div[@data-testid='conversation-compose-box-input']"))
        )
        
        js_executor = driver.execute_script("return arguments[0]", message_box)

        # Simulate Enter key press using JavaScript
        js_executor.send_keys(Keys.ENTER)

        print(f'Message sent to: {number}\n')
        
        sleep(5)
    
    except Exception as e:
        print(f"Failed to send message to {number}\n")

# Close WebDriver after sending all messages
driver.quit()
print("The script executed successfully.\n")
