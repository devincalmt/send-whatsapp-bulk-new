# Program to send bulk messages through WhatsApp web from an excel sheet without saving contact numbers
# Author @inforkgodara

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from time import sleep
import pandas
import urllib.parse
import os
import platform

if platform.system() == 'Darwin':
    # MACOS Path
    chrome_default_path = os.getcwd() + '\\driver\\chromedriver'
else:
    # Windows Path
    chrome_default_path = os.getcwd() + '\\driver\\chromedriver.exe'

excel_data = pandas.read_excel('Recipients data.xlsx', sheet_name='Recipients')

count = 0

chrome_options = Options()
chrome_options.add_argument('--user-data-dir={}'.format(os.getcwd()+'\\User_Data'))

driver = webdriver.Chrome(executable_path=chrome_default_path, options=chrome_options)
driver.get('https://web.whatsapp.com')
input("Press ENTER after login into Whatsapp Web and your chats are visiable.")

fail_list = []

for column in excel_data['Contact'].tolist():
    try:
        url = 'https://web.whatsapp.com/send?phone=' + str(int(excel_data['Contact'][count])) + '&text=' + urllib.parse.quote(excel_data['Message'][count])
        sent = False
        # It tries 3 times to send a message in case if there any error occurred
        driver.get(url)
        try:
            click_btn = WebDriverWait(driver, 60).until(
                EC.element_to_be_clickable((By.CLASS_NAME, '_4sWnG')))
        except Exception as e:
            fail_list.append(excel_data['Contact'][count])
            print("Sorry message could not sent to " + str(excel_data['Contact'][count]))
        else:
            sleep(2)
            click_btn.click()
            sent = True
            sleep(5)
            print('Message sent to: ' + str(excel_data['Contact'][count]))
        count = count + 1
    except Exception as e:
        fail_list.append(excel_data['Contact'][count])
        print('Failed to send message to ' + str(excel_data['Contact'][count]) + str(e))
driver.quit()
print("The script executed successfully.")

if len(fail_list) > 0:
    print("List Fail to Send")
    for fail in fail_list:
        print(fail)
