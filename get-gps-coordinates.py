import pandas as pd
# import selenium, use headless
# import requests
import os
from bs4 import BeautifulSoup
import requests
import timeit
import xlsxwriter
import time
from selenium import webdriver
from fake_useragent import UserAgent
#pip install xlrd


# dynamic pathname based on different device, instead of hard coding the pathname
from selenium.webdriver.common.keys import Keys

institution_list_path = os.path.join(os.getcwd(), 'institution-details.xlsx')
test_output_path = os.path.join(os.getcwd(), 'test-output.xlsx')
print(institution_list_path)

# open the final_file.xlsx
institution_list_data = pd.read_excel(institution_list_path)
test_output_data = pd.read_excel(test_output_path)

# drop rows which university name values are null,
institution_list_data = institution_list_data.dropna(axis=0, subset=('Univeristy_Name', ))

address_list = institution_list_data.iloc[:, 9]

#  Change according to the homepage of the site
Homepage = 'https://www.gps-coordinates.net/'

user = UserAgent().random
headers = {'User-Agent': user}

# Setup Chrome display
options = webdriver.ChromeOptions()
options.add_argument('--ignore-certificate-errors')
options.add_argument("--disable-notifications")
prefs = {"profile.default_content_setting_values.geolocation" : 2}
options.add_experimental_option("prefs", prefs)
options.add_argument("--test-type")
options.add_argument(f'user-agent={user}')
options.add_argument('--disable-gpu')
driver = webdriver.Chrome(chrome_options=options, executable_path=r'C:\Users\Nicholas\Documents\Summer intern @ Seeka\chromedriver.exe')
driver.get(Homepage)

#look for the address input line
element = driver.find_element_by_id("address")
element.send_keys(Keys.CONTROL + "a")
element.send_keys(Keys.DELETE)
element.send_keys('Kepong')


