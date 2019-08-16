import csv
import pandas as pd
import os
from bs4 import BeautifulSoup
import xlsxwriter
import time
from selenium import webdriver
from fake_useragent import UserAgent
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
import timeit
import multiprocessing

#  Change according to the homepage of the site
Homepage = 'https://www.gps-coordinates.net/'

user = UserAgent().random
headers = {'User-Agent': user}

# Chrome display set up
# NOTE: DO NOT ADD --headless argument
options = webdriver.ChromeOptions()
options.add_argument('--ignore-certificate-errors')
options.add_argument("--disable-notifications")
prefs = {"profile.default_content_setting_values.geolocation" : 2}
options.add_experimental_option("prefs", prefs)
options.add_argument("--test-type")
options.add_argument(f'user-agent={user}')
options.add_argument('--disable-gpu')


# create new function () input : address and institution name; output : latitude and longitude values

def search_by_name_and_address(institution_list_data):  # compare country with id window box
    input_string_name = institution_list_data[0]
    input_string_address = institution_list_data[9]
    driver = webdriver.Chrome(options=options,
                              executable_path=r'C:\Users\Nicholas\Documents\Summer intern @ Seeka\chromedriver.exe')
    driver.get(Homepage)

    # look for the search by address input
    element = driver.find_element_by_id("address")
    time.sleep(0.5)
    element.send_keys(Keys.CONTROL + "a")
    time.sleep(0.5)
    element.send_keys(Keys.DELETE)
    time.sleep(0.5)
    element.send_keys(input_string_name)
    time.sleep(3)

    # add action to handle drop down menu
    driver.find_element_by_xpath("/html/body/div[1]/div/a[1]").click()
    driver.find_element_by_class_name("btn").click()
    time.sleep(1)

    # error handling try catch
    # Details = ['Institution Name', 'Provided Address', 'Matched Address using address search',
    # 'Latitude_address', 'Longitude_address', 'Matched Address using name search', 'Latitude_name', 'Longitude_name']
    details = ['', '', '', '', '', '', '', '']
    details[0] = input_string_name

    try:
        # EC.alert_is_present: # popup exist, code returns N/A for lat and long
        time.sleep(2)
        alert = driver.switch_to.alert
        alert.accept()
        driver.quit()
        details[2], details[3], details[4] = 0, 0, 0

    except:
        try:
            delay = 10
            myElem = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.ID, 'info_window')))
            soup = BeautifulSoup(driver.page_source, "lxml")
            x = soup.find(id='info_window').get_text()
            y = x.split("Latitude: ")
            details[2] = y[0].lstrip(' ')
            z = y[1].strip("Get Altitude").split(" | Longitude: ")
            details[3] = z[0]
            details[4] = z[1]

            driver.quit()

        except TimeoutException:
            print("Loading took too much time!")

    return details


def multi_pool(func, input_name_list, procs):
    templist = []
    # counter = len(input_name_list)
    pool = multiprocessing.Pool(processes=procs)
    # print('Total number of processes: ' + str(procs))
    for a in pool.imap(func, input_name_list):
        templist.append(a)
        # print('Number of links left: ' + str(counter - len(templist)))
    pool.terminate()
    pool.join()
    return templist


def main():
    start = timeit.default_timer()

    # move open files into main as well, then add multiple processes, then add the search by name feature
    # dynamic pathname based on different device, instead of hard coding the pathname
    institution_list_path = os.path.join(os.getcwd(), 'institution-details.xlsx')
    test_output_path = os.path.join(os.getcwd(), 'test-output.xlsx')

    # open the final_file.xlsx,  drop rows which university name values are null
    rawdata = pd.read_excel(institution_list_path)
    rawdata = rawdata.dropna(axis=0, subset=('Univeristy_Name',))
    institution_list_data = rawdata.values.tolist()
    #print(institution_list_data)

    # Multiprocessing Collect_Data()
    all_data = multi_pool(search_by_name_and_address, institution_list_data, 10)


    #writing into an output file
    with open('C:/Users/Nicholas/Documents/Summer intern @ Seeka/get-gps-coordinates/output-institutions.csv', 'wt', encoding="utf-8", newline='') as website:
        writer = csv.writer(website)
        print("Writing details to CSV File now....")
        for a in all_data:
            writer.writerow(a)
    print("Total number of rows written to test-output-1 file: " + str(len(all_data)))

    stop = timeit.default_timer()
    time_sec = stop - start
    time_min = int(time_sec / 60)
    time_hour = int(time_min / 60)

    time_run = str(format(time_hour, "02.0f")) + ':' + str(
        format((time_min - time_hour * 60), "02.0f") + ':' + str(format(time_sec - (time_min * 60), "^-05.1f")))
    print("This code has completed running in: " + time_run)

if __name__ == '__main__':
    main()
