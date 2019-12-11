import csv
import datetime
import sys
import time

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup


# args
args = sys.argv
if len(args) > 2:
    print("Invalid arguments!")
    print(r"Usage: .\collect_zendesk_my.py")
    print(r"Usage: .\collect_zendesk_my.py 7")
    exit(1)

if len(args) == 2:
    days_to_collect = int(args[1])
else:
    days_to_collect = 99999


# you must set first
email = 'yuhashimoto@asgent.co.jp'
password = 'nishikiori@1'


# log file
current_date = datetime.datetime.now().strftime('%Y%m%d')
log_name = "zendesk_my_{}.csv".format(current_date)


# zendesk login
driver = webdriver.Chrome()
driver.set_page_load_timeout(10)
driver.get('https://votiro.zendesk.com/')
form = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "/html/body/div/iframe")))
driver.switch_to.frame(form)
driver.find_element_by_id('user_email').send_keys(email)
driver.find_element_by_id('user_password').send_keys(password)
login_button = driver.find_element_by_xpath('//*[@id="login-form"]/input[9]')
login_button.click()

# scraping first page
driver.get('https://votiro.zendesk.com/hc/en-us/requests')
page_source = driver.page_source
soup = BeautifulSoup(page_source, 'lxml')

print('Writing to csv ...')
try:
    with open(log_name, 'w', encoding='utf-8_sig') as f:
        writer = csv.writer(f, lineterminator='\n')
        # title ('Subject', 'Id', 'Created', 'Last activity', 'Status')
        title = soup.find('table').find('thead').find('tr')
        csv_data_title = []
        csv_data_title.append(title.find_all('th')[0].text.strip())
        csv_data_title.append(title.find_all('th')[1].text.strip())
        csv_data_title.append(title.find_all('th')[2].text.strip())
        title.find_all('th')[3].find('span').decompose()
        csv_data_title.append(title.find_all('th')[3].text.strip())
        csv_data_title.append(title.find_all('th')[4].text.strip())
        writer.writerow(csv_data_title)

        # recode
        rows = soup.find('table').find('tbody').find_all('tr')
        for row in rows:
            csv_data = []
            row.find_all('td')[0].find('div').decompose()
            csv_data.append(row.find_all('td')[0].text.strip())
            csv_data.append(row.find_all('td')[1].text.strip())
            csv_data.append(row.find_all('td')[2].text.strip())
            csv_data.append(row.find_all('td')[3].find('time')['title'])
            csv_data.append(row.find_all('td')[4].text.strip())
            print(csv_data)
            writer.writerow(csv_data)
except OSError as e:
    print('OS Error: {}'.format(e))
else:
    print('Writing to csv is successfully ended.')

time.sleep(1)


# scraping second page and more
counter = 2
flag = False
while True:
    url = "https://votiro.zendesk.com/hc/en-us/requests?page={}#requests".format(counter)
    driver.get(url)
    page_source = driver.page_source
    soup = BeautifulSoup(page_source, 'lxml')
    if soup.find_all('table'):
        # print("This page is {}.".format(counter))
        print('Writing to csv ...')
        try:
            with open(log_name, 'a', encoding='utf-8_sig') as f:
                writer = csv.writer(f, lineterminator='\n')
                rows = soup.find('table').find('tbody').find_all('tr')
                for row in rows:
                    now = datetime.datetime.now()
                    record_time = row.find_all('td')[3].find('time')['title']
                    if (datetime.datetime.strptime(record_time, '%Y-%m-%d %H:%M') + datetime.timedelta(days=days_to_collect)) <= now:
                        flag = True
                        break

                    csv_data = []
                    row.find_all('td')[0].find('div').decompose()
                    csv_data.append(row.find_all('td')[0].text.strip())
                    csv_data.append(row.find_all('td')[1].text.strip())
                    csv_data.append(row.find_all('td')[2].text.strip())
                    csv_data.append(record_time)
                    csv_data.append(row.find_all('td')[4].text.strip())
                    print(csv_data)
                    writer.writerow(csv_data)
        except OSError as e:
            print('OS Error: {}'.format(e))
        else:
            print('Writing to csv is successfully ended.')
            if flag:
                break
    else:
        print("No requests found.")
        break

    time.sleep(1)
    counter += 1


# finalize
driver.quit()

