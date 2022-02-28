from datetime import *
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import time
import requests
import os
import pandas as pd
import selenium.common.exceptions


class Parser:

    def __init__(self):
        self.s = Service()
        self.chrome_options = Options()
        self.chrome_options.add_argument('--headless')
        self.chrome_options.add_argument('--disable-dev-shm-usage')
        self.chrome_options.add_argument('--no-sandbox')
        self.driver = webdriver.Chrome(service=self.s, options=self.chrome_options)
        self.company_names = []
        self.dates = []
        self.links = []
        self.event_names = []
        self.changed_companies = []
        self.all_news = []
        self.bot_token = os.getenv('TELEGRAM_TOKEN')
        self.chat_id = os.getenv('CHAT_ID')

    def read_excel(self):
        """"This function opens the given Excel file and returns the list of companies"""
        dataframe = pd.read_excel('list_of_companies.xlsx', engine='openpyxl', usecols='L')
        self.company_names = [element for element in dataframe['EMITENT_FULL_NAME'].values]
        return self.company_names

    def open_website(self, keyword):
        """This function opens the website and sends keyword. This keyword is a criterion of filtering news."""
        self.driver.get('https://www.e-disclosure.ru/poisk-po-soobshheniyam')
        time.sleep(3)
        try:
            time.sleep(2)
            event_field = self.driver.find_element(By.ID, "textfieldEvent")
            event_field.send_keys(keyword)
        except selenium.common.exceptions.NoSuchElementException:
            self.send_a_message("This code doesn't work. It seems like one of the selectors has been changed")

    def find_button_and_form(self):
        """"This function is made to find two element, such as Button and Company_Field"""
        try:
            search_button_element = self.driver.find_element(By.ID, "butt")
            company_field_element = self.driver.find_element(By.ID, "textfieldCompany")
        except selenium.common.exceptions.NoSuchElementException:
            self.send_a_message("This code doesn't work. It seems like one of the selectors has been changed")
        return search_button_element, company_field_element

    def look_for_results(self):
        """This function send keys to the input and clicks the button 'Search'. It returns the table with all news and
        company name."""
        company_field.send_keys(company)
        time.sleep(3)
        search_button.click()
        time.sleep(3)
        try:
            results_element = self.driver.find_element(By.XPATH, '//*[@id="searchResults"]/div')
        except selenium.common.exceptions.NoSuchElementException:
            self.send_a_message('Results element has been updated.')
        return results_element, company

    def collect_data(self, name):
        """This function is responsible for parsing data from three last news. It checks the date of publication and if
        it was published not later than 5 years ago, it parses link and the title of document. Otherwise, it skips this
        element. As the result, it return the list of all news."""
        time.sleep(2)
        for i in range(1, 4):
            for element in self.driver.find_elements(By.XPATH, f'//*[@id="cont_wrap"]/table/tbody/tr[{i}]/td[1]'):
                converted_date = datetime.strptime(element.text, '%d.%m.%Y %H:%M')
                if converted_date > datetime.now() - timedelta(hours=24):
                    self.dates.append(converted_date)
                    print('Yes')
                    for link in self.driver.find_elements(By.XPATH,
                                                          f'//*[@id="cont_wrap"]/table/tbody/tr[{i}]/td[2]/a[2]'):
                        self.links.append(link.get_attribute('href'))
                        self.event_names.append(link.text)
                        self.changed_companies.append(name)
                else:
                    print('Yes, but date does not work')
                    company_field.clear()
        print(self.all_news)
        self.all_news = list(zip(self.dates, self.event_names, self.links, self.changed_companies))
        return self.all_news

    def send_a_message(self, msg):
        """This function sends a message. It receives the text as input"""
        send_text = 'https://api.telegram.org/bot' + self.bot_token + '/sendMessage?chat_id=' + self.chat_id + '&text=' + msg
        requests.get(send_text)


parser = Parser()
companies = parser.read_excel()
parser.open_website(keyword='Дивиденды')
search_button, company_field = parser.find_button_and_form()
for company in companies:
    results, company_name = parser.look_for_results()
    if results.text == 'Ничего не найдено.':
        company_field.clear()
        print('Nothing')
    else:
        parser.collect_data(company_name)
if len(parser.all_news) == 0:
    parser.send_a_message(msg='Нет изменений в акциях за последние 5 часов')
else:
    for item in parser.all_news:
        parser.send_a_message(msg=f'Дата публикации: {item[0].strftime("%H:%M")}\nНазвание компании: {item[3]}\n'
                                  f'Название события: {item[1]}\nСсылка: {item[2]}')
