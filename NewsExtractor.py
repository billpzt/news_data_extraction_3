import time
import logging
import openpyxl

from Utils import Utils
from Locators import Locators as loc

from RPA.Browser.Selenium import By, Selenium
from robocorp.tasks import task
from robocorp import browser

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import StaleElementReferenceException

class NewsExtractor:
    def __init__(self, search_phrase,  months=None, news_category=None, local=True):
        self.search_phrase = search_phrase
        self.months = months
        self.news_category = news_category
        self.browser = browser
        self.base_url = "https://gothamist.com/"
        self.results = []
        self.results_count = 0
        # Configure logging
        self.logger = logging.getLogger('NewsExtractor')
        self.logger.setLevel(logging.INFO)
        handler = logging.FileHandler('news_extractor.log')
        formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        handler.setFormatter(formatter)
        self.logger.addHandler(handler)
        self.local = local

    def open_site(self):
        """Open the news site"""
        browser.goto(self.base_url)
        self.logger.info("Site opened - OK")          

    def click_on_search_button(self):
        """Click on search button to open search bar"""
        page = browser.page()
        page.click(loc.search_button_xpath)
        self.logger.info("Clicked on search button - OK")
                
    def enter_search_phrase(self):
        page = browser.page()
        page.fill(loc.searchbar_xpath, self.search_phrase)
        page.click()
        self.logger.info("Typed search phrase in search bar - OK")

    def filter_newest(self):
        page = browser.page()
        page.select_option(loc.dropdown_xpath, 3)
        # try:
        # # Wait up to 10 seconds before throwing a TimeoutException unless it finds the element to return
        #     WebDriverWait(self.browser.driver, 10).until(
        #         EC.element_to_be_clickable((By.XPATH, loc.dropdown_xpath))
        #     )
        #     self.browser.select_from_list_by_value(loc.dropdown_xpath, "3")
        # except Exception as error:
        #     self.logger.warning(f"Option not available - {str(error)}")
    
    def click_on_news_category(self):
        page = browser.page()
        page.click(loc.category_menu_xpath)
        category_text = self.news_category
        category_checkbox_xpath = f'//div/div/label[span[contains(text(), {category_text})]]/input'
        page.set_checked(category_checkbox_xpath)

    def click_on_next_page(self):
        page = browser.page()
        page.click(loc.next_results_xpath)

    def extract_articles_data(self):
        """Extract data from news articles"""
        articles = self.browser.get_webelements(loc.articles_xpath)
        
        for r in articles:
            browser = self.browser.driver
            months = self.months

            # Capture and convert date string to datetime object
            date, valid_date = Utils.date_extraction_and_validation(browser=browser, months=months, article=r)

            if (valid_date):
                title = Utils.title_extraction(browser=browser, article=r)

                description = Utils.description_extraction(article=r)

                # Download picture if available and extract the filename
                picture_filename = Utils.picture_extraction(self.local, article=r)

                # Count search phrase occurrences in title and description
                count_search_phrases = (title.count(self.search_phrase) + description.count(self.search_phrase))

                # Check if title or description contains any amount of money
                monetary_amount = Utils.contains_monetary_amount(title) or Utils.contains_monetary_amount(description)

                # Store extracted data in a dictionary
                article_data = {
                    "title": title,
                    "date": date,
                    "description": description,
                    "picture_filename": picture_filename,
                    "count_search_phrases": count_search_phrases,
                    "monetary_amount": monetary_amount
                }
                
                self.results.append(article_data)
                self.results_count += 1
            else:
                break

        return valid_date

    def paging_for_extraction(self, goto_next_page=True):
        while (goto_next_page):
            goto_next_page = self.extract_articles_data()
            self.click_on_next_page()
        print(f"Extracted data from {self.results_count} articles")
    
    def close_site(self):
        # Close the browser
        page = browser.page()
        page.close()

    def run(self):
        # Execute the entire news extraction process
        self.open_site()
        self.click_on_search_button()
        self.enter_search_phrase()
        self.click_on_news_category()
        self.filter_newest()
        # self.paging_for_extraction()
        # if self.local:
        #     Utils.LOCAL_save_to_excel(self.results)
        # else:
        #     Utils.save_to_excel(self.results)
        self.close_site()
