import os
import re
import requests

import openpyxl
from robocorp import storage

from dateutil.parser import parse
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta

from RPA.Browser.Selenium import By, Selenium
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from Locators import Locators as loc

class Utils:
    def date_formatter(date_str):
        try:
        # Attempt to parse the input date string
            parsed_date = parse(date_str)
        # If successful, return the date part of the parsed date
            return parsed_date.date()
        except ValueError:
        # If parsing fails, return today's date
            return datetime.today().date()

    def date_checker(date_to_check, months):
        number_of_months = months
        if (months == 0):
            number_of_months = 1
        # Calculate the date 'months' months ago from today
        cutoff_date = datetime.today().date() - relativedelta(months=number_of_months)
        # Check if the date_to_check is within the last 'months' months
        return date_to_check >= cutoff_date
    
    def contains_monetary_amount(text):
        # Check if the text contains any monetary amount (e.g. $11.1, 11 dollars, etc.)
        pattern = r"\$?\d+(\.\d{2})? dollars?|USD|euro|€"
        return bool(re.search(pattern, text))

    def LOCAL_download_picture(picture_url):
        # Prepare the local path for the picture
        output_dir = os.path.join(os.getcwd(), "output", "images")  # Using current working directory
        os.makedirs(output_dir, exist_ok=True)  # Ensure the directory exists

        sanitized_filename = re.sub(r'[\\/*?:"<>|]', "", os.path.basename(picture_url))
        filename_root, filename_ext = os.path.splitext(sanitized_filename)
        if filename_ext.lower() != '.jpg':
            sanitized_filename += ".jpg"
        picture_filename = os.path.join(output_dir, sanitized_filename)

        try:
            # Download the picture
            response = requests.get(picture_url, stream=True)
            if response.status_code == 200:
                with open(picture_filename, 'wb') as file:
                    for chunk in response.iter_content(chunk_size=1024):
                        if chunk:
                            file.write(chunk)
                print(f"Picture downloaded successfully: {picture_filename}")
                return picture_filename
            else:
                print(f"Failed to download picture from {picture_url}. Status code: {response.status_code}")
        except Exception as e:
            print(f"Error downloading picture: {e}")

        return None

    def download_picture(picture_url):
        # Prepare the local path for the picture
        output_dir = os.path.join(os.getcwd(), "output", "images")  # Using current working directory
        os.makedirs(output_dir, exist_ok=True)  # Ensure the directory exists

        sanitized_filename = re.sub(r'[\\/*?:"<>|]', "", os.path.basename(picture_url))
        filename_root, filename_ext = os.path.splitext(sanitized_filename)
        if filename_ext.lower() != '.jpg':
            sanitized_filename += ".jpg"
        picture_filename = os.path.join(output_dir, sanitized_filename)

        try:
            # Download the picture
            response = requests.get(picture_url, stream=True)
            if response.status_code == 200:
                with open(picture_filename, 'wb') as file:
                    for chunk in response.iter_content(chunk_size=1024):
                        if chunk:
                            file.write(chunk)
                print(f"Picture downloaded successfully: {picture_filename}")
                
                # Store the picture as an asset in Control Room
                asset_name = sanitized_filename  # You can customize the asset name as needed
                storage.set_file(asset_name, picture_filename)
                print(f"Picture stored as asset: {asset_name}")
                
                return picture_filename
            else:
                print(f"Failed to download picture from {picture_url}. Status code: {response.status_code}")
        except Exception as e:
            print(f"Error downloading picture: {e}")

        return None
    
    def picture_extraction(local, article):
        try:
            e_img = article.find_element(By.XPATH, loc.article_image_xpath)
            print(f"Encontrou imagem: {e_img}")
        except:
            picture_url = ''
            picture_filename = ''
            print("Image not found")
        else:
            picture_url = e_img.get_attribute("src")
            if local:
                picture_filename = Utils.LOCAL_download_picture(picture_url)
            else:
                picture_filename = Utils.download_picture(picture_url)
            print(f"URL da imagem: {picture_url}")
            print(f"Nome do arquivo: {picture_filename}")
        return picture_filename
    
    def date_extraction_and_validation(browser, months, article):
        # WebDriverWait(browser, 20).until(
        #         EC.presence_of_element_located((By.XPATH, loc.article_date_xpath))
        #     )
        raw_date = article.find_element(By.XPATH, loc.article_date_xpath).text
        date = Utils.date_formatter(date_str=raw_date)
        months = months
        valid_date = Utils.date_checker(date_to_check=date, months=months)
        return date, valid_date
    
    def title_extraction(browser, article):
        try:
            title = WebDriverWait(browser, 20).until(
                EC.presence_of_element_located((By.XPATH, loc.article_title_xpath))
            )
            title = article.find_element(By.XPATH, loc.article_title_xpath).text
        except:
            title = 'Title not found'
        return title
    
    def description_extraction(article):
        try:
            description = article.find_element(By.XPATH, loc.article_description_xpath).text
        except:
            description = 'Description not found'
        return description

    def LOCAL_save_to_excel(results):
        wb = openpyxl.Workbook()
        ws = wb.active

        headers = ["Title", "Date", "Description", "Picture Filename", "Count of Search Phrases", "Monetary Amount"]
        ws.append(headers)

        # Write the data rows
        for result in results:
            # print(result)
            row = [
                result["title"],
                result["date"],
                result["description"],
                result["picture_filename"],
                result["count_search_phrases"],
                result["monetary_amount"]
            ]
            ws.append(row)
        wb.save('./output/results.xlsx')

    def save_to_excel(results):
        wb = openpyxl.Workbook()
        ws = wb.active

        headers = [
            "Title", 
            "Date", 
            "Description", 
            "Picture Filename", 
            "Count of Search Phrases", 
            "Monetary Amount"
        ]
        ws.append(headers)

        # Write the data rows
        for result in results:
            row = [
                result["title"],
                result["date"],
                result["description"],
                result["picture_filename"],
                result["count_search_phrases"],
                result["monetary_amount"]
            ]
            ws.append(row)
        
        excel_path = './output/results.xlsx'
        wb.save(excel_path)
        print(f"Excel file saved: {excel_path}")
        
        # Store the Excel file as an asset in Control Room
        asset_name = "results.xlsx"  # You can customize the asset name as needed
        storage.set_file(asset_name, excel_path)
        print(f"Excel file stored as asset: {asset_name}")
