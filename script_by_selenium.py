import time
import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import pandas as pd


# Update file path
file_path = r"E:\DjangoProject\PlayWright\4Beats\Excel.xlsx"

def day_name():
    # Get today's date
    today = datetime.date.today()

    # Get the name of the day
    day_name = today.strftime('%A')

    print(day_name)

    # Return the name of the day
    return day_name


def read_my_excel_file():
    df = pd.read_excel(file_path, sheet_name=day_name())
    return df

def read_search_value():
    df = read_my_excel_file()
    return df['Search'].tolist()


def scrape_data(page_content):
    soup = BeautifulSoup(page_content, 'html.parser')
    all_search = soup.find_all('div', class_='wM6W7d')
    print(all_search)

    all_search_list = []

    for div in all_search:
        span_text = div.find('span').get_text()  # Find the <span> inside the <div> and get the text
        all_search_list.append(span_text)

    print(all_search_list)

    # Use list comprehension to remove empty strings
    filtered_data = [item for item in all_search_list if item != '']

    # Print the filtered list
    print(filtered_data)

    # Use list comprehension with max to find the element with the maximum length
    max_length_element = max(filtered_data, key=len)
    min_length_element = min(filtered_data, key=len)

    # Print the element with the maximum length
    print(max_length_element)
    print(min_length_element)

    return max_length_element, min_length_element


def insert_searched_data(max_len_list, min_len_list):
    df = read_my_excel_file()

    # Drop existing columns
    df = df.drop(columns=['Longest Option', 'Shortest Option'])

    # Insert new columns
    df['Longest Option'] = max_len_list
    df['Shortest Option'] = min_len_list

    # Open the Excel file and write to the existing sheet
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name=day_name(), index=False)


def run():
    max_len_list, min_len_list = [], []
    search_list = read_search_value()

    # Set up Chrome options
    options = Options()
    options.headless = False  # Set to True if you want to run in headless mode

    # Set up WebDriver
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    for search in search_list:
        print(search)

        # Navigate to Google
        driver.get("https://www.google.com/")
        time.sleep(1)

        # Find the search input box and click it (equivalent to playwright's "get_by_label")
        search_box = driver.find_element(By.NAME, "q")
        search_box.click()
        time.sleep(1)

        # Fill in the search query
        search_box.send_keys(search)
        time.sleep(2)

        # Submit the search by hitting the Enter key (equivalent to playwright's "fill")
        # search_box.send_keys(Keys.RETURN)
        time.sleep(2)

        # Get page content
        page_content = driver.page_source
        # print(page_content)
        
        max_len, min_len = scrape_data(page_content)

        max_len_list.append(max_len)
        min_len_list.append(min_len)

    # Insert the extracted data into the Excel file
    insert_searched_data(max_len_list, min_len_list)

    # Close the WebDriver
    driver.quit()


if __name__ == "__main__":
    run()
