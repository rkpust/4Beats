import asyncio
from playwright.async_api import Playwright, async_playwright, expect
from bs4 import BeautifulSoup
import pandas as pd
import time
import datetime
from openpyxl import load_workbook


def day_name():
    # Get today's date
    today = datetime.date.today()

    # Get the name of the day
    day_name = today.strftime('%A')

    print(day_name)

    #Return the name of the day
    return day_name


def read_my_excel_file():
    df = pd.read_excel(r"E:\DjangoProject\PlayWright\4Beats\Excel.xlsx", sheet_name=day_name())
    # df_rows = df.index.stop
    # print(df)
    # print(df_rows)

    return df

def read_search_value():
    df = read_my_excel_file()
    # print(df)
    # print(df['Search'].tolist())

    return df['Search'].tolist()


def scrape_data(page_content):
    soup = BeautifulSoup(page_content, 'html.parser')
    # print(soup.prettify())
    all_search = soup.find_all('div', class_='wM6W7d')

    # print(all_search)
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

    # print(max_len_list, min_len_list)
    # print(df['Longest Option'])
    # print(df['Shortest Option'])
    print(df)


    # Assuming 'df' is your DataFrame
    df = df.drop(columns=['Longest Option', 'Shortest Option'])

    # Alternatively, you can use the `iloc` method to drop the last two columns:
    # df = df.iloc[:, :-2]

    # Check the DataFrame to confirm the columns are dropped
    # print(df)

    # Assuming 'df' is your existing DataFrame
    df['Longest Option'] = max_len_list
    df['Shortest Option'] = min_len_list

    # Check the updated DataFrame
    print(df)


    # Assuming 'df' is your updated DataFrame
    file_path = r"E:\DjangoProject\PlayWright\4Beats\Excel.xlsx"

    # Open the Excel file and write to the existing sheet
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        # Write the updated DataFrame to the existing sheet
        df.to_excel(writer, sheet_name=day_name(), index=False)


async def run(playwright: Playwright) -> None:
    max_len_list, min_len_list = [], []
    search_list = read_search_value()
 
    browser = await playwright.chromium.launch(headless=False)
    context = await browser.new_context()
    page = await context.new_page()

    for search in search_list:
        print(search)
        await page.goto("https://www.google.com/")
        time.sleep(1)
        await page.get_by_label("সার্চ করুন", exact=True).click()
        time.sleep(1)
        await page.get_by_label("সার্চ করুন", exact=True).fill(search)
        time.sleep(2)


        page_content = await page.content()
        max_len, min_len = scrape_data(page_content)
        # print(f'1. {max_len}\n 2. {min_len}')
        

        max_len_list.append(max_len)
        min_len_list.append(min_len)
    

    insert_searched_data(max_len_list, min_len_list)

    

    # ---------------------
    await context.close()
    await browser.close()


async def main() -> None:
    async with async_playwright() as playwright:
        await run(playwright)


asyncio.run(main())