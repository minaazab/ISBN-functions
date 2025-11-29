from openpyxl import load_workbook
from selenium import webdriver
import time
from bs4 import BeautifulSoup

# chooses which URL to use based on user choice
def search(isbn, website):
    if(website == "amazon"):
        base_url = "https://www.amazon.com/s?k="
    if(website == "barnes"):
        base_url = "http://barnesandnoble.com/s/"
    if(website == "rutgers"):
        base_url = "https://www.rutgersuniversitypress.org/search-results-grid/?keyword="
    search_url = f"{base_url}{isbn}"

    return search_url

# chooses what to parse from the HTML based on user chosen website
def parsed(soup, website):
    if(website == "amazon"):
        text = soup.find('h2', class_ ="a-size-base a-spacing-small a-spacing-top-small a-text-normal")
    if(website == "barnes"):
        text = soup.find('td')
    if(website == "rutgers"):
        text = soup.find("h2", class_ ="supapress-results-count")
    return text


# uses selenium to open the website, refresh to get past any bot detection, and edit the excel sheet based on 
def look_up(file_path, website):
    website = website.lower()
    y = 1
    wb = load_workbook(file_path)
    sheet = wb['Sheet1']     
    while sheet['A'+ str(y)].value is not None:
        cell = sheet['A'+ str(y)].value
        driver = webdriver.Chrome()
        driver.get(search(int(cell), website))
        time.sleep(3)
        driver.refresh()
        time.sleep(1)
        soup = BeautifulSoup(driver.page_source, "html.parser")
        text = parsed(soup, website)
        print(text)
        if(text == None):
            sheet['B' + str(y)] = "No"
        else:
            sheet['B' + str(y)] = "Yes"

        wb.save("book.xlsx")
        y += 1

        driver.quit()
