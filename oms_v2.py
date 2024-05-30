import mechanicalsoup
import os
import getpass


def login():
    URL = "https://www.bestellsystem.at/Members/Login.aspx"
    if os.path.isfile('./login.txt'):
        with open('login.txt', 'r') as file:
            lines = file.readlines()
            LOGIN = lines[0].strip()
            PASSWORD = lines[1].strip()

    else:
        LOGIN = input("Please enter your username:  ")
        PASSWORD = getpass.getpass('Please enter your password:  ')


    browser = mechanicalsoup.StatefulBrowser(soup_config={'features': 'html5lib'})
    browser.open(URL)
    browser.select_form('form[action="./Login.aspx"]')
    browser["textUsername"] = LOGIN
    browser["textPassword"] = PASSWORD
    response = browser.submit_selected()
    return browser


def find_projects():
    browser.open('https://www.bestellsystem.at/Members/Reports.aspx')
    page = browser.get_current_page()
    dropdown = page.find('select', {'id': 'ctl00_contentMain_dropProjects'})
    options = dropdown.find_all('option')

    items = []
    for option in options:
        item_text = option.get_text(strip=True)
        item_value = option['value']
        items.append((item_text, item_value))
    return items







for item_text, item_value in projects[1:]:
    browser.open('https://www.bestellsystem.at/Members/Reports.aspx')

    page = browser.get_current_page()
    beginn_datum_input = page.find('input', {'id': 'ctl00_contentMain_textBeginnDatum'})
    beginn_datum_input['value'] = '01.01.2024'

    browser.select_form()
    browser["ctl00$contentMain$dropProjects"] = item_value
    browser.submit_selected(btnName='ctl00$contentMain$btnSubmit')
    soup = browser.get_current_page()


table = soup.find('table', {'id': 'ctl00_contentMain_GridView1'})
rows = table.find_all('tr')

for row in rows:
    infos = [info.text.strip() for info in row if info.text.strip()]
    ordering_date = infos[8]
    company = infos[7]
    product = infos[3]
    catalog_number = infos[4]
    quantity = infos[2]
    price = infos[6].split(' EUR')[0].replace(',', '.')
    # orderer = infos[-2]
    # status = infos[10]
    #project = infos[1].split('-')[-1][0:31]
    print(ordering_date)













browser.open('https://www.bestellsystem.at/Members/Reports.aspx')
browser.select_form()
browser["ctl00$contentMain$dropProjects"] = "4507"
browser.submit_selected(btnName='ctl00$contentMain$btnSubmit')
soup = browser.get_current_page()

https://www.bestellsystem.at/Members/Reports.aspx
https://www.bestellsystem.at/Members/Reports/ReportOrders.aspx


browser.open('https://www.bestellsystem.at/Members/Reports.aspx')
page = browser.get_current_page()
dropdown_items = page.select('.dropdown-menu .dropdown-item')

