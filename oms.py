import mechanicalsoup
import os
import getpass


class OMS_item:
    def __init__(self, ordering_date, company, product, catalog_number, quantity, price, price_brutto, orderer, status,
                 project, row=None):
        self.ordering_date = ordering_date
        self.company = company
        self.product = product
        self.catalog_number = catalog_number
        self.quantity = quantity
        self.price = price
        self.price_brutto = price_brutto
        self.orderer = orderer
        self.status = status
        self.project = project
        self.row = row

    def __eq__(self, other):
        return self.ordering_date == other.ordering_date and self.company == other.company and\
               self.product == other.product and\
               self.quantity == other.quantity and self.orderer == other.orderer and\
               self.project == other.project


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
browser.open('https://www.bestellsystem.at/Members/CurrentOrders.aspx')
soup = browser.get_current_page()

divs = soup.findAll('div')
u=divs[7].find_all('tr')



all_projects = list()
all_items = list()

for row in u[1:-1]:
    infos = [info.text.strip() for info in row if info.text.strip()]
    ordering_date = infos[-3]
    company = infos[-4]
    product = infos[-5]
    catalog_number = infos[4]
    quantity = infos[3]
    price = infos[-1].split(' EUR')[0].replace(',', '.')
    orderer = infos[-2]
    status = infos[1]
    project = infos[2][0:31]

    all_projects.append(project)
    if '2024' in ordering_date:
        all_items.append(OMS_item(ordering_date=ordering_date, company=company, product=product,
                                  catalog_number=catalog_number, quantity=quantity, price=price,
                                  price_brutto=float(price)*1.2, orderer=orderer, status=status, project=project))

all_projects = [i[0:31] for i in set(all_projects)]


class OMS_item_list:
    def __init__(self):
        self.items = []

    def add_item(self, item):
        self.items.append(item)

    def __contains__(self, item_of_interest):
        for item in self.items:
            if item == item_of_interest:
                return True
        return False

    def __len__(self):
        return len(self.items)

    def index(self, item_of_interest):
        for index, item in enumerate(self.items):
            if item == item_of_interest:
                return index


class Project:
    def __init__(self, name):
        self.name = name
        self.known_items = OMS_item_list()
        self.new_items = list()
        self.last_row = None


def get_last_row(sheet_name):
    ws = master_sheets[sheet_name]
    last_row = ws.range('A' + str(ws.cells.last_cell.row)).end('up').row
    project_dict[sheet_name].last_row = last_row+1


def get_old_items(sheet_name):
    ws = master_sheets[sheet_name]
    last_row = project_dict[sheet_name].last_row
    for row in range(12, last_row):
        ordering_date = ws.range(f'A{row}').value
        company = ws.range(f'B{row}').value
        product = ws.range(f'C{row}').value
        catalog_number = ws.range(f'D{row}').value
        quantity = ws.range(f'E{row}').value
        price = ws.range(f'F{row}').value
        price_brutto = ws.range(f'G{row}').value
        orderer = ws.range(f'H{row}').value
        status = ws.range(f'I{row}').value
        project = ws.range(f'J{row}').value
        project_dict[sheet_name].known_items.add_item(OMS_item(ordering_date=ordering_date, company=company,
                                                             product=product, catalog_number=catalog_number,
                                                             quantity=quantity, price=price, price_brutto=price_brutto,
                                                             orderer=orderer, status=status, project=project, row=row))


def correct_old_items(sheet_name):
    ws = master_sheets[sheet_name]
    print(sheet_name)
    for thing in all_items:
        if thing.project == sheet_name:
            if thing in project_dict[sheet_name].known_items:
                index = project_dict[sheet_name].known_items.index(thing)
                if thing.status != project_dict[sheet_name].known_items.items[index].status:
                    row = project_dict[sheet_name].known_items.items[index].row
                    ws.range(f'I{row}').value = thing.status



def get_new_items(sheet_name):
    ws = master_sheets[sheet_name]
    for thing in all_items:
        if thing.project == sheet_name:
            if thing not in project_dict[sheet_name].known_items:
                project_dict[sheet_name].new_items.append(thing)


def write_new_items(sheet_name):
    ws = master_sheets[sheet_name]
    last_row = project_dict[sheet_name].last_row
    if last_row < 12:
        last_row = 12
    if project_dict[sheet_name].new_items:
        project_dict[sheet_name].new_items = project_dict[sheet_name].new_items[::-1]
        for row in range(len(project_dict[sheet_name].new_items)):
            ws.range(f'A{row+last_row}').value = project_dict[sheet_name].new_items[row].ordering_date
            ws.range(f'B{row+last_row}').value = project_dict[sheet_name].new_items[row].company
            ws.range(f'C{row+last_row}').value = project_dict[sheet_name].new_items[row].product
            ws.range(f'D{row+last_row}').value = project_dict[sheet_name].new_items[row].catalog_number
            ws.range(f'E{row+last_row}').value = project_dict[sheet_name].new_items[row].quantity
            ws.range(f'F{row+last_row}').value = project_dict[sheet_name].new_items[row].price
            ws.range(f'G{row+last_row}').value = project_dict[sheet_name].new_items[row].price_brutto
            ws.range(f'H{row+last_row}').value = project_dict[sheet_name].new_items[row].orderer
            ws.range(f'I{row+last_row}').value = project_dict[sheet_name].new_items[row].status
            ws.range(f'J{row+last_row}').value = project_dict[sheet_name].new_items[row].project


import xlwings as xw

with xw.App(visible=True) as app:
    path = r"Y:\ARRRGH_Users\Bernhard\1_Projects\OMS\Bestellliste_TU_2024_testing.xlsx"
    # master_wb = xw.Book(r"Y:\Lab Management\Bestellliste_TU_2024.xlsx")
    master_wb = xw.Book(path)

    master_sheets = master_wb.sheets
    master_sheets_names = [s.name for s in master_sheets]


    project_dict = dict()
    for item in all_projects:
        project_dict[item] = Project(item)
        get_last_row(item)
        get_old_items(item)

        correct_old_items(item)
        get_new_items(item)
        write_new_items(item)

    import time
    # time.sleep(100)

    master_wb.save(path)
    master_wb.close()
