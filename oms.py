import mechanicalsoup


class OMS_item:
    def __init__(self, ordering_date, company, product, catalog_number, quantity, price, orderer, status, project,
                 row=None):
        self.ordering_date = ordering_date
        self.company = company
        self.product = product
        self.catalog_number = catalog_number
        self.quantity = quantity
        self.price = price
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

with open('login.txt', 'r') as file:
    lines = file.readlines()
    LOGIN = lines[0].strip()
    PASSWORD = lines[1].strip()

browser = mechanicalsoup.StatefulBrowser(soup_config={'features': 'html5lib'})
browser.open(URL)
browser.select_form('form[action="./Login.aspx"]')

browser["textUserName"] = LOGIN
browser["textPasswort"] = PASSWORD

response = browser.submit_selected()
browser.open('https://www.bestellsystem.at/Members/CurrentOrders.aspx')
soup = browser.get_current_page()

divs = soup.findAll('div')
u=divs[52].find_all('tr')
u[1].find_all('td')


all_rows = soup.find_all('tr')

all_projects = list()
all_items = list()

for row in all_rows:
    if '<input border="0" id="ctl00_contentMain_GridView1_ctl' in str(row):
        info = row.find_all('td')
        ordering_date = info[10].text
        company = info[9].text
        product = info[8].text
        catalog_number = info[7].text
        quantity = info[6].text
        price = info[12].text.replace(',', '.')
        orderer = info[11].text
        status = info[3].text.strip()
        project = info[5].text[0:31]

        all_projects.append(project)
        if '2024' in ordering_date:
            all_items.append(OMS_item(ordering_date=ordering_date, company=company, product=product,
                                      catalog_number=catalog_number, quantity=quantity, price=price, orderer=orderer,
                                      status=status, project=project))

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
        orderer = ws.range(f'G{row}').value
        status = ws.range(f'H{row}').value
        project = ws.range(f'I{row}').value
        project_dict[sheet_name].known_items.add_item(OMS_item(ordering_date=ordering_date, company=company,
                                                             product=product, catalog_number=catalog_number,
                                                             quantity=quantity, price=price, orderer=orderer,
                                                             status=status, project=project, row=row))


def correct_old_items(sheet_name):
    ws = master_sheets[sheet_name]
    print(sheet_name)
    for thing in all_items:
        if thing.project == sheet_name:
            if thing in project_dict[sheet_name].known_items:
                index = project_dict[sheet_name].known_items.index(thing)
                if thing.status != project_dict[sheet_name].known_items.items[index].status:
                    row = project_dict[sheet_name].known_items.items[index].row
                    ws.range(f'H{row}').value = thing.status



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
            ws.range(f'G{row+last_row}').value = project_dict[sheet_name].new_items[row].orderer
            ws.range(f'H{row+last_row}').value = project_dict[sheet_name].new_items[row].status
            ws.range(f'I{row+last_row}').value = project_dict[sheet_name].new_items[row].project


import xlwings as xw
master_wb = xw.Book(r"Y:\Lab Management\Bestellliste_TU_2024.xlsx")
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
