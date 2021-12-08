from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from RPA.Tables import Tables

from os.path import abspath
from time import sleep


lib = Selenium()
lib_files = Files()
lib_tables = Tables()


def open_the_website(url):
    lib.open_available_browser(url)

def click_link(locator):
    lib.click_link(locator)

def click_element(locator):
    lib.click_element(locator)

def get_text_elements(locator):
    elements = lib.get_webelements(locator)
    text_list = [lib.get_text(element) for element in elements]
    return text_list

def get_table(cells, col_num):
    extracted_rows = []
    row_cells = []
    for cell in cells:
        if len(row_cells) < col_num:
            row_cells.append(lib.get_text(cell))
        else:
            extracted_rows.append(row_cells)
            row_cells = []
            row_cells.append(lib.get_text(cell))
    return extracted_rows

def get_elements(locator):
    elements = lib.get_webelements(locator)
    return elements

# excel file
def create_workbook(path, wb_name):
    lib_files.create_workbook(path+wb_name)

def create_worksheet(ws_name, data=None, header=None):
    lib_files.create_worksheet(ws_name, data, header)

def rename_sheet(ws_name):
    lib_files.rename_worksheet('Sheet', ws_name)

def append_to_worksheet(data, ws_name=None, header=False, start=None):
    lib_files.append_rows_to_worksheet(data, ws_name, header, start)

def write_data(start_row, start_col, list):
    for num in range(len(list)):
        lib_files.set_cell_value(
            row=start_row+num,
            column=start_col,
            value=str(list[num])
        )

def save_workbook():
    lib_files.save_workbook()

def create_table(data, columns):
    table = lib_tables.create_table(data, columns=columns)
    return table

# await
def wait_for_element(locator, time=20):
    lib.wait_until_element_is_visible(locator, time)

# download part
def download_pdf(agency_url, link_labels, download_locator):
    for label in link_labels:
        lib.go_to(agency_url+f'/{label}')
        wait_for_element(download_locator, 20)
        click_link(download_locator)
        sleep(10)

def get_url():
    return lib.get_location()

def main():
    # base url
    URL = 'https://itdashboard.gov/'
    # path
    path = abspath('.')+'\\'
    # locators
    dive_in_locator = "#home-dive-in"
    dep_name_locator = 'xpath: //div[@id="agency-tiles-container"]//span[contains(@class,"h4")]'
    spending_locator = 'xpath: //div[@id="agency-tiles-container"]//span[contains(@class,"h1")]'
    view_buttons_locator = "//div[@id='agency-tiles-container']//a[text()='view']"
    selector_locator = "//select[contains(@name, 'investments')]//option[text()='All']"
    table_header_locator = "//div[@class='dataTables_scroll']//th[@tabindex]"
    cells_locator = "//div[@class='dataTables_scroll']//tbody//td"
    link_locator = "//div[@class='dataTables_scroll']//tbody//td/a"
    download_locator = "//div[@id='business-case-pdf']/a"
    # names
    wb_name = 'extracted_data.xlsx'
    worksheet_1 = 'Agencies'
    worksheet_2 = 'Individual Investments'
    # agency to scrape
    agency_to_view = 'Department of Justice'

    try:
        open_the_website(URL)
        click_link(dive_in_locator)
        # wait til content loads
        wait_for_element(dep_name_locator)
        dep_list = get_text_elements(dep_name_locator)
        spend_list = get_text_elements(spending_locator)

        # excel
        create_workbook(path, wb_name)
        rename_sheet(worksheet_1)
        write_data(start_row=1, start_col=1, list=dep_list)
        write_data(start_row=1, start_col=2, list=spend_list)

        # click view
        buttons = get_elements(view_buttons_locator)
        click_link(buttons[dep_list.index(agency_to_view)])
        wait_for_element(selector_locator)
        click_element(selector_locator)

        sleep(10)

        # excel
        table_headers = get_text_elements(table_header_locator)
        col_num = len(table_headers)
        cells = get_elements(cells_locator)
        table_rows = get_table(cells, col_num)
        table = create_table(table_rows, table_headers)
        create_worksheet(worksheet_2)
        append_to_worksheet(table, worksheet_2, True)

        # agency url for download
        agency_url = get_url()

        # download PDF files
        lib.set_download_directory(path)
        # get links names
        link_labels = get_text_elements(link_locator)
        download_pdf(agency_url, link_labels, download_locator)

    finally:
        save_workbook()
        lib.close_all_browsers()

if __name__ == "__main__":
    main()