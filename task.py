from RPA.FileSystem import FileSystem
from RPA.Excel.Files import Files
from RPA.Tables import Tables
from Browser.utils.data_types import SelectAttribute
import Browser

from time import time
from PyPDF2 import PdfFileReader


lib_files = Files()
lib_tables = Tables()
fs_lib = FileSystem()
browser = Browser.Browser()


# work with browser
def open_browser(url, path):
    """open the page with modified options"""
    
    browser.new_browser(downloadsPath=path, headless=False)
    browser.new_context(acceptDownloads=True)
    browser.new_page(url)

def click_element(locator):
    """click the element"""
    browser.click(locator)

def set_option(selector_locator, option):
    """set number of entries in the table"""
    browser.select_options_by(selector_locator, SelectAttribute.label, option)

def get_elements(locator):
    """get and return a list of elements"""
    elements = browser.get_elements(locator)
    return elements

def get_text_elements(locator):
    """get and return a text from the elements"""
    elements = get_elements(locator)
    text_list = [browser.get_text(element) for element in elements]
    return text_list

def get_url():
    """get a URL of the current page"""
    return browser.get_url()

def download_pdf(agency_url, label, locator, path):
    """download a PDF file by label"""
    path_to_file = path + f'{label}.pdf'
    browser.go_to(agency_url + f'/{label}')
    click_element(locator)
    browser.promise_to_wait_for_download(path_to_file)
    return path_to_file

# work with tables
def get_table(cells, col_num):
    """get and return a list of the table rows"""
    extracted_rows = []
    row_cells = []
    for cell in cells:
        if len(row_cells) < col_num:
            row_cells.append(browser.get_text(cell))
        else:
            extracted_rows.append(row_cells)
            row_cells = []
            row_cells.append(browser.get_text(cell))
    extracted_rows.append(row_cells)
    return extracted_rows

def create_table(data, columns):
    """create and return a new table"""
    table = lib_tables.create_table(data, columns=columns)
    return table

# work with Excel
def create_workbook(path, wb_name):
    """create a new workbook"""
    lib_files.create_workbook(path+wb_name)

def create_worksheet(ws_name, data=None, header=None):
    """create a new worksheet"""
    lib_files.create_worksheet(ws_name, data, header)

def rename_sheet(ws_name):
    """rename the worksheet"""
    lib_files.rename_worksheet('Sheet', ws_name)

def append_to_worksheet(data, ws_name=None, header=False, start=None):
    """fill the worksheet with a table"""
    lib_files.append_rows_to_worksheet(data, ws_name, header, start)

def write_data(start_row, start_col, list):
    """write the data by columns to a table"""
    for num in range(len(list)):
        lib_files.set_cell_value(
            row=start_row+num,
            column=start_col,
            value=str(list[num])
        )

def save_workbook():
    """save the workbook"""
    lib_files.save_workbook()

# waiting
def wait_for_element(locator, timeout=20):
    """wait until the element isn't visible"""
    browser.wait_for_elements_state(locator, timeout=timeout)

def wait_full_table(elements_locator):
    """wait for all elements in the table"""
    table_elements = get_elements(elements_locator)
    timer_start = time()
    timer_tick = 0
    try:
        while len(table_elements) <= 10:
            if timer_tick - timer_start >= 45:
                raise LookupError
            table_elements = get_elements(elements_locator)
            timer_tick = time()
    except LookupError:
        print("!"*10 + "Error loading all table elements. Or number of items <= 10" + "!"*10)

# work with file system
def check_file(
    path_to_file, filename,
    to_find_1, to_find_2,
    search_limit, table_rows
):
    """check if the file was downloaded and compare the data"""

    downloaded = False
    timer_start = time()
    timer_tick = 0
    try:
        while not downloaded:
            if timer_tick - timer_start >= 40:
                raise FileExistsError
            downloaded = fs_lib.does_file_exist(path_to_file)
            timer_tick = time()
        else:
            print("-"*10 + f"{filename} downloaded" + "-"*10)
        if fs_lib.does_file_exist(path_to_file):
            compare_data(path_to_file, to_find_1, to_find_2, search_limit, table_rows)
    except FileExistsError:
        print("!"*10 + f"Something wrong with downloading {filename}" + "!"*10)

# work with PDF
def get_pdf_text(pdf_file):
    """get and return a text from the first page of the pdf"""
    with open(pdf_file, 'rb') as pdf:
        pdf_reader = PdfFileReader(pdf)
        page1 = pdf_reader.getPage(0)
        return page1.extractText()

# work with text
def get_string(text, start_locator, end_locator):
    """find the needed string in the extracted text and return it"""
    start_index = text.index(start_locator) + len(start_locator)
    end_index = text.index(end_locator)
    return text[start_index:end_index].strip()

def compare_data(path_to_pdf, to_find_1, to_find_2, search_limit, table_rows):
    """compare the file data with the table rows"""
    text = get_pdf_text(path_to_pdf)
    investment_name = get_string(text, to_find_1, to_find_2)
    uii = get_string(text, to_find_2, search_limit)
    for row in table_rows:
        if row[0] == uii and row[2] == investment_name:
            founded = True
            break
        else:
            founded = False
    if founded:
        print(f"The investment {uii} - {investment_name} found at the {table_rows.index(row) + 1} row.")
    else:
        print(f"{uii} {investment_name} not found.")


def main():
    # base url
    URL = 'https://itdashboard.gov/'
    # path
    output_folder = 'output'
    path = fs_lib.absolute_path('.') + f'\\{output_folder}\\'
    # locators
    dive_in_locator = "//a[@href='#home-dive-in']"
    dep_name_locator = "//div[@id='agency-tiles-container']//span[contains(@class,'h4')]"
    spending_locator = "//div[@id='agency-tiles-container']//span[contains(@class,'h1')]"
    view_buttons_locator = "//div[@id='agency-tiles-container']//a[text()='view']"
    selector_locator = "//*[@id='investments-table-object_length']/label/select"
    table_header_locator = "//div[@class='dataTables_scroll']//th[@tabindex]"
    cells_locator = "//div[@class='dataTables_scroll']//tbody//td"
    link_locator = "//div[@class='dataTables_scroll']//tbody//td/a"
    download_locator = "//div[@id='business-case-pdf']/a"
    elements_locator = "//*[@id='investments-table-object']/tbody/tr"
    # options
    option_to_select = "All"
    # names
    wb_name = "extracted_data.xlsx"
    worksheet_1 = "Agencies"
    worksheet_2 = "Individual Investments"
    # agency to scrape
    agency_to_view = "Department of Justice"
    # to find in PDF
    to_find_1 = '1. Name of this Investment:'
    to_find_2 = '2. Unique Investment Identifier (UII):'
    search_limit = 'Section B:'

    try:
        # go to website and collect text data
        open_browser(URL, path)
        click_element(dive_in_locator)
        dep_list = get_text_elements(dep_name_locator)
        spend_list = get_text_elements(spending_locator)

        # create an Excel workbook, rename the worksheet to "Agencies", and write agencies and expenses there
        create_workbook(path, wb_name)
        rename_sheet(worksheet_1)
        write_data(start_row=1, start_col=1, list=dep_list)
        write_data(start_row=1, start_col=2, list=spend_list)

        # go to the agency page and select all elements in the investments table
        buttons = get_elements(view_buttons_locator)
        click_element(buttons[dep_list.index(agency_to_view)])
        wait_for_element(selector_locator)
        set_option(selector_locator, option_to_select)
        wait_full_table(elements_locator)

        # collect table data, create a new excel worksheet: "Individual Investments", write collected data
        table_headers = get_text_elements(table_header_locator)
        col_num = len(table_headers)
        cells = get_elements(cells_locator)
        table_rows = get_table(cells, col_num)
        table = create_table(table_rows, table_headers)
        create_worksheet(worksheet_2)
        append_to_worksheet(table, worksheet_2, True)

        # get agency URL for download PDF files
        agency_url = get_url()

        # download PDF files and compare data
        link_labels = get_text_elements(link_locator)
        for label in link_labels:
            path_to_file = download_pdf(agency_url, label, download_locator, path)
            check_file(path_to_file, label+'.pdf', to_find_1, to_find_2, search_limit, table_rows)

    finally:
        save_workbook()
        lib_files.close_workbook()
        browser.close_browser()

if __name__ == "__main__":
    main()