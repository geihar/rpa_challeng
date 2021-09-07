import time
import os

from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from RPA.Tables import Tables


browser = Selenium()
excel = Files()
tables = Tables()

agency = 'Department of Commerce'
path = os.path.dirname(os.path.abspath(__file__))
source_pdf = os.path.join(path, 'output/')
source_excel = os.path.join(path, 'output/collected_data.xlsx')


def get_data_from_dive_in():
    browser.set_download_directory(source_pdf)
    browser.open_available_browser("https://itdashboard.gov/")
    browser.click_link('xpath://*[@id="node-23"]/div/div/div/div/div/div/div/a')
    browser.wait_until_element_is_visible('xpath://*[@id="agency-tiles-widget"]/div')
    items_text = browser.get_text('xpath://*[@id="agency-tiles-widget"]/div[1]')
    clean_text = items_text.replace("Total FY2021 Spending:\n", "").split('\n')
    columns = ['Name', 'Total FY2021 Spending:']
    return [dict(zip(columns, clean_text[i:i + 2])) for i in range(0, len(clean_text), 3)]


def save_excel(data: list):
    excel.create_workbook(source_excel)
    excel.rename_worksheet(src_name='Sheet', dst_name='Agencies')
    excel.set_worksheet_value(1, 1, 'Name')
    excel.set_worksheet_value(1, 2, 'Total FY2021 Spending:')
    table = tables.create_table(data=data)
    excel.append_rows_to_worksheet(content=table)
    excel.save_workbook(source_excel)


def save_individual_investments(data: list):
    excel.open_workbook(source_excel)
    excel.create_worksheet(name=agency)
    table = tables.create_table(data=data)
    excel.append_rows_to_worksheet(content=table)
    excel.save_workbook(source_excel)


def get_individual_investments_data():
    browser.click_element(f'partial link:{agency}')
    browser.wait_until_element_is_visible('xpath://*[@id="investments-table-object_length"]/label/select', timeout=15)
    browser.mouse_down('xpath://*[@id="investments-table-object_length"]/label/select')
    browser.page_should_contain_element('xpath://*[@id="investments-table-object_length"]/label/select/option[4]')
    browser.click_element('xpath://*[@id="investments-table-object_length"]/label/select/option[4]')
    browser.wait_until_page_does_not_contain_element('xpath://*[@id="investments-table-object_paginate"]/span/a[2]',
                                                     timeout=30)
    rows_count = browser.get_element_count('xpath://*[@id="investments-table-object"]/tbody/tr')
    cols_count = browser.get_element_count('xpath://*[@id="investments-table-object"]/tbody/tr[1]/td')

    table_data = []
    cols = []
    for i in range(1, cols_count + 1):
        th = browser.get_text(
            f'xpath://*[@id="investments-table-object_wrapper"]/div[3]/div[1]/div/table/thead/tr[2]/th[{i}]')
        cols.append(th)
    table_data.append(cols)

    for i in range(1, rows_count + 1):
        row = []
        for n in range(1, len(cols) + 1):
            data = browser.get_text(f'xpath://*[@id="investments-table-object"]/tbody/tr[{i}]/td[{n}]')
            row.append(data)
        table_data.append(row)

    return [dict(zip(cols, i)) for i in table_data]


def save_files():
    link_count = browser.get_element_count('xpath://*[@id="investments-table-object"]/tbody/tr/td/a')
    link_list = []
    for i in range(1, link_count + 1):
        link = browser.get_element_attribute(f'xpath:// *[@id="investments-table-object"]/tbody/tr[{i}]/td[1]/a',
                                             'href')
        link_list.append(link)
    for i in link_list:
        browser.go_to(i)
        browser.wait_until_element_is_visible('xpath://*[@id="business-case-pdf"]/a', timeout=30)
        browser.click_element('xpath://*[@id="business-case-pdf"]/a')
        time.sleep(5)


def main():
    try:
        clean_text = get_data_from_dive_in()
        save_excel(clean_text)
        ii_data = get_individual_investments_data()
        save_individual_investments(data=ii_data)
        save_files()

    finally:
        browser.close_all_browsers()


if __name__ == '__main__':
    main()
