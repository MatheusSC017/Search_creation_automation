from selenium import webdriver
from selenium.webdriver.common.by import By
from openpyxl import load_workbook, Workbook
import time


def rpa_form():
    browser = webdriver.Firefox()

    data = get_data()

    link_list = list()
    for row in data:
        browser.get('https://ferendum.com/pt/')
        fill_form(browser, row)
        link_search, link_admin = send_form(browser)
        link_list.append((link_search, link_admin))
        time.sleep(1)

    save_links(link_list)


def get_data():
    wb = load_workbook(filename='dados.xlsx')
    sheet = wb['Pesquisas']
    data = list()
    row = 1
    row_data = list()
    for key, cell in sheet._cells.items():
        if key[0] != row:
            data.append(row_data)
            row_data = list()
            row = key[0]

        if key[1] == 4:
            row_data.append([v.strip() for v in cell.value.split(',')])
        elif key[1] >= 5:
            row_data.append(0 if cell.value == 'Sim' else 1)
        else:
            row_data.append(cell.value)

    data.append(row_data)
    return data[1:]


def save_links(link_list, filename='links.xlsx'):
    wb = Workbook()
    ws1 = wb.active
    ws1.title = 'Links das pesquisas'

    for row, links in enumerate(link_list):
        ws1.cell(row=row + 1, column=1, value=links[0])
        ws1.cell(row=row + 1, column=2, value=links[1])

    wb.save(filename=filename)


def fill_form(driver, data):
    driver.find_element(By.NAME, 'titulo').send_keys(data[0])
    driver.find_element(By.NAME, 'descripcion').send_keys(data[1])
    driver.find_element(By.NAME, 'creador').send_keys(data[2])
    for i, op in enumerate(data[3]):
        driver.find_element(By.ID, f'op{i + 1}').send_keys(op)
    driver.find_elements(By.NAME, 'config_anonimo')[data[4]].click()
    driver.find_elements(By.NAME, 'config_priv_pub')[data[5]].click()
    driver.find_elements(By.NAME, 'config_un_solo_voto')[data[6]].click()
    driver.find_elements(By.NAME, 'config_aut_req')[data[7]].click()
    driver.find_element(By.NAME, 'accept_terms_checkbox').click()


def send_form(driver):
    driver.find_element(By.CLASS_NAME, 'btn-primary').click()

    time.sleep(1)

    driver.find_element(By.CLASS_NAME, 'btn-primary').click()

    link_search = driver.find_element(By.ID, 'textoACopiar').get_property('href')
    link_admin = driver\
        .find_element(By.CSS_SELECTOR, 'div.card:nth-child(1) > div:nth-child(1) > a:nth-child(22)')\
        .get_property('href')

    return link_search, link_admin


if __name__ == '__main__':
    rpa_form()
