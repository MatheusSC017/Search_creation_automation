from selenium import webdriver
from selenium.webdriver.common.by import By
from openpyxl import load_workbook, Workbook
import time


class RpaFerendum:
    def __init__(self, researches, output_file):
        self._researches = researches
        self._output_file = output_file
        self._researches_list = list()

    @property
    def researches(self):
        return self._researches

    @researches.setter
    def researches(self, researches):
        self._researches = researches

    @property
    def output_file(self):
        return self._output_file

    @output_file.setter
    def output_file(self, output_file):
        self._output_file = output_file

    def run(self):
        browser = webdriver.Firefox()

        data = self._get_data()

        self._researches_list = list()
        for row in data:
            browser.get('https://ferendum.com/pt/')
            self._fill_form(browser, row)
            link_search, link_admin = self._send_form(browser)
            self._researches_list.append((link_search, link_admin))
            time.sleep(1)

    def _get_data(self):
        wb = load_workbook(filename=self._researches)
        sheet = wb['Pesquisas']
        data = list()
        for row in sheet:
            row_values = [cell.value for cell in row]
            row_values[3] = [v.strip() for v in row_values[3].split(',')]
            row_values[4:] = [0 if v == 'Sim' else 1 for v in row_values[4:]]
            data.append(row_values)
        return data[1:]

    def get_researches_list(self):
        return self._researches_list

    def save(self):
        wb = Workbook()
        ws1 = wb.active
        ws1.title = 'Links das pesquisas'

        for row, search in enumerate(self._researches_list):
            ws1.cell(row=row + 1, column=1, value=search[0])
            ws1.cell(row=row + 1, column=2, value=search[1])

        wb.save(filename=self._output_file)

    @staticmethod
    def _fill_form(driver, data):
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

    @staticmethod
    def _send_form(driver):
        driver.find_element(By.CLASS_NAME, 'btn-primary').click()

        time.sleep(1)

        driver.find_element(By.CLASS_NAME, 'btn-primary').click()

        link_search = driver.find_element(By.ID, 'textoACopiar').get_property('href')
        link_admin = driver\
            .find_element(By.CSS_SELECTOR, 'div.card:nth-child(1) > div:nth-child(1) > a:nth-child(22)')\
            .get_property('href')

        return link_search, link_admin


if __name__ == '__main__':
    rpa = RpaFerendum(researches='dados.xlsx', output_file='links.xlsx')
    rpa.run()
    rpa.save()
