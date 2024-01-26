from collections.abc import Iterable
from time import sleep

import requests
import xlsxwriter
from bs4 import BeautifulSoup

BASE_URL: str = 'http://www.any_site.ru/Reestr/view?id='
START_ID: int = 1
FINISH_ID: int = 1000
XLS_NAME = 'demo.xlsx'
WORKSHEET_NAME = 'Itery_1'


class ParseSiteData:
    '''Класс для сбора, фильтрации и сохранения данных в xls.'''
    row = 1

    def get_url(self, id: int) -> str:
        '''Собираем урл для запроса.'''
        url = BASE_URL + str(id)
        return url

    def check_conection(self, url: str) -> Iterable:
        '''
        Проверяем код страницы.
        Если отдается 200, двигаем дальше.
        '''
        check = requests.get(url)
        sleep(0.2)
        if check.status_code == 200:
            return check
        raise ValueError('Данная страница не доступна или отствует')

    def parse_html(self, check: Iterable):
        '''Получаем разметку всей страницы.'''
        text = check.text
        return text

    def bs_page(self, text: str) -> str:
        '''Находим, и получаем нужные участки.'''
        soup = BeautifulSoup(text, 'html.parser')
        name_org: str = soup.find('any-tag').get_text()
        town: str = soup.find_all('any-tag', class_="any-class")[3].get_text()
        inn: str = soup.find_all('any-tag', class_="any-class")[4].get_text()
        ur_adr: str = soup.find_all('any-tag', class_="any-class").get_text()
        fact_adr: str = soup.find_all('any-tag', class_="any-class").get_text()
        director: str = soup.find_all('any-tag', class_="any-class").get_text()
        site: str = soup.find_all('any-tag', class_="any-class").get_text()
        telephon: str = soup.find_all('any-tag', class_="any-class").get_text()
        status: str = soup.find_all('any-tag', class_="any-class").get_text()
        return (
            name_org, town, inn, ur_adr, director,
            fact_adr, site, telephon, status
        )

    def create_excel(self):
        '''Создаем xls и лист в нем.'''
        workbook = xlsxwriter.Workbook(XLS_NAME)
        worksheet = workbook.add_worksheet(WORKSHEET_NAME)

        bold = workbook.add_format({'bold': 1})

        worksheet.write('A1', 'name_org', bold)
        worksheet.write('B1', 'town', bold)
        worksheet.write('C1', 'inn', bold)
        worksheet.write('D1', 'ur_adr', bold)
        worksheet.write('E1', 'fact_adr', bold)
        worksheet.write('F1', 'director', bold)
        worksheet.write('G1', 'site', bold)
        worksheet.write('H1', 'telephon', bold)
        worksheet.write('I1', 'status', bold)
        return worksheet, workbook

    def write_excel(
        self, name_org: str, town: str, inn: str, ur_adr: str,
        director: str, fact_adr: str, site: str,
        telephon: str, status: str, worksheet
    ):
        '''Записывает полученные данные.'''
        col = 0
        worksheet.write_string(self.row, col, name_org)
        worksheet.write_string(self.row, col + 1, town)
        worksheet.write_string(self.row, col + 2, inn)
        worksheet.write_string(self.row, col + 3, ur_adr)
        worksheet.write_string(self.row, col + 4, fact_adr)
        worksheet.write_string(self.row, col + 5, director)
        worksheet.write_string(self.row, col + 6, site)
        worksheet.write_string(self.row, col + 7, telephon)
        worksheet.write_string(self.row, col + 8, status)

    def main(self):
        '''Функция запуска функциональности.'''
        try:
            worksheet, workbook = self.create_excel()
            for id in range(START_ID, FINISH_ID):
                url = self.get_url(id)
                check = self.check_conection(url)
                text = self.parse_html(check)
                data = self.bs_page(text)
                self.write_excel(*data, worksheet)
                self.row += 1
            workbook.close()
        except Exception as error:
            print(f'возникла ошибка: {error}')


if __name__ == '__main__':
    test = ParseSiteData()
    test.main()
