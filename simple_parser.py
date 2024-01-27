from time import sleep

import requests
import xlsxwriter
from bs4 import BeautifulSoup
from requests.models import Response

from decorators import count_time_of_programm

BASE_URL: str = 'http://www.any_site.ru/Reestr/view?id='
START_ID: int = 1
FINISH_ID: int = 1000
XLS_NAME = 'demo.xlsx'
WORKSHEET_NAME = 'Itery_1'
STRUCTURE_OF_COLS: tuple[tuple[str, str]] = (
            ('A1', 'name_org'),
            ('B1', 'town'),
            ('C1', 'inn'),
            ('D1', 'ur_adr'),
            ('E1', 'fact_adr'),
            ('F1', 'director'),
            ('G1', 'site'),
            ('H1', 'telephon'),
            ('I1', 'status')
            )


class ParseSiteData:
    '''Класс для сбора, фильтрации и сохранения данных в xls.'''
    row = 1

    def get_url(self, id: int) -> str:
        '''Собираем урл для запроса.'''
        url = BASE_URL + str(id)
        return url

    def check_conection(self, url: str) -> Response:
        '''
        Проверяем код страницы.
        Если отдается 200, двигаем дальше.
        '''
        check = requests.get(url)
        sleep(0.2)
        if check.status_code == 200:
            return check
        raise ValueError('Данная страница не доступна или отствует')

    def parse_html(self, check: Response):
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
        data = (
            name_org, town, inn, ur_adr, director,
            fact_adr, site, telephon, status
        )
        return data

    def create_excel(self):
        '''Создаем xls и лист в нем.'''
        workbook = xlsxwriter.Workbook(XLS_NAME)
        worksheet = workbook.add_worksheet(WORKSHEET_NAME)

        bold = workbook.add_format({'bold': 5})

        for col_key, col_name in STRUCTURE_OF_COLS:
            worksheet.write(col_key, col_name, bold)

        return worksheet, workbook

    def write_excel(self, *data, worksheet):
        '''Записывает полученные данные.'''
        col = 0

        for index_col, value in enumerate(data):
            worksheet.write_string(self.row, col + index_col, value)
        self.row += 1

    @count_time_of_programm
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
            workbook.close()
        except Exception as error:
            print(f'возникла ошибка: {error}')


if __name__ == '__main__':
    test = ParseSiteData()
    test.main()
