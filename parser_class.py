import zipfile
import os
import shutil
from bs4 import BeautifulSoup
import codecs
from datetime import datetime


class Parser:
    """
    :param file_list: Список файлов которые могут находиться в zip-архиве
    :type file_list: dict
    """
    file_list = {
        0: 'Общий отчет.htm',
        1: 'Сердечно-сосудистая система.htm',
        2: 'Функция ЖКТ.htm',
        3: 'Состояние печени.htm',
        4: 'Функция толстого кишечника.htm',
        5: 'Функция желчного пузыря.htm',
        6: 'Функция поджелудочной железы.htm',
        7: 'Функция почек.htm',
        8: 'Функция легких.htm',
        9: 'Состояние черепных нервов.htm',
        10: 'Состояние костей.htm',
        11: 'Минеральная плотность костной ткани.htm',
        12: 'Ревматоидные костные заболевания.htm',
        13: 'Индекс роста костей.htm',
        14: 'Сахар в крови.htm',
        15: 'Микроэлементы.htm',
        16: 'Витамины (Нутритивный статус).htm',
        17: 'Аминокислоты.htm',
        18: 'Коферменты.htm',
        19: 'Жирные кислоты.htm',
        20: 'Эндокринная система.htm',
        21: 'Иммунная система.htm',
        22: 'Щитовидная железа.htm',
        23: 'Токсины.htm',
        24: 'Тяжёлые металлы.htm',
        25: 'Основные физические качества.htm',
        26: 'Аллергия.htm',
        27: 'Ожирение.htm',
        28: 'Состояние кожного покрова.htm',
        29: 'Глаза.htm',
        30: 'Коллаген.htm',
        31: 'Пульс сердца и кровоснабжение мозга.htm',
        32: 'Липиды крови.htm',
        33: 'Предстательная железа.htm',
        34: 'Мужская половая функция.htm',
        35: 'Сперма.htm',
        36: 'Гинекология.htm',
        37: 'Молочные железы.htm',
        38: 'Менструальный цикл.htm',
        99: 'Состав тела.htm',
    }

    def __init__(self, path_file):
        """
        :param path_file: путь к файлу архива zip
        :param unpack_dir: директория разархивирования
        :param home_dir: домашняя директория
        :param dict_pars: чистые данные из файлов
        :type dict_pars: dict
        """
        self.path_file = path_file
        self.unpack_dir = path_file.split(".")[0]
        self.home_dir = os.getcwd()
        self.dict_pars = {
            'Date': None
        }

    def unpack(self):
        """Разоархивирование"""

        myfile_zip = zipfile.ZipFile(self.path_file, 'r')
        myfile_zip.extractall(self.unpack_dir)

    def delete(self):
        """Удаляет разархивированую папку."""

        if os.path.exists(self.unpack_dir):
            shutil.rmtree(self.unpack_dir, ignore_errors=False, onerror=None)
            print(f"Папка {self.unpack_dir} удалена")
        else:
            print("Нету разархивированного архива")

    def rename(self):
        """Перекодирование файлов в архиве и
         переименования согласно отчетам в биоанализаторе."""

        try:
            for name in os.listdir(self.unpack_dir):
                if name.split(".")[1] == 'htm':
                    unicode_name = name.encode('cp437').decode('cp866')
                    name_value = unicode_name.split('-', maxsplit=1)[1]
                    for key, value in self.file_list.items():
                        if value == name_value:
                            new_name = str(key) + '-' + value
                            break
                    os.chdir(self.unpack_dir)
                    os.rename(name, new_name)
                    os.chdir('..')
            os.chdir(self.home_dir)
        except Exception:
            print("Переименование файлов не возможно")

    def add_date(self):
        """Дабавление даты и времени в dict_pars."""

        soup = self.get_soup('0-Общий отчет.htm')
        data = soup.find_all('table')[1].find_all('td')[6]
        update = datetime.strptime(data.text[20:], '%d.%m.%Y %H:%M')
        self.dict_pars.update({"Date": update})

    def get_soup(self, file_name):
        """Возвращает сырые данные из файла."""

        file_pars = self.unpack_dir + '/' + file_name
        with codecs.open(file_pars, 'r', encoding='windows-1251') as f:
            soup = BeautifulSoup(f, 'html.parser')
        return soup

    def result(self, file_name):
        """Извлекает данные из файла и формирует словарь dict_pars."""

        file_dict = {}
        num_td = 0
        try:
            num_file = int(file_name.split('-')[0])
            data = self.get_soup(file_name).find('table', class_="table").find_all('td', align="middle")
            for td in data:
                num_td += 1
                if num_td == 1:
                    key_name = td.text
                    if key_name == 'Сывороточный глобулин (A/G)':
                        key_name = 'Сывороточный глобулин (AG)'
                elif num_td == 3:
                    value = float(td.text.replace(',', '.'))
                    file_dict.update({key_name: value})
                elif num_td == 4:
                    num_td = 0
            self.update_dict_pars(num_file, file_dict)
        except Exception:
            print(f"Из файла {file_name} невозможно извлечь данные")

    def update_dict_pars(self, key, file_dict):
        """Обновления dict_pars"""
        self.dict_pars.update({key: file_dict})

    def run(self):
        """Запуск функций в хронологическом порядке."""
        self.unpack()
        self.rename()
        self.add_date()
        for num, name in self.file_list.items():
            file = str(num) + '-' + name
            self.result(file)
        self.delete()
