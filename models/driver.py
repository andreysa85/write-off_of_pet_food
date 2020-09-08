""" импортируем библиотеку пандас для работы с ексеселм
и библиотеку конфигпарсер для работы с ини файлом """
import pandas as pd
import configparser


class ReadReference(object):
    """ Чтение справочника вет.препаратов"""
    def __init__(self):
        # Создаем обьект класса конфиг.файла
        self.__conf = Config()
        # читаем конфиг.файл
        self.__path = self.__conf.read_config()

    def read_sp(self):
        # читаем файл с вет препаратами, подаставляем путь с конфиг файла self.__path[0]
        reference = pd.read_excel(self.__path[0])
        # преобразовуем колонку номер с числового в строку
        reference['nomer'] = reference['nomer'].apply(str)
        # возвращаем датафрейм
        return reference


class ReadXLS(object):
    """ Класс для чтения и разбора файлов полученыз с Юниформа, в контруктор получем путь к файлу и
    ферму с которого файл взят """
    def __init__(self, path: str, farm_name: str):
        self.__path = path
        self.__farm_name = farm_name
        # Создаем екзепляр класа конфиругационного файла
        self.__conf_a = Config()
        # читем конфиг. данные
        self.__conf = self.__conf_a.read_config()

    def read_file(self):
        """ Функция для чтения полученних файлов  """
        # создаем обьект класса справочника
        reference = ReadReference()
        # читаем информацию из датафрейма полученную от функции read_sp
        reference = reference.read_sp()
        # преобразовуем колонку номер с числового в строку
        reference['nomer'] = reference['nomer'].apply(str)
        # читаем файл полученый в конструктор, за хедер берем 3 строку
        file = pd.read_excel(self.__path, header=3)
        # отбираем колонки дата, вет.препарат, количество
        df_sort = file[['Дата', self.__farm_name, 'Qty']]
        # преобразовуем колонку дата к дататайму
        df_sort['Дата'] = pd.to_datetime(df_sort['Дата'].astype(str), errors='coerce')
        # в получившемся датафрейме меняем номер партии на название препарата из справочного файла
        df_sort[self.__farm_name] = df_sort[self.__farm_name].replace(reference.set_index('nomer')['name'].dropna())
        # создаем годовой диапазон дат self.__conf[3] и end=self.__conf[4] берем из ини файла
        date = pd.date_range(start=self.__conf[3], end=self.__conf[4])
        # счетчик для подсчета количества обработанных файлов
        count = 0
        # счетчик для прогрессбара
        progress = 0
        # цикл для поиска дат в файле и прохождение по ним
        for i in date:
            progress += 10
            # создаем датафрейм tmp при совпадение даты
            tmp = df_sort[df_sort['Дата'] == str(i)]
            # удаляем из датафрейма строку с "Перевирка норми"
            tmp.drop(tmp[tmp[self.__farm_name] == "100"].index, inplace=True)
            # если tmp пустой (дата не обнаруженна)
            if tmp.empty:
                # ничего не делаем
                pass
            # если tmp не пустой, даты есть
            else:
                # сплитуем дату для дальнейшего использование при сохранении файла
                name = str(i).split(' ')
                # групируем датафрейм по дате и вет.препарату и суммируем количество
                df = tmp.groupby(['Дата', self.__farm_name], sort=True).sum()[['Qty']].reset_index()
                # из датафрейма отбираем колонку вет.препарат
                df_name = df[self.__farm_name]
                # из датафрейма отбираем колонку количество
                df_qty = df['Qty']
                # определяем принадлежность файла к ферме маковище по 'Реєестраці'
                if self.__farm_name == 'Реєестраці':
                    # создаем обьект для сохранение методом вызова класса SaveFile и передаем ему
                    # путь для сохранения взятый из ини файла, дату, фрейм вет.преп, фрейм кол-ва
                    save_file = SaveFile(self.__conf[1], name, df_name, df_qty)
                    # визиваем функциию сохранения
                    save_file.save()
                    # меняем счетчик для кол-ва файлов
                    count += 1
                    # меняем счетчик для прогрессбара
                    #progress += 100
                # определяем принадлежность файла к ферме маковище по 'Реєестраці'
                elif self.__farm_name == 'Реєестраційни':
                    # создаем обьект для сохранение методом вызова класса SaveFile и передаем ему
                    # путь для сохранения взятый из ини файла, дату, фрейм вет.преп, фрейм кол-ва
                    save_file = SaveFile(self.__conf[2], name, df_name, df_qty)
                    # визиваем функциию сохранения
                    save_file.save()
                    # меняем счетчик для кол-ва файлов
                    count += 1
                    # меняем счетчик для прогрессбара
                    #progress += 100
        # фозвращаем счетчики
        return count, progress


class SaveFile(object):
    """ Класс для сохранения файлов, на вход получаем путь куда сохранить, дату, фреймы для сохранения """
    def __init__(self, path: str, name: list, df_name, df_gty):
        self.__path = path
        self.__name = name
        self.__df_name = df_name
        self.__df_gty = df_gty

    def save(self):
        """ Функция для сохранения файлов"""
        # создаем обьект writer для записи self.__path + self.__name[0] - это путь и имя файла
        # их берем из ини файла
        writer = pd.ExcelWriter(self.__path + self.__name[0] + '.xls', engine='xlsxwriter')
        # Сохраняем дафреймы в разных колонках, 0 и 14
        self.__df_name.to_excel(writer, sheet_name='Sheet1', startrow=1, startcol=0, index=False)
        self.__df_gty.to_excel(writer, sheet_name='Sheet1', startrow=1, startcol=14, index=False)
        # сохраняем  файл
        writer.save()
        # Закрываем обьект writer
        writer.close()


class Config(object):
    """
    Класс для чтения конфиг(ини) файла
    """
    def __init__(self):
        # Указываем путь к файлу
        self.__path = 'D:\\my_config.ini'
        # Создаем обьект для работы с ини файлом
        self.__config = configparser.ConfigParser()

    def create_config(self):
        """
        функция для создания конфиг файла
        """
        # Создаем секция для хранения переменных "Settings"
        self.__config.add_section("Settings")
        # добавляем переменны
        self.__config.set("Settings", "path_reference", "d:\\vet_lib.xlsx")
        self.__config.set("Settings", "path_save_mak", "d:\\WWW\\mak\\")
        self.__config.set("Settings", "path_save_kar", "d:\\WWW\\kar\\")
        self.__config.set("Settings", "date_start", "1/1/2019")
        self.__config.set("Settings", "date_stop", "31/12/2019")
        # сохраняем файла по указанному ранее пути
        with open(self.__path, "w") as config_file:
            self.__config.write(config_file)

    def read_config(self):
        """ Функция для чтения конфиг файла """
        # Читаем конфики из файла
        self.__config.read(self.__path)
        # Берем каждую настройку
        path_reference = self.__config.get("Settings", "path_reference")
        path_save_mak = self.__config.get("Settings", "path_save_mak")
        path_save_kar = self.__config.get("Settings", "path_save_kar")
        date_start = self.__config.get("Settings", "date_start")
        date_stop = self.__config.get("Settings", "date_stop")
        # Возвращаем настройки
        return path_reference, path_save_mak, path_save_kar, date_start, date_stop
