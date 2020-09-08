from tkinter import Tk, Label, Radiobutton, Button, PhotoImage, IntVar, CENTER, Menu
import tkinter.ttk as ttk
from tkinter import filedialog as fd
from tkinter import messagebox as mb
import os
from models.driver import ReadXLS


class MainWindow(object):
    """Класс главного окна приложения  """
    def __init__(self):
        # Создаем окно
        self.__win = Tk()
        # Создаем лейбла для фото
        self.__label_photo = Label(self.__win)
        # Создаем радиокнопку выбора фермы маковище
        self.__radio_mak = Radiobutton(self.__win)
        # Создаем радиокнопку выбора фермы карашин
        self.__radio_kar = Radiobutton(self.__win)
        # Создаем кнопку открыть файла
        self.__btn_open = Button(self.__win)
        # Создаем кнопку закрыть программу
        self.__btn_close = Button(self.__win)
        # Создаем обьект PhotoImage передаем имя картинки
        self.__img = PhotoImage(file='vet.png')
        # Создаем обьект прогрессбара
        self.__pb = ttk.Progressbar(self.__win)
        # переменная для определение радиокнопок(ферм)
        self.__var = IntVar()
        # Дефолтное значение ферма маковище
        self.__var.set(0)
        # Создаем главно меню
        self.__main_menu = Menu()

    def config(self):
        """ Функция для настройки виджетов """
        self.__win.title('Списання вет.ліків ')
        self.__win.geometry('350x300')
        self.__win.config(menu=self.__main_menu)
        self.__radio_mak.config(text="Ферма Маковище", variable=self.__var, value=0)
        self.__radio_kar.config(text="Ферма Карашин", variable=self.__var, value=1)
        self.__btn_open.config(text='Відкрити та обробити файл', command=self.btn_open)
        self.__pb = ttk.Progressbar(mode="determinate")
        self.__btn_close.config(text='Вихід', command=self.__win.destroy)
        self.__label_photo.config(image=self.__img)
        self.__main_menu.add_command(label='Открыть файл', command=self.btn_open)
        self.__main_menu.add_command(label='Налаштування', command=self.settings)

    def layout(self):
        """ Функция расположения виджетов """
        self.__label_photo.pack()
        self.__radio_mak.pack(anchor=CENTER)
        self.__radio_kar.pack(anchor=CENTER)
        self.__btn_open.pack(padx=5, pady=5, anchor=CENTER)
        self.__pb.pack()
        self.__btn_close.pack(padx=5, pady=5, anchor=CENTER)

    @staticmethod
    def settings():
        win = SettingsWindow()
        win.show()

    def btn_open(self):
        """ Функция для обработки нажатия на кнопку откыть файл  """
        # Создаем диалоговое окно выбора файла
        filename = fd.askopenfilename()
        # Берем абсолютный путь к выбранному файлу
        path = os.path.abspath(filename)
        # если выбрана верма 0 - маковище
        if self.__var.get() == 0:
            # Создаем обьект класса ReadXLS, передав ему путь к файлу и идентификатор ферму
            f = ReadXLS(path, 'Реєестраці')
            try:
                # Вызываем функцию чтения файла
                count = f.read_file()
                # меняем прогрессбар
                self.__pb['value'] = count[1]
                # выводим информационное сообщение
                mb.showinfo('OK', f'Вивантаження завершено, \n було обробленно {count[0]} днів')
                # Обрабатываем ошибку
            except LookupError:
                mb.showerror("Ошибка", "Файл не з тієї ферми")
        # если выбрана верма 1 - карашин
        elif self.__var.get() == 1:
            # Создаем обьект класса ReadXLS, передав ему путь к файлу и идентификатор ферму
            f = ReadXLS(path, 'Реєестраційни')
            try:
                # Вызываем функцию чтения файла
                count = f.read_file()
                # меняем прогрессбар
                self.__pb['value'] = count[1]
                # выводим информационное сообщение
                mb.showinfo('OK', f'Вивантаження завершено, \n було обробленно {count[0]} днів')
                # Обрабатываем ошибку
            except LookupError:
                mb.showerror("Ошибка", "Файл не з тієї ферми")

    def show(self):
        """ Функция для запуска основного цикла приложения  """
        self.config()
        self.layout()
        self.__win.mainloop()


class SettingsWindow(object):

    def __init__(self):
        self.__win_settings = Tk()

    def config(self):
        self.__win_settings.title('Налаштування')
        self.__win_settings.geometry('150x150')

    def show(self):
        self.__win_settings.config()
        self.__win_settings.mainloop()
