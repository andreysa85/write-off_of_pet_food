from gui.gui import MainWindow
from models.driver import Config
if __name__ == '__main__':
    # Создаем екземпляр класса главного окна
    win = MainWindow()
    # Создаем екземпляр класса чтения конфиг файла
    config = Config()
    # Читаем конфиг файл
    config.read_config()
    # Отображаем главное окно приложения
    win.show()
