from tkinter import Frame, Toplevel
from abc import ABC, abstractmethod


class MainWindowClass(ABC, Frame):
    """Абстрактный класс основного окна"""
    def __init__(self, parent):
        """Инициируем класс основного окна"""
        Frame.__init__(self, parent)

    @abstractmethod
    def start_program(self):
        """Функция, управляющая началом целевой работы программы"""
        pass


class MainChildWindowClass(ABC, Toplevel):
    """Абстрактный класс всех дочерних окон"""
    def __init__(self):
        """Инициируем класс всех дочерних окон"""
        Toplevel.__init__(self)


class MainSearcherClass(ABC):
    """Абстрактный класс, отвечающий за реализацию поисковой части программы"""
    @abstractmethod
    def connection_to_db(self, *args):
        """Реализация подкючения к базе данных."""
        pass

    @abstractmethod
    def searching_for_match(self, *args):
        """Функция-поисковик, реализующая поиск совпадений значения поля запроса
        с каким-либо значением внутри определенных полей БД"""
        pass


class MainInputClass(ABC):
    """Абстроктный класс, отвечающий за реализацию помещения данных в БД"""
    @abstractmethod
    def correct_input(self, *args):
        """Функция-помощник, вносящая изменения и обновляющая нужную часть БД, в
        зависимости от внесенного запроса"""
        pass


class MainAppenderClass(ABC):
    """Абстрактный класс, отвечающий за реализацию вставки новых значений
    в БД"""
    @abstractmethod
    def connection_to_db(self, *args):
        """Реализация подкючения к базе данных."""
        pass

    @abstractmethod
    def append_to_db(self, *args):
        """Функция-помощник, которая вносит новые значения"""
        pass


class MainReporterClass(ABC):
    """Абстрактный класс, отвечающий за реализацию создания отчетов
    пользователей"""
    @abstractmethod
    def create_done_data_frame(self, *args):
        """Создание датафрейма израсходованных стандартов"""
        pass

    @abstractmethod
    def create_not_done_data_frame(self, *args):
        """Создание датафрейма не израсходованных стандартов"""
        pass

