class OperatorNameError(Exception):
    """Класс ошибки, связанной с отсутстивем имени оператора"""
    def __init__(self, text):
        self.text = text
        print(self.text)


class ExcessOfTheValue(Exception):
    """Класс ошибки, связанной с превышением доступной навески"""
    def __init__(self, text):
        self.text = text
        print(self.text)


class StringIsEmpty(Exception):
    """Класс ошибки, связанной с пустым полем ввода ID стандарта"""
    def __init__(self, text):
        self.text = text
        print(self.text)


class StandartIsDone(Exception):
    """Класс ошибки, связанной с израсходованием стандарта"""
    def __init__(self, text):
        self.text = text
        print(self.text)


class IdIsEmpty(Exception):
    """Класс ошибки, связанной с отсутствием последнего
     зарегистрированного ID стандарта"""
    def __init__(self, text):
        self.text = text
        print(self.text)


class WaitingTime(Exception):
    """Класс ошибки, связанной со временем ожидания
    для сохранения файла"""
    def __init__(self, text):
        self.text = text
        print(self.text)


class NoProjectName(Exception):
    """Класс ошибки, связанной с отсутствием имени
     проекта, при добавлении в БД"""
    def __init__(self, text):
        self.text = text
        print(self.text)


class NoStandardName(Exception):
    """Класс ошибки, связанной с отсутствием ввдённого
     имени стандарта, при добавлении в БД"""
    def __init__(self, text):
        self.text = text
        print(self.text)


class NoSerialNumber(Exception):
    """Класс ошибки, связанной с отсутствием ввдённого
     серийного номера стандарта, при добавлении в БД"""
    def __init__(self, text):
        self.text = text
        print(self.text)


class NoCurrentValue(Exception):
    """Класс ошибки, связанной с отсутствием ввдённого
     номинального объёма стандарта, при добавлении в БД"""
    def __init__(self, text):
        self.text = text
        print(self.text)


class StrValidationError(Exception):
    """Класс ошибки, связанной с неправильном сохранением
     в файле базы даных строки с идентификационным номером"""
    def __init__(self, text):
        self.text = text
        print(self.text)
