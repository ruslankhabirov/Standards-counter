from Architecture import MainWindowClass, MainChildWindowClass
from Searcher import SearcherClass, InputClass, connection_to_list_db, CorrectMistakes, write_off_standart, safe_saving
from Exceptions import OperatorNameError, ExcessOfTheValue, StringIsEmpty, StandartIsDone, IdIsEmpty
from Exceptions import WaitingTime, NoProjectName, NoStandardName, NoSerialNumber, NoCurrentValue, StrValidationError
from Appender import find_all_projects, AppenderClass, connect_to_word_file, fill_word_file
from Reporter import ReporterClass, datetime_getter, add_total_summ, merge_dataframes
from Reporter import summary_dataframe, create_excel_file, projects_list_getter
import tkinter as tk
from tkinter import ttk
from tkinter import scrolledtext
import winsound as ws
import math as m
import threading

projects_names = find_all_projects()
employees = connection_to_list_db()
text_dictionary = {"OperatorNameError": "Имя оператора не выбрано!",
                   "ValueError": "Введены неверные данные!",
                   "TypeError": "Стандарт с данным идентификационным\nномером не обнаружен!",
                   "ExcessOfTheValue": "Введена избыточная навеска!\nОстаток по данному флакону: {} мг",
                   "StringIsEmpty": "Строка запроса пуста!",
                   "KeyError": "Кэш последней строки пуст!",
                   "StandartIsDone": "Стандарт уже израсходован и списан!\nУтилизируйте данный флакон",
                   "Successful_removal": "Строка {} успешно удалена",
                   "Successful_change": "Строка {} успешно изменена",
                   "Successful_write_off": "Стандарт успешно списан",
                   "Successful_sample_write_off": "Навеска успешно списана",
                   "WaitingTime": "Книга {} используется\n другим пользователем!",
                   "SimpleWaitingTime": "Таблица excel уже используется\nдругим пользователем!",
                   "SimpleWordWaitingTime": "Документ Word уже используется\nдругим пользователем!",
                   "IdIsEmpty": "В таблице отсутствует ID-номер!",
                   "NoProjectName": "Наименование проекта не выбрано!",
                   "NoStandardName": "Не введено наименование стандарта!",
                   "NoSerialNumber": "Не введён серийный номер стандарта!",
                   "NoCurrentValue": "Не введен номинальный объём стандарта!"
                   }
global_cache = {}


class AutocompleteCombobox(ttk.Combobox):
    """Функция автозаполнения для ComboBox, взятая со StackOverflow"""
    def __init__(self, completion_list):
        ttk.Combobox.__init__(self, width=25)
        self.completion_list = completion_list
        self.hits = []
        self.hit_index = 0
        self.position = 0
        self.bind('<KeyRelease>', self.handle_keyrelease)
        self['values'] = self.completion_list
        self.key_dict = {"81": "й", "87": "ц", "69": "у", "82": "к", "84": "е", "89": "н", "85": "г", "73": "ш",
                         "79": "щ", "219": "х", "221": "ъ", "65": "ф", "83": "ы", "68": "в", "70": "а", "71": "п",
                         "72": "р", "74": "о", "75": "л", "76": "д", "186": "ж", "222": "э", "90": "я", "88": "ч",
                         "67": "с", "86": "м", "66": "и", "78": "т", "77": "ь", "188": "б", "190": "ю"}

    def autocomplete(self, delta=0):
        if delta:
            self.delete(self.position, tk.END)
        else:
            self.position = len(str(self.get()))
        hits = []
        for element in self.completion_list:
            if element.lower().startswith(str(self.get()).lower()):
                hits.append(element)
        if hits != self.hits:
            self.hit_index = 0
            self.hits = hits
        if hits == self.hits and self.hits:
            self.hit_index = (self.hit_index + delta) % len(str(self.hits))
        if self.hits:
            self.delete(0, tk.END)
            self.insert(0, self.hits[self.hit_index])
            self.select_range(self.position, tk.END)

    def handle_keyrelease(self, event):
        if event.keycode == 8:
            self.delete(self.index(tk.INSERT), tk.END)
            self.position = self.index(tk.END)
        if event.keycode == 37:
            if self.position < self.index(tk.END):  # delete the selection
                self.delete(self.position, tk.END)
            else:
                self.position = self.position - 1  # delete one character
                self.delete(self.position, tk.END)
        if event.keycode == 39:
            self.position = self.index(tk.END)  # go to end (no selection)
        if str(event.keycode) in self.key_dict:
            self.autocomplete()


class MainWindow(MainWindowClass):
    def __init__(self, parent):
        super(MainWindow, self).__init__(parent)
        self.parent = parent
        self.pack(fill=tk.BOTH, expand=1)

        self.parent.title("BIOCAD Standarts counter")

        w = 370
        h = 140

        sw = self.parent.winfo_screenwidth()
        sh = self.parent.winfo_screenheight()

        x = (sw - w) / 2
        y = (sh - h) / 2
        self.parent.geometry('%dx%d+%d+%d' % (w, h, x, y))

        self.parent.resizable(False, False)

        self.input_index = tk.StringVar()
        self.input_sample = tk.StringVar()

        self.input_index_row = tk.Entry(self.parent, width=28, textvariable=self.input_index)
        self.input_index_row.place(x=180, y=12)
        self.input_index_row.focus_set()
        self.input_sample_row = tk.Entry(self.parent, width=28, textvariable=self.input_sample)
        self.input_sample_row.place(x=180, y=42)

        self.variable = tk.StringVar()
        global employees
        self.box = AutocompleteCombobox(employees)
        self.box.place(x=180, y=72)

        tk.Label(self.parent, text="Индекс образца:").place(x=12, y=12)
        tk.Label(self.parent, text="Масса навески (мг):").place(x=12, y=42)
        tk.Label(self.parent, text="Имя оператора:").place(x=12, y=72)

        tk.Button(self.parent, text="OK", width=20, command=self.start_program, bd=2).place(relx=0.32, y=100)

        """Добавим меню для выгрузки данных и добавления новых строк в таблицы"""
        main_menu = tk.Menu(self.parent)
        self.parent.config(menu=main_menu)
        file_menu = tk.Menu(main_menu, tearoff=0)
        file_menu.add_command(label="Добавить стандарты", command=start_appender_window)
        file_menu.add_command(label="Создать отчет", command=start_report_window)
        file_menu.add_command(label="Помощь")

        main_menu.add_cascade(label="Меню", menu=file_menu)

        """Теперь добавим меню для исправления ошибок, которые пользователь мог совершить"""
        help_menu = tk.Menu(main_menu, tearoff=0)
        help_menu.add_command(label="Исправить навеску", command=input_new_subtractor)
        help_menu.add_command(label="Убрать последнюю строку", command=delete_last_row)
        help_menu.add_command(label="Списать стандарт", command=self.write_off)

        main_menu.add_cascade(label="Редактирование", menu=help_menu)

        """Создадим функционал перехода на новый виджет через нажатие на клавишу
        Enter"""

        self.input_index_row.bind("<Return>",lambda func: self.input_sample_row.focus())
        self.input_sample_row.bind("<Return>",lambda func: self.box.focus())
        self.box.bind("<Return>",lambda func: self.input_index_row.focus())


    def start_program(self):
        """Основная функция, отвечающая за реализацию
        пользовательского функционала счетчика"""

        try:
            employee_name = str(self.box.get())
            if not employee_name:
                raise OperatorNameError("Не передано имя оператора!")
            print("Получено имя оператора")

            index = str(self.input_index_row.get())
            if not index:
                raise StringIsEmpty("Строка запроса пуста!")
            print("Получено значение индекса")

            """Приведем навеску к int-виду"""
            first_sample = self.input_sample_row.get()
            if "," in first_sample:
                first_sample = first_sample.replace(",", ".")
            sample = m.ceil(float(first_sample))

            """Навеска обязательно должна быть больше нуля"""
            if sample <= 0:
                raise ValueError
            print("Получено значение навески")
            search = SearcherClass(index)
            print("Создан экземпляр класса открытия и поиска")
            work_book, name_of_work_book = search.connection_to_db()
            print("Получен экземпляр книги")
            page, carriage = search.find_page_and_carriage_position(work_book)
            print("Получен экземпляр объекта листа книги и каретка")
            row, cas_cache, catalogue_cache, name_cache = search.searching_for_match(page, carriage)
            print("Найдена нужная строка и получены кэш-словари")
            print(row)
            input_class = InputClass(page, sample, row, carriage, cas_cache,
                                     catalogue_cache, name_cache, employee_name)
            print("Создан экземпляр класса вставки нужных значений")
            current_value = input_class.correct_input()

            if current_value:
                ChildWindow(240, 120).create_window(21, 23, 45, 73,
                                                    text_dictionary["ExcessOfTheValue"].format(current_value))
                raise ExcessOfTheValue(
                    "Введена избыточная навеска! Остаток по данному флакону: {}".format(current_value)
                )
            print("Введен основной массив значений")
            input_class.correct_cas_input()
            print("Введены CAS-значений")
            saving_status = safe_saving(work_book, name_of_work_book, 2)
            work_book.close()
            if not saving_status:
                ChildWindow(240, 120).create_window(15, 23, 45, 73,
                                                    text_dictionary["WaitingTime"].format(name_of_work_book[:-5]))
                raise WaitingTime("Книга уже кем-то используется! Сохранение невозможно")
            print("Excel-файл сохранён и закрыт \n")
            ChildWindow(240, 120).create_window(41, 33, 45, 63, text_dictionary["Successful_sample_write_off"], "Успех!")

            """Формируем глобальный кэш для возможных исправлений"""
            global global_cache
            global_cache["table_name"] = name_of_work_book
            global_cache["current_row"] = carriage + 1

        except StringIsEmpty:
            ChildWindow(240, 130).create_window(54, 33, 45, 63, text_dictionary["StringIsEmpty"])
        except OperatorNameError:
            ChildWindow(240, 130).create_window(39, 33, 45, 63, text_dictionary["OperatorNameError"])
        except ValueError:
            ChildWindow(240, 120).create_window(42, 33, 45, 63, text_dictionary["ValueError"])
            print("Введена неверная навеска!")
        except TypeError:
            ChildWindow(240, 120).create_window(3, 23, 45, 73, text_dictionary["TypeError"])
            print("Стандарт с данным идентификационным номером не обнаружен!")
        except StandartIsDone:
            ChildWindow(240, 120).create_window(17, 23, 45, 73,
                                                text_dictionary["StandartIsDone"])
        except StrValidationError as x:
            print(x)
        finally:
            self.input_index_row.delete(0, tk.END)
            self.input_sample_row.delete(0, tk.END)
            self.input_index_row.focus_set()

    def write_off(self):
        """Функция, отвечающая за списание стандарта пользователем"""
        try:
            employee_name = str(self.box.get())
            if not employee_name:
                print("Имя оператора не выбрано!")
                ChildWindow(240, 130).create_window(39, 33, 45, 63, text_dictionary["OperatorNameError"])
                raise OperatorNameError("Не передано имя оператора!")

            index = str(self.input_index_row.get())
            if not index:
                ChildWindow(240, 130).create_window(54, 33, 45, 63, text_dictionary["StringIsEmpty"])
                raise StringIsEmpty("Строка запроса пуста!")
            write_off_standart(index, employee_name)
            ChildWindow(240, 120).create_window(41, 33, 45, 63,
                                                text_dictionary["Successful_write_off"], "Успех!")
        except TypeError:
            ChildWindow(240, 120).create_window(3, 23, 45, 73, text_dictionary["TypeError"])
            print("Стандарт с данным идентификационным номером не обнаружен!")
        except StandartIsDone:
            ChildWindow(240, 120).create_window(17, 23, 45, 73,
                                                text_dictionary["StandartIsDone"])
        except WaitingTime:
            ChildWindow(240, 120).create_window(22, 23, 45, 73, text_dictionary["SimpleWaitingTime"])
        finally:
            self.input_index_row.delete(0, tk.END)
            self.input_sample_row.delete(0, tk.END)
            self.input_index_row.focus_set()


class ChildWindow(MainChildWindowClass):
    def __init__(self, w: int, h: int):
        super(ChildWindow, self).__init__()

        self.w = w
        self.h = h

        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x = (sw - w) / 2
        y = (sh - h) / 2

        self.geometry('%dx%d+%d+%d' % (w, h, x, y))

        """Запрещаем работу с главным окном, пока не будет закрыто дочернее окно"""
        self.grab_set()
        self.focus_set()

    def create_window(self, x_label: int, y_label: int,
                      x_button: int, y_button: int,
                      label_text: str, header="Ошибка!"):
        self.title(header)

        """Размещаем Label с нужным текстом"""
        tk.Label(self, text=label_text).place(x=x_label, y=y_label)

        """Размещаем кнопку"""
        tk.Button(self, text="OK", width=20, command=self.destroy).place(x=x_button, y=y_button)

        """Запрещаем изменять размер экрана"""
        self.resizable(False, False)

        """Воспроизводим звук системной ошибки каждый раз, когда запускается экземпляр класса с данным атрибутом"""
        ws.PlaySound("*", ws.SND_ASYNC)

    def private_window(self, table_name, current_row):
        """Приватное окно, которое отвечает за исправление пользователем навески"""
        self.title("Навеска")
        tk.Label(self, text="Введите навеску:").place(x=15, y=33)
        input_subtractor = tk.IntVar()
        input_subtractor_row = tk.Entry(self, width=15, textvariable=input_subtractor)
        input_subtractor_row.place(x=130, y=33)

        """Функция, которая запускается при нажатии кнопки.
        Ссылается на функцию-исполнитель, которая и проверяет """
        def starting_private_window_button():
            try:
                """Приводим навеску к int-виду"""
                first_subtractor = input_subtractor_row.get()
                if "," in first_subtractor:
                    first_subtractor = first_subtractor.replace(",", ".")
                subtractor = m.ceil(float(first_subtractor))

                checker = CorrectMistakes(table_name, current_row, subtractor).correct_subtractor()
                """Проверим, соблюдено ли условие допустимого размера навески.
                Если нет, возбуждаем ошибку превышения навески и выводим
                пользователю окно ошибки"""
                if checker:
                    ChildWindow(240, 120).create_window(21, 23, 45, 73,
                                                        text_dictionary["ExcessOfTheValue"].format(checker))
                    raise ExcessOfTheValue("Введена избыточная навеска!")
            except ValueError:
                ChildWindow(240, 120).create_window(42, 33, 45, 63, text_dictionary["ValueError"])
                print("Введена неверная навеска!")
            except WaitingTime:
                ChildWindow(240, 120).create_window(22, 23, 45, 73, text_dictionary["SimpleWaitingTime"])
            else:
                ChildWindow(240, 120).create_window(41, 33, 45, 63,
                                                    text_dictionary["Successful_change"].format(current_row),
                                                    "Успех!")
            finally:
                input_subtractor_row.delete(0, tk.END)
                input_subtractor_row.focus_set()

        tk.Button(self, text="OK", width=20, command=starting_private_window_button).place(x=45, y=63)
        self.resizable(False, False)
        self.grab_set()
        self.focus_set()

    def append_window(self):
        """Приватное окно, которое позволяет вносить пользователю новые стандарты"""
        self.title("Внесение стандартов")

        tk.Label(self, text="Проект:").place(x=10, y=20)
        global projects_names
        box = tk.ttk.Combobox(self, state='readonly', width=30)
        box['values'] = projects_names
        box.place(x=130, y=20)

        tk.Label(self, text="Номер по каталогу:").place(x=10, y=45)
        input_catalogue_number = tk.StringVar()
        input_catalogue_number_row = tk.Entry(self, width=33, border=2, textvariable=input_catalogue_number)
        input_catalogue_number_row.place(x=130, y=45)

        tk.Label(self, text="CAS-номер:").place(x=10, y=70)
        input_cas_number = tk.StringVar()
        input_cas_number_row = tk.Entry(self, width=33, border=2, textvariable=input_cas_number)
        input_cas_number_row.place(x=130, y=70)

        tk.Label(self, text="Наименование:").place(x=10, y=95)
        input_name = scrolledtext.ScrolledText(self, width=25, height=2, border=2)
        input_name.place(x=130, y=95)

        tk.Label(self, text="Производитель:").place(x=10, y=138)
        input_manufacturer = tk.StringVar()
        input_manufacturer_row = tk.Entry(self, width=33, border=2, textvariable=input_manufacturer)
        input_manufacturer_row.place(x=130, y=138)

        tk.Label(self, text="Серийный номер:").place(x=10, y=163)
        input_serial_number = tk.StringVar()
        input_serial_number_row = tk.Entry(self, width=33, border=2, textvariable=input_serial_number)
        input_serial_number_row.place(x=130, y=163)

        tk.Label(self, text="Номинал (мг):").place(x=10, y=188)
        input_nominal_value = tk.IntVar()
        input_nominal_value_row = tk.Entry(self, width=8, border=2, textvariable=input_nominal_value)
        input_nominal_value_row.place(x=130, y=188)

        tk.Label(self, text="Несн. остаток:").place(x=185, y=188)
        input_untouchable_value = tk.IntVar()
        input_untouchable_value_row = tk.Entry(self, width=8, border=2, textvariable=input_untouchable_value)
        input_untouchable_value_row.place(x=279, y=188)

        current_value_state = tk.StringVar()
        current_value_state.set("disabled")

        tk.Label(self, text="Остаток во флаконе:").place(x=10, y=213)
        input_current_value = tk.IntVar()
        input_current_value_row = tk.Entry(self, width=8, border=2, textvariable=input_current_value,
                                           state=current_value_state.get())
        input_current_value_row.place(x=130, y=213)

        def change_current_value_status():
            if current_value_state == "disabled":
                input_current_value_row.config(state="disabled")
            else:
                input_current_value_row.config(state="normal")
            input_current_value_row.config(state=current_value_state.get())

        current_value_status = tk.Checkbutton(self,
                                              text="Текущее количество субстанции \n во флаконе равно номинальному",
                                              variable=current_value_state, onvalue="disabled", offvalue="normal",
                                              command=change_current_value_status)
        current_value_status.place(x=125, y=238)

        tk.Label(self, text="Количество:").place(x=10, y=283)
        number_of_standards = tk.Spinbox(self, width=8, from_=1, to=100, border=2)
        number_of_standards.place(x=130, y=283)

        self.resizable(False, False)
        self.grab_set()
        self.focus_set()

        def start_appender_program():
            try:
                project = str(box.get())
                if not project:
                    ChildWindow(240, 120).create_window(20, 33, 45, 63, text_dictionary["NoProjectName"])
                    raise NoProjectName("Имя проекта не выбрано!")
                print("Получено значение проекта")

                catalogue_number = input_catalogue_number_row.get()
                print("Получено значение каталожного номера")

                cas_number = input_cas_number_row.get()
                print("Получено значение cas-номера")

                name = input_name.get('1.0', 'end-1c')
                name = name.rstrip()
                if not name:
                    ChildWindow(240, 120).create_window(15, 33, 45, 63, text_dictionary["NoStandardName"])
                    raise NoStandardName("Не введено имя стандарта!")
                print("Получено значение наименования стандарта")

                manufacturer = input_manufacturer_row.get()
                print("Получено значение наименования производителя")

                serial_number = input_serial_number_row.get()
                if not serial_number:
                    ChildWindow(240, 120).create_window(13, 33, 45, 63, text_dictionary["NoSerialNumber"])
                    raise NoSerialNumber("Не введён серийный номер стандарта!")
                print("Получено значение серийного номера")

                nominal_value = input_nominal_value_row.get()
                if not nominal_value:
                    ChildWindow(240, 120).create_window(2, 33, 45, 63, text_dictionary["NoCurrentValue"])
                    raise NoCurrentValue("Не введен номинальный объём стандарта!")
                print("Получено значение номинального объёма")

                untouchable_value = input_untouchable_value_row.get()
                print("Получено значение неснижаемого остатка")

                current_value = input_current_value_row.get()
                print("Получено значение остатка субстанции во флаконе")
                print(current_value)

                number_of_standard = int(number_of_standards.get())
                print("Получено значение количества вносимых флаконов")

                appender = AppenderClass(project, catalogue_number, cas_number, name, manufacturer,
                                         serial_number, nominal_value, number_of_standard, untouchable_value,
                                         current_value)
                print("Создан экземпляр класса вставщика")

                book, file_name = appender.connection_to_db()
                print("Получены объект книги и имя файла")

                main_page, id_page, identification, carriage = appender.get_id(book)
                print("Получены объекты двух страниц книги, последний идентификатор и каретка")

                previous_mass = appender.get_previous_mass(main_page, carriage)
                print("Получено значение предыдущей общей массы")

                id_numbers = appender.append_to_db(main_page, id_page, identification, carriage, previous_mass)
                print("Осуществлена вставка новых значений в таблицу")

                document = connect_to_word_file()
                print("Создан объект документа Word")

                fill_word_file(document, name, serial_number, nominal_value, current_value, id_numbers)
                print("Документ Word заполнен")

                word_saving_status = safe_saving(document, "Идентификаторы.docx", 2)

                if not word_saving_status:
                    ChildWindow(240, 120).create_window(22, 23, 45, 73, text_dictionary["SimpleWordWaitingTime"])
                    raise WaitingTime("Word-файл уже кем-то используется! Сохранение невозможно")
                print("Документ Word сохранён и закрыт")

                saving_status = safe_saving(book, file_name, 2)
                book.close()

                if not saving_status:
                    ChildWindow(240, 120).create_window(22, 23, 45, 73, text_dictionary["SimpleWaitingTime"])
                    raise WaitingTime("Книга уже кем-то используется! Сохранение невозможно")
                print("Документ Excel сохранён и закрыт")
            except ValueError:
                ChildWindow(240, 120).create_window(42, 33, 45, 63, text_dictionary["ValueError"])
                print("Введены неверные значения полей!")
            except IdIsEmpty:
                ChildWindow(240, 120).create_window(25, 33, 45, 63, text_dictionary["IdIsEmpty"])
                print("Значение поля идентификатора пусто!")
            finally:
                input_catalogue_number_row.delete(0, tk.END)
                input_cas_number_row.delete(0, tk.END)
                input_name.delete("1.0", tk.END)
                input_manufacturer_row.delete(0, tk.END)
                input_serial_number_row.delete(0, tk.END)
                input_nominal_value_row.delete(0, tk.END)
                input_current_value_row.delete(0, tk.END)
                input_untouchable_value_row.delete(0, tk.END)
                input_catalogue_number_row.focus_set()

        tk.Button(self, text="Добавить", width=20, command=start_appender_program, bd=3).place(relx=0.05, y=313)
        tk.Button(self, text="Закрыть", width=20, command=self.destroy, bd=3).place(relx=0.52, y=313)

    def report_window(self):
        """Приватное окно, которое позволяет пользователю собирать отчет"""
        self.title("Сбор отчета")

        text_state = tk.StringVar()
        text_state.set("disabled")

        def change_text_status():
            projects_to_report.delete("1.0", tk.END)
            if text_state.get() == "normal":
                box.config(state="readonly")
            else:
                box.config(state="disabled")
            projects_to_report.config(state=text_state.get())

        def combobox_selected(event):
            combobox_value = box.get()
            projects_to_report.insert(tk.END, combobox_value + "; ")

        tk.Label(self, text="Список проектов:").place(x=10, y=20)
        global projects_names
        box = tk.ttk.Combobox(self, state="disabled", width=30)
        box['values'] = projects_names
        box.bind("<<ComboboxSelected>>", combobox_selected)
        box.place(x=130, y=20)

        tk.Label(self, text="Проекты для отчета:").place(x=10, y=50)
        projects_to_report = tk.Text(self, width=25, height=3, border=2, state=text_state.get())
        projects_to_report.place(x=130, y=50)

        tk.Label(self, text="Введите даты:").place(x=10, y=110)
        input_date = tk.StringVar()
        input_date_value = tk.Entry(self, width=33, border=2, textvariable=input_date)
        input_date_value.place(x=130, y=110)

        global_report = tk.Checkbutton(self, text="Собрать глобальный отчет", variable=text_state, onvalue="disabled",
                                       offvalue="normal", command=change_text_status)
        global_report.place(x=125, y=135)

        def start_report():

            dates = str(input_date_value.get())
            
            if dates == "" or not dates:
                date_list = []
                file_name = "Отчет.xlsx"
            else:
                try:
                    datetime_getter(dates)
                except ValueError:
                    ChildWindow(240, 120).create_window(42, 33, 45, 63, text_dictionary["ValueError"])
                date_list = datetime_getter(dates)
                file_name = "Отчет за {}.xlsx".format(dates)

            progress = ChildWindow(300, 100)
            progress.title("Прогресс выполнения")
            pb = ttk.Progressbar(progress, length=280, mode="determinate")
            pb.pack(expand=1)

            total_data_list = []  

            if text_state.get() == "normal":
                project_string = projects_to_report.get('1.0', 'end')
                projects = projects_list_getter(project_string)
                for i in range(len(projects)):
                    if "BCD" not in projects[i]:
                        projects.pop(i)
            else:
                global projects_names
                projects = projects_names

            k = 0
            length = len(projects)
            tk.Label(progress, text="Проектов найдено: {}".format(length)).place(relx=0.25, rely=0.15)

            def start_count():
                """Потоково выполняемая функция, реализованная для
                корректного воплощения ProgressBar, открывающегося
                при нажатии на кнопку создания отчёта"""
                for project in projects:
                    try:
                        data = ReporterClass(project)
                    except FileNotFoundError:
                        nonlocal k
                        k += 1
                        pb['value'] = int((k / length) * 100)
                        continue
                    if data.df.empty:
                        k += 1
                        pb['value'] = int((k / length) * 100)
                        continue
                    else:
                        print(project)
                        done_data, list_of_total_data = data.create_done_data_frame(date_list)
                        not_done_data, list_of_not_done_data = data.create_not_done_data_frame()
                        cas_unique_names, catalogue_unique_names, unique_names = data.create_single_not_done_data(
                            not_done_data)
                        cas_unique_names, catalogue_unique_names, unique_names = data.empty_standards(
                            done_data, cas_unique_names, catalogue_unique_names, unique_names)
                        cas_unique_names, catalogue_unique_names, unique_names = data.not_empty_data(
                            list_of_total_data, done_data, cas_unique_names, catalogue_unique_names, unique_names)
                        cas_unique_names, catalogue_unique_names, unique_names = add_total_summ(
                            cas_unique_names, catalogue_unique_names, unique_names)
                        cas_unique_names, catalogue_unique_names, unique_names = data.add_balance(
                            cas_unique_names, catalogue_unique_names, unique_names)
                        total_data = merge_dataframes(cas_unique_names, catalogue_unique_names, unique_names)
                        total_data_list.append(total_data)
                        k += 1
                        pb['value'] = int((k / length) * 100)

                final_df = summary_dataframe(total_data_list)
                create_excel_file(final_df, file_name)
                progress.destroy()

            thread1 = threading.Thread(target=start_count)
            thread1.start()

        tk.Button(self, text="Создать отчет", width=20, command=start_report, bd=3).place(relx=0.05, y=170)
        tk.Button(self, text="Закрыть", width=20, command=self.destroy, bd=3).place(relx=0.52, y=170)

        self.resizable(False, False)
        self.grab_set()
        self.focus_set()


def delete_last_row():
    """Функция отвечает за удаление последней внесённой строки в БД"""
    try:
        table_name = global_cache["table_name"]
        current_row = global_cache["current_row"]
        deleter = CorrectMistakes(table_name, current_row, 0)
        deleter.delete_current_row()
        global_cache.clear()
        ChildWindow(240, 120).create_window(43, 33, 45, 63,
                                            text_dictionary["Successful_removal"].format(current_row),
                                            "Успех!")
    except KeyError:
        ChildWindow(240, 120).create_window(38, 33, 45, 63, text_dictionary["KeyError"])
    except WaitingTime:
        ChildWindow(240, 120).create_window(22, 23, 45, 73, text_dictionary["SimpleWaitingTime"])


def input_new_subtractor():
    """Функция отвечает за исправление последней внесённой строки в БД"""
    try:
        table_name = global_cache["table_name"]
        current_row = global_cache["current_row"]
        ChildWindow(240, 120).private_window(table_name, current_row)
    except KeyError:
        ChildWindow(240, 120).create_window(38, 33, 45, 63, text_dictionary["KeyError"])
    global_cache.clear()


def start_appender_window():
    ChildWindow(352, 350).append_window()


def start_report_window():
    ChildWindow(352, 210).report_window()


def main():
    root_element = tk.Tk()
    MainWindow(root_element)
    root_element.mainloop()


if __name__ == "__main__":
    main()
