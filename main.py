import sys
import win32com.client
import xlwings as xw
from PyQt5.QtWidgets import *
from PyQt5 import QtCore, QtGui
from PyQt5.QtGui import QFont, QColor
import pandas as pd
from openpyxl import load_workbook, workbook


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # Создаём QTabWidget
        # self.table5 = None
        # self.table4 = None
        # self.table3 = None
        # self.table2 = None
        # self.table1 = None
        self.tabs = QTabWidget()

        self.setWindowTitle("Рассылка ГО и ЧС")
        self.setGeometry(0, 0, 1920, 1080)

        # Установка шрифта для подсказок
        QToolTip.setFont(QFont('Times New Roman', 12))

        # Создаем верхние кнопки и инициализируем их работу
        self.check_button = QPushButton('Проверка', self)
        self.check_button.setToolTip("<h3>Пройти верификацию</h3>")
        self.check_button.clicked.connect(self.check_action)

        self.staff_button = QPushButton('Проверка и рассылка сотрудникам', self)
        self.staff_button.setToolTip("<h3>Пройти верификацию и разослать уведомление сотрудникам</h3>")
        self.staff_button.clicked.connect(self.staff_action)

        self.boss_button = QPushButton('Рассылка начальникам отдела', self)
        self.boss_button.setToolTip("<h3>Разослать уведомление начальникам отделов</h3>")
        self.boss_button.clicked.connect(self.boss_action)

        self.responsible_button = QPushButton('Рассылка ответственным', self)
        self.responsible_button.setToolTip("<h3>Рассылка ответственным</h3>")
        self.responsible_button.clicked.connect(self.responsible_action)

        # Создаем QTableWidget
        self.tableWidget = QTableWidget()
        self.tableWidget.setEditTriggers(QTableWidget.AllEditTriggers)  # Разрешаем редактирование ячеек

        # Создаем QLabel и QLineEdit для ввода текста для поиска
        self.searchLabel = QLabel("Поиск:")
        self.searchLabel.setStyleSheet("\n" "font: 15pt \"Open Sauce One SemiBold\"")
        self.searchLineEdit = QLineEdit()
        self.searchLineEdit.textChanged.connect(self.search_table)  # Подключаем функцию поиска

        # Создаем кнопку для сохранения файла
        self.button1 = QPushButton("Сохранить файл", self)
        self.button1.clicked.connect(self.save_table)

        # Создаем кнопку для добавления нового столбца
        self.addButtonColumn = QPushButton("Добавить столбец", self)
        self.addButtonColumn.clicked.connect(self.add_column)

        # Создаем кнопку для добавления новой строки
        self.addButtonRow = QPushButton("Добавить строку", self)
        self.addButtonRow.clicked.connect(self.add_row)

        def apply_shadow_effect(widget):  # Функция прикрепляющая эффект тени к кнопке
            shadow = QGraphicsDropShadowEffect()
            shadow.setBlurRadius(20)  # Радиус размытия
            shadow.setXOffset(5)  # Смещение тени по горизонтали
            shadow.setYOffset(5)  # Смещение тени по вертикали
            widget.setGraphicsEffect(shadow)

        for button in [self.button1, self.check_button, self.staff_button, self.boss_button, self.responsible_button,
                       self.addButtonColumn, self.addButtonRow]:
            button.setStyleSheet("\n" "font: 15pt \"Open Sauce One SemiBold\"")
            apply_shadow_effect(button)

        # Объявление функций создания таблиц
        self.load_excel_data()
        self.load_excel_data2()
        self.load_excel_data3()
        self.load_excel_data4()
        self.load_excel_data5()
        self.load_excel_data6()

        # Добавляем таблицы на вкладки
        self.tabs.addTab(self.table1, "Сотрудники")
        self.tabs.addTab(self.table2, "Электробезопасность")
        self.tabs.addTab(self.table3, "Охрана труда")
        self.tabs.addTab(self.table4, "Пожарная безопасность")
        self.tabs.addTab(self.table5, "Мероприятия")
        self.tabs.addTab(self.table6, "Ответственные")

        # Привязка сигнала переключения вкладки к функции и индекс вкладки
        self.current_tab_index = 0
        self.tabs.currentChanged.connect(self.tab_changed)

        # Создание layout'а
        hlayout = QHBoxLayout()

        # Создаем горизонтальный QVBoxLayout и добавляем в него верхние кнопки
        horizontal_layout = QHBoxLayout()
        horizontal_layout.addWidget(self.check_button)
        horizontal_layout.addWidget(self.staff_button)
        horizontal_layout.addWidget(self.boss_button)
        horizontal_layout.addWidget(self.responsible_button)

        # Создание "контейнера"
        container = QWidget()
        # Установка размеров "контейнера"
        container.setFixedSize(1300, 100)
        # Установка layout'а для "контейнера"
        container.setLayout(horizontal_layout)

        # "Загоняем" контейнер обратно в горизонтальный layout
        hlayout.addWidget(container)
        # Устанавливаем отступы для layout'а
        hlayout.setContentsMargins(0, 0, 150, 0)

        # Создаем вертикальный QVBoxLayout и добавляем в него компоненты
        layout = QVBoxLayout()
        layout.setContentsMargins(50, 60, 50, 30)
        layout.addLayout(hlayout)
        layout.addSpacing(80)
        horizontal_layout2 = QHBoxLayout()
        horizontal_layout2.addWidget(self.searchLabel)
        horizontal_layout2.addWidget(self.searchLineEdit)
        layout.addLayout(horizontal_layout2)
        layout.addWidget(self.tabs)
        layout.addSpacing(30)

        # Создаем горизонтальный макет и добавляем в него компоненты
        horizontal_layout1 = QHBoxLayout()
        horizontal_layout1.addWidget(self.addButtonColumn)
        horizontal_layout1.addWidget(self.addButtonRow)
        horizontal_layout1.addWidget(self.button1)

        # Добавляем горизонтальный макет в вертикальный макет
        layout.addLayout(horizontal_layout1)

        # Создаем QWidget и устанавливаем QVBoxLayout в качестве макета
        widget = QWidget()
        widget.setLayout(layout)
        self.setCentralWidget(widget)

        # Загружаем данные из файла Excel
        self.showMaximized()
        self.show()

    # Функция, определяющая включенную вкладку
    def tab_changed(self, index):
        # Сохраняем индекс текущей вкладки
        self.current_tab_index = index
        # Печать названия текущей вкладки
        print("Вкладка переключена на:", self.tabs.tabText(index))

    def add_column(self):
        current_column_count = self.tabs.currentWidget().columnCount()
        self.tabs.currentWidget().setColumnCount(current_column_count+1)

    def add_row(self):
        current_row_count = self.tabs.currentWidget().rowCount()
        self.tabs.currentWidget().setRowCount(current_row_count+1)

    # Функция, отвечающая за проверку
    def check_action(self):
        vba_book = xw.Book("filename.xlsm")
        vba_macro2 = vba_book.macro("Макрос1(0)")
        vba_macro2()
        vba_book.save()  # Сохраняем изменения в файле
        vba_book.close()  # Закрываем файл
        print("Проверка")
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Quit()

    # Функция, отвечающая за проверку и рассылку сотрудникам
    def staff_action(self):
         vba_book = xw.Book("filename.xlsm")
         vba_macro3 = vba_book.macro("Макрос1(1)")
         vba_macro3()
         vba_book.save()  # Сохраняем изменения в файле
         vba_book.close()  # Закрываем файл
         print("Рассылка сотрудникам")
         excel = win32com.client.Dispatch("Excel.Application")
         excel.Quit()

    # Функция, отвечающая за рассылку начальникам отдела
    def boss_action(self):
         vba_book = xw.Book("filename.xlsm")
         vba_macro4 = vba_book.macro("Макрос1(2)")
         vba_macro4()
         vba_book.save()  # Сохраняем изменения в файле
         vba_book.close()  # Закрываем файл
         print("Рассылка начальникам отдела")
         excel = win32com.client.Dispatch("Excel.Application")
         excel.Quit()

    # Функция, отвечающая за рассылку ответственным
    def responsible_action(self):
         vba_book = xw.Book("filename.xlsm")
         vba_macro5 = vba_book.macro("Макрос1(3)")
         vba_macro5()
         vba_book.save()  # Сохраняем изменения в файле
         vba_book.close()  # Закрываем файл
         print("Рассылка ответственным")
         excel = win32com.client.Dispatch("Excel.Application")
         excel.Quit()

    def load_excel_data(self):
        # Загружаем файл Excel с помощью библиотеки pandas
        df = pd.read_excel('filename.xlsm', sheet_name='Сотрудники', skiprows=7, header = 0, usecols=(lambda x: x != '№ п/п'))

        self.table1 = QTableWidget()
        self.table1.setSortingEnabled(True)
        # Устанавливаем количество строк и столбцов в QTableWidget
        self.table1.setRowCount(df.shape[0])
        self.table1.setColumnCount(df.shape[1])

        # Устанавливаем заголовок
        headers = df.columns.tolist()
        self.table1.setHorizontalHeaderLabels(headers)

        # Заполняем QTableWidget данными из DataFrame
        for i, row in df.iterrows():
            for j, value in enumerate(row):
                if pd.isnull(value):
                    value = ""
                elif str(value).__contains__(':') or str(value).__contains__('-'):
                    year = value.year
                    month = value.month
                    day = value.day
                    value = f'{day}.{month}.{year}'
                item = QTableWidgetItem(str(value))
                self.table1.setItem(i, j, item)

        # Расширяем столбцы таблицы для соответствия содержимому
        self.table1.resizeColumnsToContents()

        # Устанавливаем режим растягивания последнего столбца таблицы
        self.table1.horizontalHeader().setStretchLastSection(True)

    def search_table(self, text):
        # Получаем текст для поиска
        currentTable = self.tabs.currentWidget()
        # очистка выделения
        currentTable.clearSelection()

        # поиск и обновление видимости строк
        for i in range(currentTable.rowCount()):
            matches = False
            for j in range(currentTable.columnCount()):
                item = currentTable.item(i, j)
                if item and text in item.text():
                    matches = True
                    break
            currentTable.setRowHidden(i, not matches)

    def load_excel_data2(self):
        # Загружаем файл Excel с помощью библиотеки pandas
        df2 = pd.read_excel("filename.xlsm", sheet_name='Электробезопасн', skiprows=7, header=0,
                            usecols=(lambda x: x != '№ п/п'))

        self.table2 = QTableWidget()
        self.table2.setSortingEnabled(True)
        # Устанавливаем количество строк и столбцов в QTableWidget
        self.table2.setRowCount(df2.shape[0])
        self.table2.setColumnCount(df2.shape[1])

        # Устанавливаем заголовок
        headers = df2.columns.tolist()
        self.table2.setHorizontalHeaderLabels(headers)

        # Заполняем QTableWidget данными из DataFrame
        for i, row in df2.iterrows():
            for j, value in enumerate(row):
                if pd.isnull(value):
                    value = ""
                if str(value).__contains__(':') or str(value).__contains__('-'):
                    year = value.year
                    month = value.month
                    day = value.day
                    value = f'{day}.{month}.{year}'
                item = QTableWidgetItem(str(value))
                self.table2.setItem(i, j, item)

        # Расширяем столбцы таблицы для соответствия содержимому
        self.table2.resizeColumnsToContents()

        # Устанавливаем режим растягивания последнего столбца таблицы
        self.table2.horizontalHeader().setStretchLastSection(True)

    def load_excel_data3(self):
        # Загружаем файл Excel с помощью библиотеки pandas
        df3 = pd.read_excel("filename.xlsm", sheet_name='Охрана труда (г', skiprows=7, header=0,
                            usecols=(lambda x: x != '№ п/п'))
        # Создаём вторую таблицу
        self.table3 = QTableWidget()
        self.table3.setSortingEnabled(True)
        self.table3.setRowCount(df3.shape[0])
        self.table3.setColumnCount(df3.shape[1])

        # Устанавливаем заголовок
        headers = df3.columns.tolist()
        self.table3.setHorizontalHeaderLabels(headers)

        # Заполняем QTableWidget данными из DataFrame
        for i, row in df3.iterrows():
            for j, value in enumerate(row):
                if pd.isnull(value):
                    value = ""
                if str(value).__contains__(':') or str(value).__contains__('-'):
                    year = value.year
                    month = value.month
                    day = value.day
                    value = f'{day}.{month}.{year}'
                item = QTableWidgetItem(str(value))
                self.table3.setItem(i, j, item)

        # Расширяем столбцы таблицы для соответствия содержимому
        self.table3.resizeColumnsToContents()
        # Устанавливаем режим растягивания последнего столбца таблицы
        self.table3.horizontalHeader().setStretchLastSection(True)

    def load_excel_data4(self):
        # Загружаем файл Excel с помощью библиотеки pandas
        df4 = pd.read_excel("filename.xlsm", sheet_name='Пожарная безопа', skiprows=7, header=0,
                            usecols=(lambda x: x != '№ п/п'))
        # Создаём вторую таблицу
        self.table4 = QTableWidget()
        self.table4.setSortingEnabled(True)
        self.table4.setRowCount(df4.shape[0])
        self.table4.setColumnCount(df4.shape[1])

        # Устанавливаем заголовок
        headers = df4.columns.tolist()
        self.table4.setHorizontalHeaderLabels(headers)

        # Заполняем QTableWidget данными из DataFrame
        for i, row in df4.iterrows():
            for j, value in enumerate(row):
                if pd.isnull(value):
                    value = ""
                if str(value).__contains__(':') or str(value).__contains__('-'):
                    year = value.year
                    month = value.month
                    day = value.day
                    value = f'{day}.{month}.{year}'
                item = QTableWidgetItem(str(value))
                self.table4.setItem(i, j, item)

        # Расширяем столбцы таблицы для соответствия содержимому
        self.table4.resizeColumnsToContents()
        # Устанавливаем режим растягивания последнего столбца таблицы
        self.table4.horizontalHeader().setStretchLastSection(True)

    def load_excel_data5(self):
        # Загружаем файл Excel с помощью библиотеки pandas
        df5 = pd.read_excel("filename.xlsm", sheet_name='Мероприятия', skiprows=1, usecols=(lambda x: x != '№ п/п'))
        # Создаём вторую таблицу
        self.table5 = QTableWidget()
        self.table5.setSortingEnabled(True)
        self.table5.setRowCount(df5.shape[0])
        self.table5.setColumnCount(df5.shape[1])

        # Устанавливаем заголовок
        # header2 = df5.columns.tolist()
        # self.table5.setHorizontalHeaderLabels(header2)

        # Заполняем QTableWidget данными из DataFrame
        for i, row in df5.iterrows():
            for j, value in enumerate(row):
                item = QTableWidgetItem(str(value))
                self.table5.setItem(i, j, item)

        # Расширяем столбцы таблицы для соответствия содержимому
        self.table5.resizeColumnsToContents()
        # Устанавливаем режим растягивания последнего столбца таблицы
        self.table5.horizontalHeader().setStretchLastSection(True)

    def load_excel_data6(self):
        # Загружаем файл Excel с помощью библиотеки pandas
        df6 = pd.read_excel("filename.xlsm", sheet_name='Отвественные', skiprows=1, usecols=(lambda x: x != '№ п/п'))
        # Создаём вторую таблицу
        self.table6 = QTableWidget()
        self.table6.setSortingEnabled(True)
        self.table6.setRowCount(df6.shape[0])
        self.table6.setColumnCount(df6.shape[1])

        # Устанавливаем заголовок
        # header2 = df5.columns.tolist()
        # self.table5.setHorizontalHeaderLabels(header2)

        # Заполняем QTableWidget данными из DataFrame
        for i, row in df6.iterrows():
            for j, value in enumerate(row):
                if pd.isnull(value):
                    value = ""
                if str(value).__contains__(':') or str(value).__contains__('-'):
                    year = value.year
                    month = value.month
                    day = value.day
                    value = f'{day}.{month}.{year}'
                item = QTableWidgetItem(str(value))
                self.table6.setItem(i, j, item)

        # Расширяем столбцы таблицы для соответствия содержимому
        self.table6.resizeColumnsToContents()
        # Устанавливаем режим растягивания последнего столбца таблицы
        self.table6.horizontalHeader().setStretchLastSection(True)

    def save_table(self):
        currentTable2 = self.tabs.currentWidget()
        workbook = xw.Book("filename.xlsm")

        sheet = workbook.sheets[self.current_tab_index]

        if self.current_tab_index == 4 or self.current_tab_index == 5:
            for row in range(currentTable2.rowCount()):
                for col in range(currentTable2.columnCount()):
                    item = currentTable2.item(row, col)
                    if item is not None:
                        sheet.range(row + 3, col + 2).value = item.text()
        else:
            for row in range(currentTable2.rowCount()):
                for col in range(currentTable2.columnCount()):
                    item = currentTable2.item(row, col)
                    if item is not None:
                        sheet.range(row + 9, col + 2).value = item.text()

        workbook.save()  # Сохраняем изменения в файле
        workbook.close() # Закрываем файл
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Quit()

stylesheet12 = """
    MyWindow {
        background-color: #D7DCEE;
        background-image: url(1234.png);
    }

    QPushButton {
    background-color: white;
    color: #021F3F;
    border-style: outset;
    padding: 2px;
    border-width: 6px;
    border-radius: 8px;
    border-color: white;    
    min-width: 100px; 
    max-width: 350px; 
    min-height: 50px; 
    max-height: 100px;
    }

    QTableWidget {
        background-color: white;
        color: #000000;
    }

    QTableWidget::item {
        background-color: white;
    }

    QTableWidget::item:selected {
        background-color: #0080FF;
        color: #FFFFFF;
    }

    QTableWidget::item:hover {
        background-color: #808080;
    }

    QTableWidget::indicator {
        width: 20px;
        height: 20px;
    }

    QHeaderView::section {
        background-color: white;
        color: black;
        padding: 5px;
        font-size: 14px;
        font-family: Arial;
        border: 1px solid #C0C0C0;
    }

    QTabBar {
    background-color: white;    
    }

    QTabBar::tab {
    background: white; 
    color: #000000;
    border: 2px solid #C4C4C3;
    border-bottom-color: #C2C7CB;
    border-top-left-radius: 4px;
    border-top-right-radius: 4px;
    min-width: 8ex;
    padding: 5px;
    }

    QTabBar::tab:selected {
    background: #ffffff;
    color: #000000;
    }

    QTabBar::tab:!selected {
    margin-top: 2px;
    }
"""

if __name__ == "__main__":
    app = QApplication([])
    app.setStyleSheet(stylesheet12)
    window = MyWindow()
    window.show()
    app.exec()



