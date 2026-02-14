import sys
import os
import datetime
import sqlite3
import xlrd
import xlwt
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
import pyqtgraph as pg
import time


def create_table():
    tables = [i[0] for i in list(cur.execute('SELECT name FROM sqlite_master WHERE type="table"'))]
    
    cur.execute('CREATE TABLE IF NOT EXISTS Contract ( id INTEGER, Year INTEGER, Month TEXT, Time INTEGER, Count_tariff INTEGER, Count_money INTEGER )')
    
    cur.execute('''CREATE TABLE IF NOT EXISTS Database
(Number TEXT, Category TEXT, Comment TEXT, FIO TEXT, Job_title TEXT, Subdivision TEXT, Current_blocking TEXT, Date_current_blocking TEXT, TariffN TEXT,
IP_address TEXT, Tariff TEXT, ICCID_number TEXT, Date_issue TEXT, Date_activation TEXT, Date_change TEXT, Discription_change TEXT, Reg_government_services TEXT )''')

    cur.execute('CREATE TABLE IF NOT EXISTS SaveINFO ( id INTEGER, ParamDB TEXT, ParamPT TEXT, Next_number INTEGER )')
    
    cur.execute('CREATE TABLE IF NOT EXISTS SavePT ( id INTEGER, Start TEXT, Time TEXT, Type TEXT, INFO TEXT )')
    
    cur.execute('CREATE TABLE IF NOT EXISTS Tariff ( id INTEGER, Description TEXT )')
    
    con.commit()
    
    if not 'Database' in tables:
        columns = [j + i for i in MONTH_SHORT for j in ['tr1', 'tr2', 'tr3',
                                                        'ac1', 'ac2', 'ac3', 'ac4', 'ac5', 'ac6', 'ac7', 'ac8', 'ac9', 'ac10', 'ac11', 'ac12',
                                                        'su1', 'su2']]
        for col in columns:
            cur.execute(f'ALTER TABLE Database ADD COLUMN {col} TEXT')
        con.commit()
    
    if not 'SaveINFO' in tables:
        cur.execute('INSERT INTO SaveINFO (id, ParamDB, Next_number) VALUES (1, "1-2-3-4-5-6-7-8-9-10-11-12", 0)')
        con.commit()

    if not 'Contract' in tables:
        cur.execute(f'INSERT INTO Contract (id, Year, Month, Time, Count_tariff, Count_money) VALUES (1, {int(datetime.datetime.now().year)}, "Январь", 12, 0, 0)')
        con.commit()

    tables = [i[0] for i in list(cur.execute('SELECT name FROM sqlite_master WHERE type="table"'))]


file_db, cur, con, start = '', '', '', True


MONTH = ['январь', 'февраль', 'март', 'апрель',
         'май', 'июнь', 'июль', 'август',
         'сентябрь', 'октябрь', 'ноябрь', 'декабрь']
MONTH_SHORT = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
MAIN_NAME_COLUMNS = ['Number', 'Category', 'Comment', 'FIO', 'Job_title', 'Subdivision',
                     'Current_blocking', 'Date_current_blocking', 'TariffN',
                     'IP_address', 'Tariff', 'ICCID_number', 'Date_issue',
                     'Date_activation', 'Date_change', 'Discription_change', 'Reg_government_services']
OTHER_NAME_COLUMNS = ['tr1Jan', 'tr2Jan', 'tr3Jan', 'ac1Jan', 'ac2Jan', 'ac3Jan', 'ac4Jan', 'ac5Jan',
                      'ac6Jan', 'ac7Jan', 'ac8Jan', 'ac9Jan', 'ac10Jan', 'ac11Jan', 'ac12Jan', 'su1Jan',
                      'su2Jan', 'tr1Feb', 'tr2Feb', 'tr3Feb', 'ac1Feb', 'ac2Feb', 'ac3Feb', 'ac4Feb',
                      'ac5Feb', 'ac6Feb', 'ac7Feb', 'ac8Feb', 'ac9Feb', 'ac10Feb', 'ac11Feb', 'ac12Feb',
                      'su1Feb', 'su2Feb', 'tr1Mar', 'tr2Mar', 'tr3Mar', 'ac1Mar', 'ac2Mar', 'ac3Mar',
                      'ac4Mar', 'ac5Mar', 'ac6Mar', 'ac7Mar', 'ac8Mar', 'ac9Mar', 'ac10Mar', 'ac11Mar',
                      'ac12Mar', 'su1Mar', 'su2Mar', 'tr1Apr', 'tr2Apr', 'tr3Apr', 'ac1Apr', 'ac2Apr',
                      'ac3Apr', 'ac4Apr', 'ac5Apr', 'ac6Apr', 'ac7Apr', 'ac8Apr', 'ac9Apr', 'ac10Apr',
                      'ac11Apr', 'ac12Apr', 'su1Apr', 'su2Apr', 'tr1May', 'tr2May', 'tr3May', 'ac1May',
                      'ac2May', 'ac3May', 'ac4May', 'ac5May', 'ac6May', 'ac7May', 'ac8May', 'ac9May',
                      'ac10May', 'ac11May', 'ac12May', 'su1May', 'su2May', 'tr1Jun', 'tr2Jun', 'tr3Jun',
                      'ac1Jun', 'ac2Jun', 'ac3Jun', 'ac4Jun', 'ac5Jun', 'ac6Jun', 'ac7Jun', 'ac8Jun',
                      'ac9Jun', 'ac10Jun', 'ac11Jun', 'ac12Jun', 'su1Jun', 'su2Jun', 'tr1Jul', 'tr2Jul',
                      'tr3Jul', 'ac1Jul', 'ac2Jul', 'ac3Jul', 'ac4Jul', 'ac5Jul', 'ac6Jul', 'ac7Jul',
                      'ac8Jul', 'ac9Jul', 'ac10Jul', 'ac11Jul', 'ac12Jul', 'su1Jul', 'su2Jul', 'tr1Aug',
                      'tr2Aug', 'tr3Aug', 'ac1Aug', 'ac2Aug', 'ac3Aug', 'ac4Aug', 'ac5Aug', 'ac6Aug',
                      'ac7Aug', 'ac8Aug', 'ac9Aug', 'ac10Aug', 'ac11Aug', 'ac12Aug', 'su1Aug', 'su2Aug',
                      'tr1Sep', 'tr2Sep', 'tr3Sep', 'ac1Sep', 'ac2Sep', 'ac3Sep', 'ac4Sep', 'ac5Sep',
                      'ac6Sep', 'ac7Sep', 'ac8Sep', 'ac9Sep', 'ac10Sep', 'ac11Sep', 'ac12Sep', 'su1Sep',
                      'su2Sep', 'tr1Oct', 'tr2Oct', 'tr3Oct', 'ac1Oct', 'ac2Oct', 'ac3Oct', 'ac4Oct',
                      'ac5Oct', 'ac6Oct', 'ac7Oct', 'ac8Oct', 'ac9Oct', 'ac10Oct', 'ac11Oct', 'ac12Oct',
                      'su1Oct', 'su2Oct', 'tr1Nov', 'tr2Nov', 'tr3Nov', 'ac1Nov', 'ac2Nov', 'ac3Nov',
                      'ac4Nov', 'ac5Nov', 'ac6Nov', 'ac7Nov', 'ac8Nov', 'ac9Nov', 'ac10Nov', 'ac11Nov',
                      'ac12Nov', 'su1Nov', 'su2Nov', 'tr1Dec', 'tr2Dec', 'tr3Dec', 'ac1Dec', 'ac2Dec',
                      'ac3Dec', 'ac4Dec', 'ac5Dec', 'ac6Dec', 'ac7Dec', 'ac8Dec', 'ac9Dec', 'ac10Dec',
                      'ac11Dec', 'ac12Dec', 'su1Dec', 'su2Dec']
MAIN_HEADER_DB = ['Абонентский номер', 'Категория', 'Коментарий в ЛК МТС', 'ФИО', 'Должность', 'Подразделение',
             'Текущие блокировки', 'Даты текущих блокировок', 'Тариф №', 'IP-адрес', 'Тариф',
             'Серийный номер ICCID SIM-карты', 'Дата выдачи', 'Дата активации', 'Дата изменения',
             'Что поменялось', 'Регистрация на Госуслугах']
OTHER_HEADER_DB = [j + i for i in MONTH for j in ['Проверка коментария за ', 'GPRS МБ. за ', 'Сумма трафика за ',
                                                  'Общие за ', 'GPRS за ', 'SMS за ', 'MMS за ', 'МТС за ',
                                                  'М. моб. за ', 'М. фикс. за ', 'МГ за ', 'МН за ', 'ВСР за ',
                                                  'МНР за ', 'Прочее за ', 'Блокировки за ', 'Даты блокировок за ']]



class Menu(QMainWindow):
    def __init__(self):
        super().__init__()
        
        self.menuBar = QMenuBar(self)
        self.menuBar.setGeometry(0, 0, 1200, 30)
        
        bdSIMCard = self.menuBar.addMenu('База данных')

        import_data = QAction('&Импорт данных', self)
        bdSIMCard.addAction(import_data)
        import_data.triggered.connect(self.F_import_data)
        
        open_bd = QAction('&Открыть базу данных', self)
        bdSIMCard.addAction(open_bd)
        open_bd.triggered.connect(self.F_open_bd)

        import_traffic = QAction('&Импорт трафика', self)
        bdSIMCard.addAction(import_traffic)
        import_traffic.triggered.connect(self.F_import_traffic)

        import_accruals = QAction('&Импорт начислений', self)
        bdSIMCard.addAction(import_accruals)
        import_accruals.triggered.connect(self.F_import_accruals)

        import_subNumbers = QAction('&Импорт абон. номеров', self)
        bdSIMCard.addAction(import_subNumbers)
        import_subNumbers.triggered.connect(self.F_import_subNumbers)

        param_visual_bd = QAction('&Параметры отображения', self)
        bdSIMCard.addAction(param_visual_bd)
        param_visual_bd.triggered.connect(self.F_set_param_visual_bd)

        param_contract = QAction('&Условия договора', self)
        bdSIMCard.addAction(param_contract)
        param_contract.triggered.connect(self.F_set_param_contract)

        pivotTable = self.menuBar.addMenu('Сводная таблица')
        
        open_pt1 = QAction('&Открыть сводную таблицу по категориям', self)
        pivotTable.addAction(open_pt1)
        open_pt1.triggered.connect(self.F_open_pt1)

        open_pt2 = QAction('&Открыть сводную таблицу по тарифам', self)
        pivotTable.addAction(open_pt2)
        open_pt2.triggered.connect(self.F_open_pt2)

        calculate_costs = QAction('&Рассчитать затраты', self)
        pivotTable.addAction(calculate_costs)
        calculate_costs.triggered.connect(self.F_calculate_costs)

        graffic_prognos = QAction('&Прогнозный график', self)
        pivotTable.addAction(graffic_prognos)
        graffic_prognos.triggered.connect(self.F_graffic_prognos)

        deleteINFO = self.menuBar.addMenu('Удаление информации')
        
        delINFO = QAction('&Удаление информации', self)
        deleteINFO.addAction(delINFO)
        delINFO.triggered.connect(self.F_delINFO)

        saveINFO = self.menuBar.addMenu('Сохраненная информация')

        allPT = QAction('&Все сводные таблицы', self)
        saveINFO.addAction(allPT)
        allPT.triggered.connect(self.F_allPT)

        self.path = QLabel(self)
        self.path.setGeometry(300, 200, 550, 100)

        self.status = QLabel(self)
        self.status.setGeometry(100, 100, 500, 50)

    def clear_all(self):
        self.path.clear()
        self.status.clear()

    def create_windows(self):
        self.W_import_data = ImportData()
        self.W_open_bd = OpenBD()
        self.W_import_traffic = ImportTraffic()
        self.W_import_accruals = ImportAccruals()
        self.W_import_subNumbers = ImportSubNumbers()
        self.W_set_param_visual_bd = SetParamVisualBD()
        self.W_set_param_contract = SetParamContract()
        self.W_open_pt1 = OpenPT1()
        self.W_open_pt2 = OpenPT2()
        self.W_calculate_costs = CalculateCosts()
        self.W_graffic_prognos = GrafficPrognos()
        self.W_delINFO = DelINFO()
        self.W_allPT = AllPT()

    def F_import_data(self):
        global file_db, cur, con, start
        try:
            if start:
                file_db = self.filename
                if file_db == '':
                    raise Exception
                os.chdir('/'.join(file_db.split('/')[:-1]))
                con = sqlite3.connect(file_db.split('/')[-1])
                cur = con.cursor()
                start = False
            wind = self
            self.create_windows()
            create_table()
            self.clear_all()
            self.W_import_data.initUI()
            self.W_import_data.show()
            wind.close()
        except Exception:
            wind = QMessageBox(self)
            wind.setWindowTitle('Ошибка')
            wind.setText('Не выбран файл')
            wind.setIcon(QMessageBox.Critical)
            wind.setStandardButtons(QMessageBox.Close)
            res = wind.exec()
            

    def F_open_bd(self):
        global file_db, cur, con, start
        try:
            if start:
                file_db = self.filename
                if file_db == '':
                    raise Exception
                os.chdir('/'.join(file_db.split('/')[:-1]))
                con = sqlite3.connect(file_db.split('/')[-1])
                cur = con.cursor()
                start = False
            wind = self
            self.create_windows()
            create_table()
            self.clear_all()
            con.commit()
            self.W_open_bd.initUI()
            self.W_open_bd.show()
            wind.close()
        except Exception:
            wind = QMessageBox(self)
            wind.setWindowTitle('Ошибка')
            wind.setText('Не выбран файл')
            wind.setIcon(QMessageBox.Critical)
            wind.setStandardButtons(QMessageBox.Close)
            res = wind.exec()

    def F_import_traffic(self):
        global file_db, cur, con, start
        try:
            if start:
                file_db = self.filename
                if file_db == '':
                    raise Exception
                os.chdir('/'.join(file_db.split('/')[:-1]))
                con = sqlite3.connect(file_db.split('/')[-1])
                cur = con.cursor()
                start = False
            wind = self
            self.create_windows()
            create_table()
            self.clear_all()
            self.W_import_traffic.initUI()
            self.W_import_traffic.show()
            wind.close()
        except Exception:
            wind = QMessageBox(self)
            wind.setWindowTitle('Ошибка')
            wind.setText('Не выбран файл')
            wind.setIcon(QMessageBox.Critical)
            wind.setStandardButtons(QMessageBox.Close)
            res = wind.exec()

    def F_import_accruals(self):
        global file_db, cur, con, start
        try:
            if start:
                file_db = self.filename
                if file_db == '':
                    raise Exception
                os.chdir('/'.join(file_db.split('/')[:-1]))
                con = sqlite3.connect(file_db.split('/')[-1])
                cur = con.cursor()
                start = False
            wind = self
            self.create_windows()
            create_table()
            self.clear_all()
            self.W_import_accruals.initUI()
            self.W_import_accruals.show()
            wind.close()
        except Exception:
            wind = QMessageBox(self)
            wind.setWindowTitle('Ошибка')
            wind.setText('Не выбран файл')
            wind.setIcon(QMessageBox.Critical)
            wind.setStandardButtons(QMessageBox.Close)
            res = wind.exec()

    def F_import_subNumbers(self):
        global file_db, cur, con, start
        try:
            if start:
                file_db = self.filename
                if file_db == '':
                    raise Exception
                os.chdir('/'.join(file_db.split('/')[:-1]))
                con = sqlite3.connect(file_db.split('/')[-1])
                cur = con.cursor()
                start = False
            wind = self
            self.create_windows()
            create_table()
            self.clear_all()
            self.W_import_subNumbers.initUI()
            self.W_import_subNumbers.show()
            wind.close()
        except Exception:
            wind = QMessageBox(self)
            wind.setWindowTitle('Ошибка')
            wind.setText('Не выбран файл')
            wind.setIcon(QMessageBox.Critical)
            wind.setStandardButtons(QMessageBox.Close)
            res = wind.exec()

    def F_set_param_visual_bd(self):
        global file_db, cur, con, start
        try:
            if start:
                file_db = self.filename
                if file_db == '':
                    raise Exception
                os.chdir('/'.join(file_db.split('/')[:-1]))
                con = sqlite3.connect(file_db.split('/')[-1])
                cur = con.cursor()
                start = False
            wind = self
            self.create_windows()
            create_table()
            self.clear_all()
            self.W_set_param_visual_bd.initUI()
            self.W_set_param_visual_bd.show()
            wind.close()
        except Exception:
            wind = QMessageBox(self)
            wind.setWindowTitle('Ошибка')
            wind.setText('Не выбран файл')
            wind.setIcon(QMessageBox.Critical)
            wind.setStandardButtons(QMessageBox.Close)
            res = wind.exec()

    def F_set_param_contract(self):
        global file_db, cur, con, start
        try:
            if start:
                file_db = self.filename
                if file_db == '':
                    raise Exception
                os.chdir('/'.join(file_db.split('/')[:-1]))
                con = sqlite3.connect(file_db.split('/')[-1])
                cur = con.cursor()
                start = False
            wind = self
            self.create_windows()
            create_table()
            self.clear_all()
            self.W_set_param_contract.initUI()
            self.W_set_param_contract.show()
            wind.close()
        except Exception:
            wind = QMessageBox(self)
            wind.setWindowTitle('Ошибка')
            wind.setText('Не выбран файл')
            wind.setIcon(QMessageBox.Critical)
            wind.setStandardButtons(QMessageBox.Close)
            res = wind.exec()

    def F_open_pt1(self):
        global file_db, cur, con, start
        try:
            if start:
                file_db = self.filename
                if file_db == '':
                    raise Exception
                os.chdir('/'.join(file_db.split('/')[:-1]))
                con = sqlite3.connect(file_db.split('/')[-1])
                cur = con.cursor()
                start = False
            wind = self
            self.create_windows()
            create_table()
            self.clear_all()
            con.commit()
            self.W_open_pt1.initUI()
            self.W_open_pt1.show()
            wind.close()
        except Exception:
            wind = QMessageBox(self)
            wind.setWindowTitle('Ошибка')
            wind.setText('Не выбран файл')
            wind.setIcon(QMessageBox.Critical)
            wind.setStandardButtons(QMessageBox.Close)
            res = wind.exec()

    def F_open_pt2(self):
        global file_db, cur, con, start
        try:
            if start:
                file_db = self.filename
                if file_db == '':
                    raise Exception
                os.chdir('/'.join(file_db.split('/')[:-1]))
                con = sqlite3.connect(file_db.split('/')[-1])
                cur = con.cursor()
                start = False
            wind = self
            self.create_windows()
            create_table()
            self.clear_all()
            con.commit()
            self.W_open_pt2.initUI()
            self.W_open_pt2.show()
            wind.close()
        except Exception:
            wind = QMessageBox(self)
            wind.setWindowTitle('Ошибка')
            wind.setText('Не выбран файл')
            wind.setIcon(QMessageBox.Critical)
            wind.setStandardButtons(QMessageBox.Close)
            res = wind.exec()

    def F_calculate_costs(self):
        global file_db, cur, con, start
        try:
            if start:
                file_db = self.filename
                if file_db == '':
                    raise Exception
                os.chdir('/'.join(file_db.split('/')[:-1]))
                con = sqlite3.connect(file_db.split('/')[-1])
                cur = con.cursor()
                start = False
            wind = self
            self.create_windows()
            create_table()
            self.clear_all()
            self.W_calculate_costs.initUI()
            self.W_calculate_costs.show()
            wind.close()
        except Exception:
            wind = QMessageBox(self)
            wind.setWindowTitle('Ошибка')
            wind.setText('Не выбран файл')
            wind.setIcon(QMessageBox.Critical)
            wind.setStandardButtons(QMessageBox.Close)
            res = wind.exec()

    def F_graffic_prognos(self):
        global file_db, cur, con,start
        try:
            if start:
                file_db = self.filename
                if file_db == '':
                    raise Exception
                os.chdir('/'.join(file_db.split('/')[:-1]))
                con = sqlite3.connect(file_db.split('/')[-1])
                cur = con.cursor()
                start = False
            wind = self
            self.create_windows()
            create_table()
            self.clear_all()
            self.W_graffic_prognos.initUI()
            self.W_graffic_prognos.show()
            wind.close()
        except Exception:
            wind = QMessageBox(self)
            wind.setWindowTitle('Ошибка')
            wind.setText('Не выбран файл')
            wind.setIcon(QMessageBox.Critical)
            wind.setStandardButtons(QMessageBox.Close)
            res = wind.exec()
            
    def F_delINFO(self):
        global file_db, cur, con, start
        try:
            if start:
                file_db = self.filename
                if file_db == '':
                    raise Exception
                os.chdir('/'.join(file_db.split('/')[:-1]))
                con = sqlite3.connect(file_db.split('/')[-1])
                cur = con.cursor()
                start = False
            wind = self
            self.create_windows()
            create_table()
            self.clear_all()
            self.W_delINFO.initUI()
            self.W_delINFO.show()
            wind.close()
        except Exception:
            wind = QMessageBox(self)
            wind.setWindowTitle('Ошибка')
            wind.setText('Не выбран файл')
            wind.setIcon(QMessageBox.Critical)
            wind.setStandardButtons(QMessageBox.Close)
            res = wind.exec()


    def F_allPT(self):
        global file_db, cur, con, start
        try:
            if start:
                file_db = self.filename
                if file_db == '':
                    raise Exception
                os.chdir('/'.join(file_db.split('/')[:-1]))
                con = sqlite3.connect(file_db.split('/')[-1])
                cur = con.cursor()
                start = False
            wind = self
            self.create_windows()
            create_table()
            self.clear_all()
            self.W_allPT.initUI()
            self.W_allPT.show()
            wind.close()
        except Exception:
            wind = QMessageBox(self)
            wind.setWindowTitle('Ошибка')
            wind.setText('Не выбран файл')
            wind.setIcon(QMessageBox.Critical)
            wind.setStandardButtons(QMessageBox.Close)
            res = wind.exec()

class DatabaseSIMCard(Menu):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setGeometry(100, 100, 1200, 600)
        self.setWindowTitle('База данных SIM-карт')

        self.filename = ''

        if self.filename:
            self.path.setText(f'Путь к выбранному файлу:\n{self.filename}')
        else:
            self.path.setText('')

        btn = QPushButton('Выбрать файл базы данных', self)
        btn.setGeometry(300, 100, 250, 40)
        btn.clicked.connect(self.F_choose_file)

    def F_choose_file(self):
        options = QFileDialog.Options()
        filename, _ = QFileDialog.getOpenFileName(self, 'Выбрать файл', '', 'Все файлы (*.db)', options=options)

        if filename:
            self.filename = filename
            self.path.setText(f'Путь к выбранному файлу:\n{self.filename}')


class ImportData(Menu):
    def __init__(self):
        super().__init__()

    def initUI(self):
        self.setGeometry(100, 100, 1200, 600)
        self.setWindowTitle('Импорт данных')

        self.filename = ''

        self.label1 = QLabel(self)
        self.label1.setGeometry(50, 60, 550, 40)
        self.label1.setText('Введите название файла, из котрого хотите перенести данные в базу данных')

        self.import_button = QPushButton('Импортировать', self)
        self.import_button.setGeometry(350, 120, 150, 40)
        self.import_button.clicked.connect(self.F_import)

        self.choose_file = QPushButton('Выбрать файл', self)
        self.choose_file.setGeometry(50, 120, 250, 40)
        self.choose_file.clicked.connect(self.F_choose_file)

        self.path.setGeometry(50, 200, 700, 100)
        if self.filename:
            self.path.setText(f'Путь к выбранному файлу:\n{self.filename}')
        else:
            self.path.setText('')

    def F_choose_file(self):
        options = QFileDialog.Options()
        filename, _ = QFileDialog.getOpenFileName(self, 'Выбрать файл', '', 'Все файлы (*.xls)', options=options)

        if filename:
            self.filename = filename
            self.path.setText(f'Путь к выбранному файлу:\n{self.filename}')

    def F_import(self):
        if os.path.exists(self.filename):
            if self.filename.split('.')[-1] in ['xls']:
                
                workbook = xlrd.open_workbook(self.filename)
                sheet = workbook.sheet_by_index(0)
                n_rows = sheet.nrows
                n_cols = sheet.ncols

                for i in range(1, n_rows):
                    vals = []
                    for j in range(1, n_cols):
                        value = sheet.cell_value(i, j)
                        if str(value) and j in [8, 13, 14, 15]:
                            try:
                                value = xlrd.xldate_as_datetime(value, 0).date()
                                value = '.'.join([str(value.day).zfill(2), str(value.month).zfill(2), str(value.year)])
                            except Exception:
                                pass
                        vals.append(str(value))
                    vals_columns = ','.join(MAIN_NAME_COLUMNS)
                    sql_command = f'INSERT INTO Database ({vals_columns}) VALUES ({",".join(["?" for i in vals])})'
                    
                    try:
                        vals[0] = int(float(vals[0]))
                    except Exception:
                        pass
                    try:
                        vals[8] = int(float(vals[8]))
                    except Exception:
                        pass
                    cur.execute(sql_command, vals)
                con.commit()
                wind = QMessageBox(self)
                wind.setWindowTitle('Процесс успешно завершен')
                wind.setText('Данные успешно импортированы')
                wind.setIcon(QMessageBox.Information)
                wind.setStandardButtons(QMessageBox.Close)
                res = wind.exec()
            else:
                wind = QMessageBox(self)
                wind.setWindowTitle('Ошибка')
                wind.setText('Файл не является документом Excel, нужно выбрать файл Excel')
                wind.setIcon(QMessageBox.Critical)
                wind.setStandardButtons(QMessageBox.Close)
                res = wind.exec()
        else:
            wind = QMessageBox(self)
            wind.setWindowTitle('Ошибка')
            wind.setText('Не найден файл с таким именем')
            wind.setIcon(QMessageBox.Critical)
            wind.setStandardButtons(QMessageBox.Close)
            res = wind.exec()


class OpenBD(Menu):
    def __init__(self):
        super().__init__()

    def initUI(self):
        self.setGeometry(100, 100, 1200, 600)
        self.setWindowTitle('База данных')
        
        request = list(cur.execute('''SELECT * FROM Database'''))

        param_contract = list(cur.execute('SELECT Year, Month, Time FROM Contract WHERE id = 1'))

        if param_contract:
            param_contract = list(param_contract[0])
            param_contract[1] = param_contract[1].lower()

            id_month = MONTH.index(param_contract[1])
        else:
            id_month = 0
        
        if request:

            button = QPushButton('Удалить', self)
            button.setGeometry(1000, 400, 100, 40)
            button.clicked.connect(self.delete)

            need_id = [(int(i) - id_month - 1) % 12 for i in list(cur.execute('SELECT ParamDB FROM SaveINFO'))[0][0].split('-') if i]
            need_id2 = [int(i) for i in list(cur.execute('SELECT ParamDB FROM SaveINFO'))[0][0].split('-') if i]
            if need_id:
                ids = [i for i in range(17)] + [17 + j for i in need_id for j in range(17 * int(i), 17 * (int(i) + 1))]
            else:
                ids = [i for i in range(17)]


            other_header_db = OTHER_HEADER_DB[17 * id_month:] + OTHER_HEADER_DB[:17 * id_month]

            other_name_columns = OTHER_NAME_COLUMNS[17 * id_month:] + OTHER_NAME_COLUMNS[:17 * id_month]


            self.need_header = [i for ind, i in enumerate(MAIN_HEADER_DB + other_header_db) if ind in ids]
            self.need_columns = [i for ind, i in enumerate(MAIN_NAME_COLUMNS + other_name_columns) if ind in ids]
            for_req = ', '.join(self.need_columns)

            request = list(cur.execute(f'''SELECT {for_req} FROM Database'''))


            self.table = QTableWidget(self)
            self.table.setGeometry(50, 50, 900, 400)
            self.table.setColumnCount(len(self.need_header))
            self.table.setRowCount(len(request) + 100)
            self.table.setHorizontalHeaderLabels(self.need_header)


            for i, elem in enumerate(request):
                for j, val in enumerate(elem):
                    if not val is None:

                        self.table.setItem(i, j, QTableWidgetItem(str(val)))
                    else:
                        self.table.setItem(i, j, QTableWidgetItem(''))

            self.table.resizeColumnsToContents()
            self.table.cellChanged.connect(self.change_add)
            self.table.horizontalHeader().sectionClicked.connect(self.filter_table)

            
            try:
                self.table.setCurrentCell(self.row, self.col)
            except Exception:
                self.table.setCurrentCell(0, 0)

            self.params = {}

            for i in range(self.table.rowCount()):
                for j in range(self.table.columnCount()):
                    item = self.table.item(i, j)
                    if not item is None:
                        val = item.text()
                    else:
                        val = 'None'
                    if not j in self.params.keys():
                        self.params[j] = [val]
                    if not val in self.params[j]:
                        self.params[j] = self.params[j] + [val]
        else:
            self.status.setText('База данных пустая')


    def filter_table(self, index):
        self.index_col = index
        all_variants = []
        self.menu = QMenu(self)

        all_header = MAIN_HEADER_DB + OTHER_HEADER_DB
        ind = all_header.index(self.need_header[index])
        data = [i[ind] for i in list(cur.execute('SELECT * FROM Database'))]
        self.check_boxes = []
        for i in ['Выбрать все'] + data:
            if not i in all_variants:
                all_variants.append(i)
                checkBox = QCheckBox(i, self.menu)

                checkableAction = QWidgetAction(self.menu)
                checkableAction.setDefaultWidget(checkBox)
                self.menu.addAction(checkableAction)

                self.check_boxes.append(checkBox)

                if i == 'Выбрать все':
                    checkBox.stateChanged.connect(self.select_all)

        s = []
        for i in range(self.table.rowCount()):
            if not self.table.isRowHidden(i):
                item = self.table.item(i, self.index_col)
                if (not item is None):
                    s.append(item.text())

        for i in self.check_boxes:
            if i.text() in s:
                i.setChecked(True)
            else:
                i.setChecked(False)

        
        btn_cancel = QPushButton('Отменить', self.menu)
        btn_cancel.clicked.connect(self.menu.close)
        checkableAction = QWidgetAction(self.menu)
        checkableAction.setDefaultWidget(btn_cancel)
        self.menu.addAction(checkableAction)
        
        btn_ok = QPushButton('OK', self.menu)
        btn_ok.clicked.connect(self.save)
        checkableAction = QWidgetAction(self.menu)
        checkableAction.setDefaultWidget(btn_ok)
        self.menu.addAction(checkableAction)

        x = self.geometry().left()
        y = self.geometry().top()

        self.menu.exec(QCursor.pos())


    def select_all(self, state):
        for i in self.check_boxes:
            i.setChecked(True)

        
    def save(self):

        s = []
        for i in self.check_boxes:
            if i.isChecked() and i.text() != 'Выбрать все':
                s.append(i.text())

        self.params[self.index_col] = s

        for i in range(self.table.rowCount()):
            flag = True
            for j in range(self.table.columnCount()):
                item = self.table.item(i, j)
                if item is None:
                    item = 'None'
                else:
                    item = item.text()
                if not item in self.params[j]:
                    flag = False
                    break
            if flag:
                self.table.setRowHidden(i, False)
            else:
                self.table.setRowHidden(i, True)
                
                
        self.menu.close()

    def change_add(self, row, column):
        row_count = len(list(cur.execute('SELECT * FROM Database')))
        if row < row_count:

            val_column = self.need_columns[column]
            new_val = self.table.item(row, column).text()
            number = self.table.item(row, 0).text()
            if str(number) == 'None' or not str(number):
                return 0
            req = val_column + ' = ' + new_val
            sql_command = f'''UPDATE Database SET {f'{val_column} = "{new_val}"'} WHERE Number = ?'''
            cur.execute(sql_command, (number,))
        else:

            val_column = self.need_columns[column]
            new_val = self.table.item(row, column).text()
            number = self.table.item(row, 0)
            if str(number) == 'None' or not str(number):
                return 0
            if column == 0:
                sql_command = f'''INSERT INTO Database (Number) VALUES ({f'"{new_val}"'})'''
            else:
                sql_command = f'''INSERT INTO Database ({val_column}) VALUES ({f'"{new_val}"'})'''
            cur.execute(sql_command)
        self.row = row
        self.col = column
        con.commit()
        self.initUI()
        self.show()

    def delete(self):
        rows = self.table.selectionModel().selectedRows()
        request = list(cur.execute('''SELECT * FROM Database'''))

        for i in rows:
            number = request[i.row()][0]
            if number:
                cur.execute('DELETE FROM Database WHERE number = ?', (number,))
            else:
                pass
        if rows[0]:
            self.row = rows[0].row()
            self.col = 0
        self.close()
        con.commit()
        self.initUI()
        self.show()
        


class ImportTraffic(Menu):
    def __init__(self):
        super().__init__()

    def initUI(self):
        self.setGeometry(100, 100, 1200, 600)
        self.setWindowTitle('Импорт трафика')

        self.num_notFind = []

        self.filename = ''

        self.label1 = QLabel(self)
        self.label1.setGeometry(50, 60, 550, 40)
        self.label1.setText('Выберете файл, из котрого хотите импортировать данные о трафике')

        self.choose_file = QPushButton('Выбрать файл', self)
        self.choose_file.setGeometry(50, 120, 250, 40)
        self.choose_file.clicked.connect(self.F_choose_file)

        self.path.setGeometry(50, 200, 700, 100)
        if self.filename:
            self.path.setText(f'Путь к выбранному файлу:\n{self.filename}')
        else:
            self.path.setText('')

        self.import_button = QPushButton('Импортировать', self)
        self.import_button.setGeometry(350, 120, 150, 40)
        self.import_button.clicked.connect(self.F_import)

        self.label2 = QLabel(self)
        self.label2.setGeometry(800, 60, 550, 40)
        self.label2.setText('Выберите месяц, которому соответствуют эти данные')

        self.month1 = QRadioButton('Январь', self)
        self.month1.setGeometry(800, 100, 150, 30)
        self.month2 = QRadioButton('Февраль', self)
        self.month2.setGeometry(800, 130, 150, 30)
        self.month3 = QRadioButton('Март', self)
        self.month3.setGeometry(800, 160, 150, 30)
        self.month4 = QRadioButton('Апрель', self)
        self.month4.setGeometry(800, 190, 150, 30)
        self.month5 = QRadioButton('Май', self)
        self.month5.setGeometry(800, 220, 150, 30)
        self.month6 = QRadioButton('Июнь', self)
        self.month6.setGeometry(800, 250, 150, 30)
        self.month7 = QRadioButton('Июль', self)
        self.month7.setGeometry(800, 280, 150, 30)
        self.month8 = QRadioButton('Август', self)
        self.month8.setGeometry(800, 310, 150, 30)
        self.month9 = QRadioButton('Сентябрь', self)
        self.month9.setGeometry(800, 340, 150, 30)
        self.month10 = QRadioButton('Октябрь', self)
        self.month10.setGeometry(800, 370, 150, 30)
        self.month11 = QRadioButton('Ноябрь', self)
        self.month11.setGeometry(800, 400, 150, 30)
        self.month12 = QRadioButton('Декабрь', self)
        self.month12.setGeometry(800, 430, 150, 30)
            

    def F_choose_file(self):
        options = QFileDialog.Options()
        filename, _ = QFileDialog.getOpenFileName(self, 'Выбрать файл', '', 'Все файлы (*.xls)', options=options)

        if filename:
            self.filename = filename
            self.path.setText(f'Путь к выбранному файлу:\n{self.filename}')

    def F_import(self):
        current_month = -1
        for ind, i in enumerate([self.month1, self.month2, self.month3, self.month4,
                                 self.month5, self.month6, self.month7, self.month8,
                                 self.month9, self.month10, self.month11, self.month12]):
            if i.isChecked():
                current_month = ind
        if os.path.exists(self.filename):
            if self.filename.split('/')[-1].split('.')[-1] in ['xls']:
                if [1 for i in [self.month1, self.month2, self.month3, self.month4,
                                self.month5, self.month6, self.month7, self.month8,
                                self.month9, self.month10, self.month11, self.month12] if i.isChecked()]:

                    workbook = xlrd.open_workbook(self.filename)
                    sheet = workbook.sheet_by_index(0)
                    n_rows = sheet.nrows

                    need_vals_columns = [OTHER_NAME_COLUMNS[i] for i in range(17 * int(current_month), 17 * int(current_month) + 3)]

                    vals_columns = [i + MONTH_SHORT[current_month] for i in ['tr1', 'tr2', 'tr3']]
                    if [j for i in list(cur.execute(f'SELECT {", ".join(vals_columns)} FROM Database')) for j in list(i) if not j is None and j != '']:
                        wind = QMessageBox(self)
                        wind.setWindowTitle('Ошибка')
                        wind.setText('Возможная ошибка: перезаполнение данных. Вам следует удалить информацию по месяцам из базы данных')
                        wind.setIcon(QMessageBox.Critical)
                        wind.setStandardButtons(QMessageBox.Close)
                        res = wind.exec()
                        return 0

                    for i in range(1, n_rows):
                        number = str(sheet.cell_value(i, 0))
                        if number:
                            try:
                                number = int(float(number))
                            except:
                                pass
                            vals = []
                            summa = 0
                            for j in [1, 2]:
                                value = sheet.cell_value(i, j)
                                vals.append(value)
                            for j in [5, 6, 7, 8, 9, 10, 11]:
                                summa += int(sheet.cell_value(i, j))
                            if list(cur.execute('SELECT Comment FROM Database WHERE Number = ?', (number,))) and str(list(cur.execute('SELECT Comment FROM Database WHERE Number = ?', (number,)))[0][0]) == str(vals[0]):
                                res = '✅'
                            else:
                                res = '❌'
                            vals = [res] + [str(vals[1])] + [summa]

                            req = ', '.join([f'{i} = "{j}"' for i, j in zip(vals_columns, vals)])
                            
                            sql_command = f'UPDATE Database SET {req} WHERE Number = ?'
                            cur.execute(sql_command, (number,))
                    
                    con.commit()

                    wind = QMessageBox(self)
                    wind.setWindowTitle('Процесс успешно заврешен')
                    wind.setText('Данные успешно импортированы')
                    wind.setIcon(QMessageBox.Information)
                    wind.setStandardButtons(QMessageBox.Close)
                    res = wind.exec()
                else:
                    wind = QMessageBox(self)
                    wind.setWindowTitle('Ошибка')
                    wind.setText('Не выбран месяц, за который импортируются данные')
                    wind.setIcon(QMessageBox.Critical)
                    wind.setStandardButtons(QMessageBox.Close)
                    res = wind.exec()
            else:
                wind = QMessageBox(self)
                wind.setWindowTitle('Ошибка')
                wind.setText('Файл не является документом Excel, нужно выбрать файл Excel')
                wind.setIcon(QMessageBox.Critical)
                wind.setStandardButtons(QMessageBox.Close)
                res = wind.exec()
        else:
            wind = QMessageBox(self)
            wind.setWindowTitle('Ошибка')
            wind.setText('Не найден файл с таким именем')
            wind.setIcon(QMessageBox.Critical)
            wind.setStandardButtons(QMessageBox.Close)
            res = wind.exec()


class ImportAccruals(Menu):
    def __init__(self):
        super().__init__()

    def initUI(self):
        self.setGeometry(100, 100, 1200, 600)
        self.setWindowTitle('Импорт начислений')

        self.num_notFind = []

        self.filename = ''

        self.label1 = QLabel(self)
        self.label1.setGeometry(50, 60, 550, 40)
        self.label1.setText('Выберете файл, из котрого хотите импортировать данные о начислениях')

        self.choose_file = QPushButton('Выбрать файл', self)
        self.choose_file.setGeometry(50, 120, 250, 40)
        self.choose_file.clicked.connect(self.F_choose_file)

        self.path.setGeometry(50, 200, 700, 100)
        if self.filename:
            self.path.setText(f'Путь к выбранному файлу:\n{self.filename}')
        else:
            self.path.clear()

        self.import_button = QPushButton('Импортировать', self)
        self.import_button.setGeometry(350, 120, 150, 40)
        self.import_button.clicked.connect(self.F_import)

        self.label2 = QLabel(self)
        self.label2.setGeometry(800, 60, 550, 40)
        self.label2.setText('Выберите месяц, которому соответствуют эти данные')

        self.month1 = QRadioButton('Январь', self)
        self.month1.setGeometry(800, 100, 150, 30)
        self.month2 = QRadioButton('Февраль', self)
        self.month2.setGeometry(800, 130, 150, 30)
        self.month3 = QRadioButton('Март', self)
        self.month3.setGeometry(800, 160, 150, 30)
        self.month4 = QRadioButton('Апрель', self)
        self.month4.setGeometry(800, 190, 150, 30)
        self.month5 = QRadioButton('Май', self)
        self.month5.setGeometry(800, 220, 150, 30)
        self.month6 = QRadioButton('Июнь', self)
        self.month6.setGeometry(800, 250, 150, 30)
        self.month7 = QRadioButton('Июль', self)
        self.month7.setGeometry(800, 280, 150, 30)
        self.month8 = QRadioButton('Август', self)
        self.month8.setGeometry(800, 310, 150, 30)
        self.month9 = QRadioButton('Сентябрь', self)
        self.month9.setGeometry(800, 340, 150, 30)
        self.month10 = QRadioButton('Октябрь', self)
        self.month10.setGeometry(800, 370, 150, 30)
        self.month11 = QRadioButton('Ноябрь', self)
        self.month11.setGeometry(800, 400, 150, 30)
        self.month12 = QRadioButton('Декабрь', self)
        self.month12.setGeometry(800, 430, 150, 30)

    def F_choose_file(self):
        options = QFileDialog.Options()
        filename, _ = QFileDialog.getOpenFileName(self, 'Выбрать файл', '', 'Все файлы (*.xls)', options=options)

        if filename:
            self.filename = filename
            self.path.setText(f'Путь к выбранному файлу:\n{self.filename}')

    def F_import(self):
        current_month = 0
        for ind, i in enumerate([self.month1, self.month2, self.month3, self.month4,
                                 self.month5, self.month6, self.month7, self.month8,
                                 self.month9, self.month10, self.month11, self.month12]):
            if i.isChecked():
                current_month = ind
        if os.path.exists(self.filename):
            if self.filename.split('/')[-1].split('.')[-1] in ['xls']:
                if [1 for i in [self.month1, self.month2, self.month3, self.month4,
                                self.month5, self.month6, self.month7, self.month8,
                                self.month9, self.month10, self.month11, self.month12] if i.isChecked()]:

                    workbook = xlrd.open_workbook(self.filename)
                    sheet = workbook.sheet_by_index(0)
                    n_rows = sheet.nrows

                    need_vals_columns = [OTHER_NAME_COLUMNS[i] for i in range(17 * int(current_month) + 3, 17 * int(current_month) + 15)]

                    vals_columns = [i + MONTH_SHORT[current_month] for i in ['ac1', 'ac2', 'ac3', 'ac4', 'ac5', 'ac6',
                                                                             'ac7', 'ac8', 'ac9', 'ac10', 'ac11', 'ac12']]
                    if [j for i in list(cur.execute(f'SELECT {", ".join(vals_columns)} FROM Database')) for j in list(i) if not j is None and j != '']:
                        wind = QMessageBox(self)
                        wind.setWindowTitle('Ошибка')
                        wind.setText('Возможная ошибка: перезаполнение данных. Вам следует удалить информацию по месяцам из базы данных')
                        wind.setIcon(QMessageBox.Critical)
                        wind.setStandardButtons(QMessageBox.Close)
                        res = wind.exec()
                        return 0
                    
                    for i in range(1, n_rows):
                        number = str(sheet.cell_value(i, 0))
                        if number:
                            try:
                                number = int(float(number))
                            except Exception:
                                pass
                            vals = []
                            for j in range(2, 14):
                                value = sheet.cell_value(i, j)
                                vals.append(str(value))
                            req = ', '.join([f'{i} = "{j}"' for i, j in zip(vals_columns, vals)])
                            sql_command = f'UPDATE Database SET {req} WHERE Number = ?'
                            cur.execute(sql_command, (number,))
                    con.commit()
                    
                    wind = QMessageBox(self)
                    wind.setWindowTitle('Процесс успешно завершен')
                    wind.setText('Данные успешно импортированы')
                    wind.setIcon(QMessageBox.Information)
                    wind.setStandardButtons(QMessageBox.Close)
                    res = wind.exec()
                else:
                    wind = QMessageBox(self)
                    wind.setWindowTitle('Ошибка')
                    wind.setText('Не выбран месяц, за который импортируются данные')
                    wind.setIcon(QMessageBox.Critical)
                    wind.setStandardButtons(QMessageBox.Close)
                    res = wind.exec()
            else:
                wind = QMessageBox(self)
                wind.setWindowTitle('Ошибка')
                wind.setText('Файл не является документом Excel, нужно выбрать файл Excel')
                wind.setIcon(QMessageBox.Critical)
                wind.setStandardButtons(QMessageBox.Close)
                res = wind.exec()
        else:
            wind = QMessageBox(self)
            wind.setWindowTitle('Ошибка')
            wind.setText('Не найден файл с таким именем')
            wind.setIcon(QMessageBox.Critical)
            wind.setStandardButtons(QMessageBox.Close)
            res = wind.exec()


class ImportSubNumbers(Menu):
    def __init__(self):
        super().__init__()

    def initUI(self):
        self.setGeometry(100, 100, 1200, 600)
        self.setWindowTitle('Импорт абонентских номеров')

        self.num_notFind = []

        self.filename = ''

        self.label1 = QLabel(self)
        self.label1.setGeometry(50, 60, 550, 40)
        self.label1.setText('Выберете файл, из котрого хотите импортировать данные о начислениях')

        self.choose_file = QPushButton('Выбрать файл', self)
        self.choose_file.setGeometry(50, 120, 250, 40)
        self.choose_file.clicked.connect(self.F_choose_file)

        self.path.setGeometry(50, 200, 700, 100)
        if self.filename:
            self.path.setText(f'Путь к выбранному файлу:\n{self.filename}')
        else:
            self.path.clear()

        self.import_button = QPushButton('Импортировать', self)
        self.import_button.setGeometry(350, 120, 150, 40)
        self.import_button.clicked.connect(self.F_import)

        self.label2 = QLabel(self)
        self.label2.setGeometry(800, 60, 550, 40)
        self.label2.setText('Выберите месяц, которому соответствуют эти данные')

        self.month1 = QRadioButton('Январь', self)
        self.month1.setGeometry(800, 100, 150, 30)
        self.month2 = QRadioButton('Февраль', self)
        self.month2.setGeometry(800, 130, 150, 30)
        self.month3 = QRadioButton('Март', self)
        self.month3.setGeometry(800, 160, 150, 30)
        self.month4 = QRadioButton('Апрель', self)
        self.month4.setGeometry(800, 190, 150, 30)
        self.month5 = QRadioButton('Май', self)
        self.month5.setGeometry(800, 220, 150, 30)
        self.month6 = QRadioButton('Июнь', self)
        self.month6.setGeometry(800, 250, 150, 30)
        self.month7 = QRadioButton('Июль', self)
        self.month7.setGeometry(800, 280, 150, 30)
        self.month8 = QRadioButton('Август', self)
        self.month8.setGeometry(800, 310, 150, 30)
        self.month9 = QRadioButton('Сентябрь', self)
        self.month9.setGeometry(800, 340, 150, 30)
        self.month10 = QRadioButton('Октябрь', self)
        self.month10.setGeometry(800, 370, 150, 30)
        self.month11 = QRadioButton('Ноябрь', self)
        self.month11.setGeometry(800, 400, 150, 30)
        self.month12 = QRadioButton('Декабрь', self)
        self.month12.setGeometry(800, 430, 150, 30)

    def F_choose_file(self):
        options = QFileDialog.Options()
        filename, _ = QFileDialog.getOpenFileName(self, 'Выбрать файл', '', 'Все файлы (*.xls)', options=options)

        if filename:
            self.filename = filename
            self.path.setText(f'Путь к выбранному файлу:\n{self.filename}')

    def F_import(self):
        current_month = 0
        for ind, i in enumerate([self.month1, self.month2, self.month3, self.month4,
                                 self.month5, self.month6, self.month7, self.month8,
                                 self.month9, self.month10, self.month11, self.month12]):
            if i.isChecked():
                current_month = ind
        if os.path.exists(self.filename):
            if self.filename.split('/')[-1].split('.')[-1] in ['xls']:
                if [1 for i in [self.month1, self.month2, self.month3, self.month4,
                                self.month5, self.month6, self.month7, self.month8,
                                self.month9, self.month10, self.month11, self.month12] if i.isChecked()]:

                    workbook = xlrd.open_workbook(self.filename)
                    sheet = workbook.sheet_by_index(0)
                    n_rows = sheet.nrows

                    need_vals_columns = [OTHER_NAME_COLUMNS[i] for i in range(17 * int(current_month) + 15, 17 * int(current_month) + 17)]

                    vals_columns = [i + MONTH_SHORT[current_month] for i in ['su1', 'su2']]
                    if [j for i in list(cur.execute(f'SELECT {", ".join(vals_columns)} FROM Database')) for j in list(i) if not j is None and j != '']:
                        wind = QMessageBox(self)
                        wind.setWindowTitle('Ошибка')
                        wind.setText('Возможная ошибка: перезаполнение данных. Вам следует удалить информацию по месяцам из базы данных')
                        wind.setIcon(QMessageBox.Critical)
                        wind.setStandardButtons(QMessageBox.Close)
                        res = wind.exec()
                        return 0
                    
                    for i in range(1, n_rows):
                        number = str(sheet.cell_value(i, 0))
                        if number:
                            try:
                                number = int(float(number))
                            except Exception:
                                pass
                            vals = []
                            for j in [5, 6]:
                                value = sheet.cell_value(i, j)
                                vals.append(value)
                            vals = [str(i) for i in vals]
                            indexes = [ind for ind, i in enumerate(vals) if i != '']

                            vals_columns = [i for ind, i in enumerate(need_vals_columns) if ind in indexes]
                            values = [i for i in vals if i != '']
                            req = ', '.join([f'{i} = "{j}"' for i, j in zip(vals_columns, values)])
                            sql_command = f'UPDATE Database SET {req} WHERE Number = ?'
                            cur.execute(sql_command, (number,))
                    con.commit()
                    
                    wind = QMessageBox(self)
                    wind.setWindowTitle('Процесс успешно завершен')
                    wind.setText('Данные успешно импортированы')
                    wind.setIcon(QMessageBox.Information)
                    wind.setStandardButtons(QMessageBox.Close)
                    res = wind.exec()
                else:
                    wind = QMessageBox(self)
                    wind.setWindowTitle('Ошибка')
                    wind.setText('Не выбран месяц, за который импортируются данные')
                    wind.setIcon(QMessageBox.Critical)
                    wind.setStandardButtons(QMessageBox.Close)
                    res = wind.exec()
            else:
                wind = QMessageBox(self)
                wind.setWindowTitle('Ошибка')
                wind.setText('Файл не является документом Excel, нужно выбрать файл Excel')
                wind.setIcon(QMessageBox.Critical)
                wind.setStandardButtons(QMessageBox.Close)
                res = wind.exec()
        else:
            wind = QMessageBox(self)
            wind.setWindowTitle('Ошибка')
            wind.setText('Не найден файл с таким именем')
            wind.setIcon(QMessageBox.Critical)
            wind.setStandardButtons(QMessageBox.Close)
            res = wind.exec()


class SetParamVisualBD(Menu):
    def __init__(self):
        super().__init__()

    def initUI(self):
        self.setGeometry(100, 100, 1200, 600)
        self.setWindowTitle('Параметры отображения')

        self.label = QLabel(self)
        self.label.setGeometry(100, 60, 800, 40)
        self.label.setText('Выберите месяца, за которые будет отображаться информация о трафике, о начислениях и об абонентских номерах')

        self.month1 = QCheckBox('Январь', self)
        self.month1.setGeometry(100, 100, 150, 30)
        self.month2 = QCheckBox('Февраль', self)
        self.month2.setGeometry(100, 130, 150, 30)
        self.month3 = QCheckBox('Март', self)
        self.month3.setGeometry(100, 160, 150, 30)
        self.month4 = QCheckBox('Апрель', self)
        self.month4.setGeometry(100, 190, 150, 30)
        self.month5 = QCheckBox('Май', self)
        self.month5.setGeometry(100, 220, 150, 30)
        self.month6 = QCheckBox('Июнь', self)
        self.month6.setGeometry(100, 250, 150, 30)
        self.month7 = QCheckBox('Июль', self)
        self.month7.setGeometry(100, 280, 150, 30)
        self.month8 = QCheckBox('Август', self)
        self.month8.setGeometry(100, 310, 150, 30)
        self.month9 = QCheckBox('Сентябрь', self)
        self.month9.setGeometry(100, 340, 150, 30)
        self.month10 = QCheckBox('Октябрь', self)
        self.month10.setGeometry(100, 370, 150, 30)
        self.month11 = QCheckBox('Ноябрь', self)
        self.month11.setGeometry(100, 400, 150, 30)
        self.month12 = QCheckBox('Декабрь', self)
        self.month12.setGeometry(100, 430, 150, 30)

        need_id = [i for i in list(cur.execute('SELECT ParamDB FROM SaveINFO'))[0][0].split('-') if i]
        q_check_boxes = [self.month1, self.month2, self.month3, self.month4, self.month5, self.month6,
                         self.month7, self.month8, self.month9, self.month10, self.month11, self.month12]
        for i in need_id:
            q_check_boxes[int(i) - 1].setChecked(True)

        self.save_button = QPushButton('Сохранить изменения', self)
        self.save_button.setGeometry(100, 470, 170, 40)
        self.save_button.clicked.connect(self.F_save)

    def F_save(self):
        all_month = []
        for ind, i in enumerate([self.month1, self.month2, self.month3, self.month4,
                                 self.month5, self.month6, self.month7, self.month8,
                                 self.month9, self.month10, self.month11, self.month12]):
            if i.isChecked():
                all_month.append(ind + 1)
        for_req = '-'.join(list(map(str, all_month)))
        sql_command = f'''UPDATE SaveINFO SET ParamDB = "{for_req}" WHERE id = 1'''
        cur.execute(sql_command)
        con.commit()
        wind = QMessageBox(self)
        wind.setWindowTitle('Процесс успешно завершен')
        wind.setText('Данные успешно сохранены')
        wind.setIcon(QMessageBox.Information)
        wind.setStandardButtons(QMessageBox.Close)
        res = wind.exec()


class SetParamContract(Menu):
    def __init__(self):
        super().__init__()

    def initUI(self):
        self.setGeometry(100, 100, 1200, 600)
        self.setWindowTitle('Условия договора')

        params = list(cur.execute('SELECT Year, Month, Time, Count_tariff, Count_money FROM Contract WHERE id = 1'))
        tariff = list(cur.execute('SELECT * FROM Tariff'))

        label1 = QLabel(self)
        label1.setGeometry(50, 50, 300, 50)
        label1.setText('Год начала действия договора')

        label2 = QLabel(self)
        label2.setGeometry(50, 150, 300, 50)
        label2.setText('Месяц начала действия договора')

        label3 = QLabel(self)
        label3.setGeometry(50, 250, 300, 50)
        label3.setText('Срок действия договора (количество месяцев)')

        label4 = QLabel(self)
        label4.setGeometry(50, 350, 300, 50)
        label4.setText('Количество тарифов')

        label5 = QLabel(self)
        label5.setGeometry(50, 450, 300, 50)
        label5.setText('Денежная сумма по договору')

        label6 = QLabel(self)
        label6.setGeometry(850, 50, 300, 50)
        label6.setText('Тарифы')

        self.year = QSpinBox(self)
        self.year.setGeometry(350, 50, 200, 50)
        self.year.setMaximum(10000)
        
        font1 = self.year.font()
        font1.setPointSize(25)
        self.year.setFont(font1)

        self.month = QComboBox(self)
        self.month.setGeometry(350, 150, 200, 50)
        self.month.addItems([i.capitalize() for i in MONTH])
        font2 = self.month.font()
        font2.setPointSize(20)
        self.month.setFont(font2)

        self.time = QSpinBox(self)
        self.time.setGeometry(350, 250, 200, 50)

        font3 = self.time.font()
        font3.setPointSize(25)
        self.time.setFont(font3)

        self.countTariff = QSpinBox(self)
        self.countTariff.setGeometry(350, 350, 200, 50)

        font4 = self.countTariff.font()
        font4.setPointSize(25)
        self.countTariff.setFont(font4)
        self.countTariff.valueChanged.connect(self.countTariffChanged)

        self.summa = QSpinBox(self)
        self.summa.setGeometry(350, 450, 200, 50)
        self.summa.setMaximum(2147483647)

        font5 = self.summa.font()
        font5.setPointSize(25)
        self.summa.setFont(font5)

        self.table = QTableWidget(self)
        self.table.setGeometry(650, 150, 500, 250)
        self.table.setColumnCount(2)
        self.table.setHorizontalHeaderLabels(['id', 'Описание'])

        self.table.itemChanged.connect(self.itemChanged)


        if not params:
            self.year.setValue(int(datetime.datetime.now().year))
            self.time.setValue(12)
        else:
            params = params[0]
            self.year.setValue(params[0])
            self.month.setCurrentText(params[1])
            self.time.setValue(params[2])
            self.countTariff.setValue(params[3])
            self.summa.setValue(params[4])
            self.table.setRowCount(params[3])

        for i, val in enumerate(tariff):
            for j, elem in enumerate(val):
                self.table.setItem(i, j, QTableWidgetItem(str(elem)))

        self.table.resizeColumnsToContents()
            
            
        button = QPushButton('Сохранить', self)
        button.setGeometry(650, 450, 200, 50)
        button.clicked.connect(self.save)

    def itemChanged(self, item):
        row = item.row()
        col = item.column()
        
        new_value = item.text()

        if new_value == self.table.item(row, col).text():
            self.table.resizeColumnsToContents()
            return 0

        self.table.setItem(row, col, QTableWidgetItem(str(new_value)))

    def countTariffChanged(self):
        self.table.setRowCount(self.countTariff.value())     

    def save(self):
        flag = True
        for i in range(self.countTariff.value()):
            for j in range(2):
                val = self.table.item(i, j)
                if val is None:
                    flag = False
                    break
                if val.text() == '':
                    flag = False
                    break
        if flag and self.countTariff.value() != 0:
            val1, val2, val3, val4, val5 = self.year.value(), self.month.currentText(), self.time.value(), self.countTariff.value(), self.summa.value()
            sql_command = f'''UPDATE Contract SET Year = {val1}, Month = "{val2}", Time = {val3}, Count_tariff = {val4}, Count_money = {val5} WHERE id = 1'''
            cur.execute(sql_command)
            con.commit()

            tariff = [i[0] for i in list(cur.execute('SELECT id FROM Tariff'))]
            for i in tariff:
                cur.execute(f'DELETE FROM Tariff WHERE id = {i}')
            con.commit()
            for i in range(self.countTariff.value()):
                s = []
                for j in range(2):
                    s.append(self.table.item(i, j).text())
                cur.execute(f'INSERT INTO Tariff (id, Description) VALUES (?, ?)', (s[0], s[1]))
            con.commit()

            wind = QMessageBox(self)
            wind.setWindowTitle('Процесс успешно завершен')
            wind.setText('Данные успешно сохранены')
            wind.setIcon(QMessageBox.Information)
            wind.setStandardButtons(QMessageBox.Close)
            res = wind.exec()


class OpenPT1(Menu):
    def __init__(self):
        super().__init__()

    def initUI(self):
        self.setGeometry(100, 100, 1200, 600)
        self.setWindowTitle('Сводная таблица по категориям')

        tables = [i[0] for i in list(cur.execute('SELECT name FROM sqlite_master WHERE type="table"'))]

        category = ', '.join(['C' + str(ind + 1) + ' TEXT' for ind, i in enumerate([i[0] for i in list(set(list(cur.execute('SELECT Category FROM Database')))) if not i[0] is None and i[0]])])
        contract_info = list(cur.execute('SELECT Year, Month, Time FROM Contract WHERE id = "1"'))
        if category and contract_info:
            contract_info = list(contract_info[0])
            
            cur.execute(f'CREATE TABLE IF NOT EXISTS PivotTableCategories ( Month_year TEXT, {category}, Summa TEXT, Delta TEXT )')
            con.commit()

            if not 'PivotTableCategories' in tables:
                month1 = MONTH[MONTH.index(contract_info[1].lower()):]
                month2 = MONTH[:MONTH.index(contract_info[1].lower())]
                all_month = month1 + month2
                for i in range(contract_info[2] // 12):
                    for mon in month1:
                        cur.execute(f'INSERT INTO PivotTableCategories (Month_year) VALUES (?)', (mon.capitalize() + ' ' + str(int(contract_info[0]) + i),))
                    for mon in month2:
                        cur.execute(f'INSERT INTO PivotTableCategories (Month_year) VALUES (?)', (mon.capitalize() + ' ' + str(int(contract_info[0]) + i + 1),))

                for i in range(contract_info[2] % 12):
                    year = contract_info[2] // 12 + int(contract_info[0])
                    if all_month[i] in month1:
                        cur.execute(f'INSERT INTO PivotTableCategories (Month_year) VALUES (?)', (all_month[i].capitalize() + ' ' + str(year),))
                    if all_month[i] in month2:
                        cur.execute(f'INSERT INTO PivotTableCategories (Month_year) VALUES (?)', (all_month[i].capitalize() + ' ' + str(year + 1),))
                        
                con.commit()

            request = [list(i) for i in list(cur.execute('SELECT * FROM PivotTableCategories'))]
            self.table = QTableWidget(self)
            self.table.setGeometry(100, 100, 900, 400)
            self.table.setColumnCount(len(['Месяц/год'] +\
                                          [str(i) for i in sorted([i for i in list(set([i[0] if not i[0] is None else '' for i in list(cur.execute('SELECT Category FROM Database'))])) if i])] +\
                                          ['Сумма', 'Разница']))
            self.table.setRowCount(len(request))
            self.table.setHorizontalHeaderLabels(['Месяц/год'] +\
                                                 [str(i) for i in sorted([i for i in list(set([i[0] if not i[0] is None else '' for i in list(cur.execute('SELECT Category FROM Database'))])) if i])] +\
                                                 ['Сумма', 'Разница'])

            for i, val in enumerate(request):
                for j, elem in enumerate(val):
                    if str(elem) == 'None':
                        elem = ''
                    self.table.setItem(i, j, QTableWidgetItem(str(elem)))
            self.table.resizeColumnsToContents()
        else:
            self.status.setText('В базе данных нет категорий')
                

class OpenPT2(Menu):
    def __init__(self):
        super().__init__()

    def initUI(self):
        self.setGeometry(100, 100, 1200, 600)
        self.setWindowTitle('Сводная таблица по тарифам')

        tables = [i[0] for i in list(cur.execute('SELECT name FROM sqlite_master WHERE type="table"'))]

        contract_info = list(cur.execute('SELECT Year, Month, Time FROM Contract WHERE id = 1'))

        tariff = ', '.join(['T' + str(i[0]) + ' TEXT' for i in list(cur.execute('SELECT id FROM Tariff'))])

        if tariff and contract_info:
            contract_info = list(contract_info[0])
            
            cur.execute(f'CREATE TABLE IF NOT EXISTS PivotTableTariffs (Month_year TEXT, {tariff}, Summa TEXT, Delta TEXT)')
            con.commit()
            
            if not 'PivotTableTariffs' in tables:
                month1 = MONTH[MONTH.index(contract_info[1].lower()):]
                month2 = MONTH[:MONTH.index(contract_info[1].lower())]
                all_month = month1 + month2
                for i in range(contract_info[2] // 12):
                    for mon in month1:
                        cur.execute(f'INSERT INTO PivotTableTariffs (Month_year) VALUES (?)', (mon.capitalize() + ' ' + str(int(contract_info[0]) + i),))
                    for mon in month2:
                        cur.execute(f'INSERT INTO PivotTableTariffs (Month_year) VALUES (?)', (mon.capitalize() + ' ' + str(int(contract_info[0]) + i + 1),))
                        
                for i in range(contract_info[2] % 12):
                    year = contract_info[2] // 12 + int(contract_info[0])
                    if all_month[i] in month1:
                        cur.execute(f'INSERT INTO PivotTableTariffs (Month_year) VALUES (?)', (all_month[i].capitalize() + ' ' + str(year),))
                    if all_month[i] in month2:
                        cur.execute(f'INSERT INTO PivotTableTariffs (Month_year) VALUES (?)', (all_month[i].capitalize() + ' ' + str(year + 1),))
                    
                con.commit()

            request = [list(i) for i in list(cur.execute('SELECT * FROM PivotTableTariffs'))]
            self.table = QTableWidget(self)
            self.table.setGeometry(100, 100, 900, 400)
            self.table.setColumnCount(len(['Месяц/год'] +\
                                          ['Тариф ' + str(i[0]) for i in list(cur.execute('SELECT id FROM Tariff'))] +\
                                          ['Сумма', 'Разница']))
            self.table.setRowCount(len(request))
            self.table.setHorizontalHeaderLabels(['Месяц/год'] +\
                                                 ['Тариф ' + str(i[0]) for i in list(cur.execute('SELECT id FROM Tariff'))] +\
                                                 ['Сумма', 'Разница'])

            for i, val in enumerate(request):
                for j, elem in enumerate(val):
                    if str(elem) == 'None':
                        elem = ''
                    self.table.setItem(i, j, QTableWidgetItem(str(elem)))

            self.table.resizeColumnsToContents()
        else:
            self.status.setText('В базе данных нет тарифов')


class CalculateCosts(Menu):
    def __init__(self):
        super().__init__()

    def initUI(self):
        self.setGeometry(100, 100, 1200, 600)
        self.setWindowTitle('Расчет затрат')

        label1 = QLabel(self)
        label1.setGeometry(100, 60, 550, 40)
        label1.setText('Выберите месяц, за который хотите сделать расчет затрат')

        self.month1 = QCheckBox('Январь', self)
        self.month1.setGeometry(100, 100, 150, 30)
        self.month2 = QCheckBox('Февраль', self)
        self.month2.setGeometry(100, 130, 150, 30)
        self.month3 = QCheckBox('Март', self)
        self.month3.setGeometry(100, 160, 150, 30)
        self.month4 = QCheckBox('Апрель', self)
        self.month4.setGeometry(100, 190, 150, 30)
        self.month5 = QCheckBox('Май', self)
        self.month5.setGeometry(100, 220, 150, 30)
        self.month6 = QCheckBox('Июнь', self)
        self.month6.setGeometry(100, 250, 150, 30)
        self.month7 = QCheckBox('Июль', self)
        self.month7.setGeometry(100, 280, 150, 30)
        self.month8 = QCheckBox('Август', self)
        self.month8.setGeometry(100, 310, 150, 30)
        self.month9 = QCheckBox('Сентябрь', self)
        self.month9.setGeometry(100, 340, 150, 30)
        self.month10 = QCheckBox('Октябрь', self)
        self.month10.setGeometry(100, 370, 150, 30)
        self.month11 = QCheckBox('Ноябрь', self)
        self.month11.setGeometry(100, 400, 150, 30)
        self.month12 = QCheckBox('Декабрь', self)
        self.month12.setGeometry(100, 430, 150, 30)

        label2 = QLabel(self)
        label2.setGeometry(700, 60, 550, 40)
        label2.setText('Выберите год, за который хотите сделать расчет затрат')

        self.year = QSpinBox(self)
        self.year.setGeometry(700, 130, 200, 50)
        self.year.setMaximum(10000)
        self.year.setValue(int(datetime.datetime.now().year))
        font1 = self.year.font()
        font1.setPointSize(25)
        self.year.setFont(font1)

        label3 = QLabel(self)
        label3.setGeometry(700, 250, 450, 40)
        label3.setText('Сделать расчет для сводной таблицы')

        self.calculate_categories = QPushButton('по категориям', self)
        self.calculate_categories.setGeometry(700, 320, 250, 70)
        self.calculate_categories.clicked.connect(self.F_calculate_categories)

        self.calculate_tariffs = QPushButton('по тарифам', self)
        self.calculate_tariffs.setGeometry(700, 430, 250, 70)
        self.calculate_tariffs.clicked.connect(self.F_calculate_tariffs)

    def F_calculate_categories(self):
        try:
            tables = [i[0] for i in list(cur.execute('SELECT name FROM sqlite_master WHERE type="table"'))]

            category = ', '.join(['C' + str(ind + 1) + ' TEXT' for ind, i in enumerate([i[0] for i in list(set(list(cur.execute('SELECT Category FROM Database')))) if not i[0] is None and i[0]])])
            cur.execute(f'CREATE TABLE IF NOT EXISTS PivotTableCategories (Month_year TEXT, {category}, Summa TEXT, Delta TEXT)')
            con.commit()

            if not 'PivotTableCategories' in tables:
                contract_info = list(list(cur.execute('SELECT Year, Month, Time FROM Contract WHERE id = 1'))[0])

                month1 = MONTH[MONTH.index(contract_info[1].lower()):]
                month2 = MONTH[:MONTH.index(contract_info[1].lower())]
                all_month = month1 + month2
                for i in range(contract_info[2] // 12):
                    for mon in month1:
                        cur.execute(f'INSERT INTO PivotTableCategories (Month_year) VALUES (?)', (mon.capitalize() + ' ' + str(int(contract_info[0]) + i),))
                    for mon in month2:
                        cur.execute(f'INSERT INTO PivotTableCategories (Month_year) VALUES (?)', (mon.capitalize() + ' ' + str(int(contract_info[0]) + i + 1),))

                for i in range(contract_info[2] % 12):
                    year = contract_info[2] // 12 + int(contract_info[0])
                    if all_month[i] in month1:
                        cur.execute(f'INSERT INTO PivotTableCategories (Month_year) VALUES (?)', (all_month[i].capitalize() + ' ' + str(year),))
                    if all_month[i] in month2:
                        cur.execute(f'INSERT INTO PivotTableCategories (Month_year) VALUES (?)', (all_month[i].capitalize() + ' ' + str(year + 1),))

                con.commit()
            
            current_months = []
            for ind, i in enumerate([self.month1, self.month2, self.month3, self.month4,
                                     self.month5, self.month6, self.month7, self.month8,
                                     self.month9, self.month10, self.month11, self.month12]):
                if i.isChecked():
                    current_months.append(ind)

            k_month = 0
            for current_month in current_months:
                year = self.year.value()
                month = MONTH[current_month].capitalize()
        
                first_col = month + ' ' + str(year)

                last_month = [i[0] for i in list(cur.execute('SELECT Month_year FROM PivotTableCategories'))]

                if first_col in last_month:
                    k_month += 1

                    categories = sorted(list(set([i[0] for i in list(cur.execute('SELECT Category FROM Database')) if (not i[0] is None) and i[0]])))

                    all_sum = []
                    for categ in categories:
                        su = sum(list(map(float, [i[0] for i in list(cur.execute(f'SELECT ac1{MONTH_SHORT[current_month]} FROM Database WHERE Category = "{categ}"')) if (not i[0] is None) and i[0]])))
                        all_sum.append(su)

                    all_sum[categories.index('Неизвестные')] += sum(list(map(float, [i[0] for i in list(cur.execute(f'SELECT ac1{MONTH_SHORT[current_month]} FROM Database WHERE Category IS NULL OR Category = "{""}"')) if (not i[0] is None) and i[0]])))

                    itog = sum(all_sum)
                    all_sum.append(itog)

                    param_contaract = [i for i in list(cur.execute('SELECT Year, Month FROM Contract'))[0]]

                    if str(year) != str(param_contaract[0]) or str(month) != str(param_contaract[1]):
                        last_su_for_month = [i[0] for i in list(cur.execute('SELECT Summa FROM PivotTableCategories'))]
                        last_su = last_su_for_month[last_month.index(first_col) - 1]
                        if last_su is None or (not last_su):
                            last_su = 0
                        else:
                            last_su = float(last_su)
                        delta = itog - last_su
                    else:
                        delta = 0
                    all_sum.append(delta)

                    categ_in_db = ['C' + str(i) for i in range(1, len(all_sum) - 1)] + ['Summa', 'Delta']
                    for_req = ', '.join([col + ' = ' + str(val) for col, val in zip(categ_in_db, all_sum)])
                    cur.execute(f'UPDATE PivotTableCategories SET {for_req} WHERE Month_year = "{first_col}"')
                    con.commit()

            if k_month:
                wind = QMessageBox(self)
                wind.setWindowTitle('Процесс успешно завершен')
                wind.setText('Расчет по категориям успешно произведен')
                wind.setIcon(QMessageBox.Information)
                wind.setStandardButtons(QMessageBox.Close)
                res = wind.exec()
        except Exception:
            wind = QMessageBox(self)
            wind.setWindowTitle('Ошибка')
            wind.setText('В сводной таблице неправильные столбцы, чтобы исправить ошибку удалите сводную таблицу')
            wind.setIcon(QMessageBox.Critical)
            wind.setStandardButtons(QMessageBox.Close)
            res = wind.exec()

    def F_calculate_tariffs(self):
        try:
            tables = [i[0] for i in list(cur.execute('SELECT name FROM sqlite_master WHERE type="table"'))]

            tariff = ', '.join(['T' + str(i[0]) + ' TEXT' for i in list(cur.execute('SELECT id FROM Tariff'))])
            cur.execute(f'CREATE TABLE IF NOT EXISTS PivotTableTariffs (Month_year TEXT, {tariff}, Summa TEXT, Delta TEXT)')
            con.commit()

            if not 'PivotTableTariffs' in tables:
                contract_info = list(list(cur.execute('SELECT Year, Month, Time FROM Contract WHERE id = 1'))[0])
                month1 = MONTH[MONTH.index(contract_info[1].lower()):]
                month2 = MONTH[:MONTH.index(contract_info[1].lower())]
                all_month = month1 + month2
                for i in range(contract_info[2] // 12):
                    for mon in month1:
                        cur.execute(f'INSERT INTO PivotTableTariffs (Month_year) VALUES (?)', (mon.capitalize() + ' ' + str(int(contract_info[0]) + i),))
                    for mon in month2:
                        cur.execute(f'INSERT INTO PivotTableTariffs (Month_year) VALUES (?)', (mon.capitalize() + ' ' + str(int(contract_info[0]) + i + 1),))

                for i in range(contract_info[2] % 12):
                    year = contract_info[2] // 12 + int(contract_info[0])
                    if all_month[i] in month1:
                        cur.execute(f'INSERT INTO PivotTableTariffs (Month_year) VALUES (?)', (all_month[i].capitalize() + ' ' + str(year),))
                    if all_month[i] in month2:
                        cur.execute(f'INSERT INTO PivotTableTariffs (Month_year) VALUES (?)', (all_month[i].capitalize() + ' ' + str(year + 1),))
                        
                con.commit()
            
            current_months = []
            for ind, i in enumerate([self.month1, self.month2, self.month3, self.month4,
                                     self.month5, self.month6, self.month7, self.month8,
                                     self.month9, self.month10, self.month11, self.month12]):
                if i.isChecked():
                    current_months.append(ind)

            k_month = 0
            for current_month in current_months:
                year = self.year.value()
                month = MONTH[current_month].capitalize()
        
                first_col = month + ' ' + str(year)

                last_month = [i[0] for i in list(cur.execute('SELECT Month_year FROM PivotTableTariffs'))]

                if first_col in last_month:
                    k_month += 1
                    
                    tariffs = [i[0] for i in list(cur.execute('SELECT id FROM TARIFF'))]
                    all_sum = []

                    for tar in tariffs:
                        su = sum(list(map(float, [i[0] for i in list(cur.execute(f'SELECT ac1{MONTH_SHORT[current_month]} FROM Database WHERE TariffN = "{tar}"')) if (not i[0] is None) and i[0]])))
                        all_sum.append(su)

                    itog = sum(all_sum)
                    all_sum.append(itog)

                    param_contaract = [i for i in list(cur.execute('SELECT Year, Month FROM Contract'))[0]]
                    if str(year) != str(param_contaract[0]) or str(month) != str(param_contaract[1]):
                        last_su_for_month = [i[0] for i in list(cur.execute('SELECT Summa FROM PivotTableTariffs'))]
                        last_su = last_su_for_month[last_month.index(first_col) - 1]
                        if last_su is None or (not last_su):
                            last_su = 0
                        else:
                            last_su = float(last_su)
                        delta = itog - last_su
                    else:
                        delta = 0
                    all_sum.append(delta)

                    tariff_in_db = ['T' + str(i) for i in tariffs] + ['Summa', 'Delta']
                    for_req = ', '.join([col + ' = ' + str(val) for col, val in zip(tariff_in_db, all_sum)])
                    cur.execute(f'UPDATE PivotTableTariffs SET {for_req} WHERE Month_year = "{first_col}"')
                    con.commit()

            if k_month:
                wind = QMessageBox(self)
                wind.setWindowTitle('Процесс успешно завершен')
                wind.setText('Расчет по тарифам успешно произведен')
                wind.setIcon(QMessageBox.Information)
                wind.setStandardButtons(QMessageBox.Close)
                res = wind.exec()
        except Exception:
            wind = QMessageBox(self)
            wind.setWindowTitle('Ошибка')
            wind.setText('В сводной таблице неправильные столбцы, чтобы исправить ошибку удалите сводную таблицу')
            wind.setIcon(QMessageBox.Critical)
            wind.setStandardButtons(QMessageBox.Close)
            res = wind.exec()


class GrafficPrognos(Menu):
    def __init__(self):
        super().__init__()

    def initUI(self):
        self.setGeometry(100, 100, 1200, 600)
        self.setWindowTitle('Прогнозный график')

        tables = [i[0] for i in list(cur.execute('SELECT name FROM sqlite_master WHERE type="table"'))]

        if 'PivotTableCategories' in tables:
            sum_for_month = [float(i[0]) for i in list(cur.execute('SELECT Summa FROM PivotTableCategories')) if not i[0] is None]
            month_year = [MONTH_SHORT[MONTH.index(str(i[0].split()[0]).lower())] + '\n' + str(i[0].split()[1]) for i in list(cur.execute('SELECT Month_year FROM PivotTableCategories'))]

            self.graph = pg.PlotWidget(self)
            self.graph.setGeometry(50, 50, 1100, 500)
            self.graph.setBackground('w')

            pen = pg.mkPen(color=(255, 0, 0), width=2)

            x = [i for i in range(list(cur.execute('SELECT time FROM Contract'))[0][0])]
            y = sum_for_month

            su_contract = int(list(cur.execute('SELECT Count_money FROM Contract'))[0][0])
            if len(y) < len(x):
                sr_zn = (su_contract - sum(y)) / len(x) - len(y)
                while len(y) < len(x):
                    y.append(sr_zn)

            val_for_x = self.graph.getAxis('bottom')

            val_for_x.tickFont = QFont()

            val_for_x.setTicks([list(zip(x, month_year))])

            self.graph.plot(x, y, pen=pen)

            self.graph.showGrid(x=True, y=True)
        else:
            self.status.setText('Нет данных')
        

class DelINFO(Menu):
    def __init__(self):
        super().__init__()

    def initUI(self):
        self.setGeometry(100, 100, 1200, 600)
        self.setWindowTitle('Удаление информации')

        label1 = QLabel(self)
        label1.setGeometry(100, 100, 400, 50)
        label1.setText('Удаление информации по месяцам из базы данных:')

        label2 = QLabel(self)
        label2.setGeometry(700, 100, 400, 50)
        label2.setText('Удаление сводной таблицы:')

        btn1 = QPushButton(self)
        btn1.setGeometry(100, 200, 400, 100)
        btn1.setText('Удалить информацию по месяцам из базы данных')
        btn1.clicked.connect(self.del1)
        
        btn2 = QPushButton(self)
        btn2.setGeometry(700, 200, 400, 100)
        btn2.setText('Удалить сводную таблицу (без сохранения)')
        btn2.clicked.connect(self.del2)

        btn3 = QPushButton(self)
        btn3.setGeometry(700, 350, 400, 100)
        btn3.setText('Удалить сводную таблицу (с сохранением)')
        btn3.clicked.connect(self.del3)

    def del1(self):
        check = QMessageBox(self)
        check.setWindowTitle('Удаление данных')
        check.setText('Вы действительно хотите удалить информацию по месяцам из базы данных?')
        check.setIcon(QMessageBox.Question)
        check.setStandardButtons(QMessageBox.Cancel | QMessageBox.Ok)

        res = check.exec()
        
        if res == QMessageBox.Ok:
            columns = [j + i for i in MONTH_SHORT for j in ['tr1', 'tr2', 'tr3',
                                                            'ac1', 'ac2', 'ac3', 'ac4', 'ac5', 'ac6', 'ac7', 'ac8', 'ac9', 'ac10', 'ac11', 'ac12',
                                                            'su1', 'su2']]
            for col in columns:
                cur.execute(f'ALTER TABLE Database DROP COLUMN {col}')
            con.commit()
            for col in columns:
                cur.execute(f'ALTER TABLE Database ADD COLUMN {col} TEXT')
            con.commit()

            wind = QMessageBox(self)
            wind.setWindowTitle('Процесс успешно завершен')
            wind.setText('Информация по месяцам удалена из базы данных')
            wind.setIcon(QMessageBox.Information)
            wind.setStandardButtons(QMessageBox.Close)
            res = wind.exec()

    def del2(self):
        check = QMessageBox(self)
        check.setWindowTitle('Удаление данных')
        check.setText('Вы действительно хотите далить сводную таблицу без сохранения?')
        check.setIcon(QMessageBox.Question)
        check.setStandardButtons(QMessageBox.Cancel | QMessageBox.Ok)

        res = check.exec()

        if res == QMessageBox.Ok:
            cur.execute('DROP TABLE IF EXISTS PivotTableCategories')
            cur.execute('DROP TABLE IF EXISTS PivotTableTariffs')
            con.commit()

            wind = QMessageBox(self)
            wind.setWindowTitle('Процесс успешно завершен')
            wind.setText('Сводная таблица удалена без сохранения')
            wind.setIcon(QMessageBox.Information)
            wind.setStandardButtons(QMessageBox.Close)
            res = wind.exec()

    def del3(self):
        check = QMessageBox(self)
        check.setWindowTitle('Удаление данных')
        check.setText('Вы действительно хотите далить сводную таблицу с сохранением?')
        check.setIcon(QMessageBox.Question)
        check.setStandardButtons(QMessageBox.Cancel | QMessageBox.Ok)

        res = check.exec()

        if res == QMessageBox.Ok:
            start = ' '.join(list(map(str, list(list(cur.execute('SELECT Month, Year FROM Contract'))[0]))))
            time = str(list(cur.execute('SELECT Time FROM Contract'))[0][0]) + ' мес.'

            type1 = 'По категориям'
            header1 = ['Месяц/год'] +\
                      [str(i) for i in sorted([i for i in list(set([i[0] if not i[0] is None else '' for i in list(cur.execute('SELECT Category FROM Database'))])) if i])] +\
                      ['Сумма', 'Разница']
            info1 = '|'.join(header1) + '\n' + '\n'.join(['|'.join(list(map(str, list(i)))) for i in list(cur.execute('SELECT * FROM PivotTableCategories'))])

            type2 = 'По тарифам'
            header2 = ['Месяц/год'] +\
                      ['Тариф ' + str(i[0]) for i in list(cur.execute('SELECT id FROM Tariff'))] +\
                      ['Сумма', 'Разница']
            info2 = '|'.join(header2) + '\n' + '\n'.join(['|'.join(list(map(str, list(i)))) for i in list(cur.execute('SELECT * FROM PivotTableTariffs'))])

            filename, _ = QFileDialog.getSaveFileName(self, 'Сохранить сводную таблицу как файл', '')
            if filename:
                if filename.split('.')[-1] != 'xls':
                    wind = QMessageBox(self)
                    wind.setWindowTitle('Ошибка')
                    wind.setText('Укажите для файла расширение xls,\nиначе данные не сохранятся,\nсейчас ничего не произошло')
                    wind.setIcon(QMessageBox.Critical)
                    wind.setStandardButtons(QMessageBox.Close)
                    res = wind.exec()
                else:
                    wb = xlwt.Workbook()
                    ws_category = wb.add_sheet('По категориям')
                    ws_tariff = wb.add_sheet('По тарифам')

                    for i, val in enumerate(info1.split('\n')):
                        for j, elem in enumerate(val.split('|')):
                            ws_category.write(i, j, str(elem))

                    for i, val in enumerate(info2.split('\n')):
                        for j, elem in enumerate(val.split('|')):
                            ws_tariff.write(i, j, str(elem))
                        
                    wb.save(filename)

                    inf_savePT = list(cur.execute('SELECT id FROM SavePT'))
                    if inf_savePT:
                        val_id = int(inf_savePT[-1][0]) + 1
                    else:
                        val_id = 1
                        
                    cur.execute(f'INSERT INTO SavePT (id, Start, Time, Type, INFO) VALUES ({val_id}, "{start}", "{time}", "{type1}", "{info1}")')
                    cur.execute(f'INSERT INTO SavePT (id, Start, Time, Type, INFO) VALUES ({val_id + 1}, "{start}", "{time}", "{type2}", "{info2}")')
                    con.commit()

                    cur.execute('DROP TABLE IF EXISTS PivotTableCategories')
                    cur.execute('DROP TABLE IF EXISTS PivotTableTariffs')
                    con.commit()

                    wind = QMessageBox(self)
                    wind.setWindowTitle('Процесс успешно завершен')
                    wind.setText('Сводная таблица удалена с сохранением. Данные успешно экспортировались в файл Excel и сохранились в окне Сохраненная информация')
                    wind.setIcon(QMessageBox.Information)
                    wind.setStandardButtons(QMessageBox.Close)
                    res = wind.exec()
            else:
                wind = QMessageBox(self)
                wind.setWindowTitle('Ошибка')
                wind.setText('Не указано место для сохранения файла')
                wind.setIcon(QMessageBox.Critical)
                wind.setStandardButtons(QMessageBox.Close)
                res = wind.exec()


class AllPT(Menu):
    def __init__(self):
        super().__init__()

    def initUI(self):
        self.setGeometry(100, 100, 1200, 600)
        self.setWindowTitle('Все сводные таблицы')

        request = [list(i) for i in list(cur.execute('SELECT id, Start, Time, Type FROM SavePT'))]

        if request:

            self.table = QTableWidget(self)
            self.table.setGeometry(100, 100, 700, 400)
            self.table.setColumnCount(4)
            self.table.setRowCount(len(request))
            self.table.setHorizontalHeaderLabels(['id', 'Месяц/год начала действия договора',
                                                  'Срок действия договора (количество месяцев)', 'Тип сводной таблицы'])

            for i, val in enumerate(request):
                for j, elem in enumerate(val):
                    self.table.setItem(i, j, QTableWidgetItem(str(elem)))

            self.table.resizeColumnsToContents()

            self.table.cellClicked.connect(self.openSavePT)

        else:
            self.status.setText('Нет сохраненных сводных таблиц')

    def openSavePT(self, row, col):
        self.close()
        self.main = DatabaseSIMCard()
        self.main.clear_all()

        self.W_open_save_pt = OpenSavePT(int(self.table.item(row, 0).text()))
        self.W_open_save_pt.initUI()
        self.W_open_save_pt.show()
        

class OpenSavePT(Menu):
    def __init__(self, id_val):
        super().__init__()
        self.id_val = id_val

    def initUI(self):
        self.setGeometry(100, 100, 1200, 600)
        self.setWindowTitle(f'Сохраненная сводная таблица с id {self.id_val}')

        info = list(cur.execute(f'SELECT INFO FROM SavePT WHERE id = {self.id_val}'))[0][0]

        rows = [i.split('|') for i in info.split('\n')]

        self.table = QTableWidget(self)
        self.table.setGeometry(100, 100, 1000, 400)
        self.table.setColumnCount(len(rows[0]))
        self.table.setRowCount(len(rows) - 1)
        self.table.setHorizontalHeaderLabels(rows[0])

        for i, cols in enumerate(rows[1:]):
            for j, elem in enumerate(cols):
                self.table.setItem(i, j, QTableWidgetItem(str(elem)))

        self.table.resizeColumnsToContents()

        
def except_hook(cls, exception, traceback):
    sys.__excepthook__(cls, exception, traceback)
    

if __name__ == '__main__':
    app = QApplication(sys.argv)
    form = DatabaseSIMCard()
    form.show()
    sys.excepthook = except_hook
    sys.exit(app.exec())
