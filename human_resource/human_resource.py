import os
import sys
import time
from datetime import datetime
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QMainWindow
from PyQt5.QtCore import QDate, QEventLoop, QThread
import tkinter
from tkinter import filedialog
from openpyxl import Workbook

import human_resource_ui
from analyze_utils.employee_analyze import EmployeeAnalyze
from analyze_utils.real_check_in_and_change_analyze import ReadCheckInAndChangeAnalyze
from analyze_utils.real_leave_analyze import RealLeaveAnalyze
from analyze_utils.expect_check_in_leave_analyze import ExpectCheckInLeaveAnalyze

class HumanResource:
    # month
    _month = None

    # workbook
    _human_resource_workbook: Workbook = None

    # employee
    _employee_excel_path = None
    _employee_analyze: EmployeeAnalyze = None

    # real check in and change excel
    _real_check_in_and_change_excel_path = None
    _real_check_in_and_change_analyze: ReadCheckInAndChangeAnalyze = None

    # real leave excel
    _real_leave_excel_path = None
    _real_leave_analyze: RealLeaveAnalyze = None

    # expect check in and leave analyze
    _expect_check_in_leave_excel_path = None
    _expect_check_in_leave_analyze: ExpectCheckInLeaveAnalyze = None


    def __init__(self, month,
                 employee_excel_path,
                 real_check_in_and_change_excel_path,
                 real_leave_excel_path,
                 expect_check_in_leave_excel_path):
        self._month = month
        self._employee_excel_path = employee_excel_path
        self._real_check_in_and_change_excel_path = real_check_in_and_change_excel_path
        self._real_leave_excel_path = real_leave_excel_path
        self._expect_check_in_leave_excel_path = expect_check_in_leave_excel_path

        self._init_employee_analyze()
        self._init_real_check_in_and_change_analyze()
        self._init_real_leave_analyze()
        self._init_expect_check_in_leave_analyze()

        self._create_empty_human_resource_excel()
        self._create_main_sheet()
        self._create_head_office()
        self._create_miopane()
        self._create_miacucina()
        self._create_other_company()
        self._create_dreamers()


    def _init_employee_analyze(self):
        print('Initialize employee ()... '.format(self._employee_excel_path), end='')
        self._employee_analyze = EmployeeAnalyze(employee_excel_path=self._employee_excel_path)
        print('done.')


    def _init_real_check_in_and_change_analyze(self):
        print('Initialize check in and change analyze unit... ', end='')
        self._real_check_in_and_change_analyze = ReadCheckInAndChangeAnalyze(excel_path=self._real_check_in_and_change_excel_path)
        print('done.')


    def _init_real_leave_analyze(self):
        print('Initialize real leave analyze unit... ', end='')
        self._real_leave_analyze = RealLeaveAnalyze(excel_path=self._real_leave_excel_path)
        print('done.')


    def _init_expect_check_in_leave_analyze(self):
        print('Initialize expect check in leave analyze unit... ', end='')
        self._expect_check_in_leave_excel_path = ExpectCheckInLeaveAnalyze(excel_path=self._expect_check_in_leave_excel_path,
                                                                           start_date_str='2024/03/01',
                                                                           end_date_str='2024/03/10')
        print('done.')


    def _create_empty_human_resource_excel(self):
        self._human_resource_workbook = Workbook()
        self._human_resource_workbook.remove(self._human_resource_workbook['Sheet'])


    def _create_main_sheet(self):
        print('Generate main sheet... ', end='')
        from main import HumanResourceMain
        human_resource_main = HumanResourceMain(human_resource_workbook=self._human_resource_workbook,
                                                month=self._month,
                                                employee_analyze=self._employee_analyze,
                                                real_check_in_and_change_analyze=self._real_check_in_and_change_analyze,
                                                real_leave_analyze=self._real_leave_analyze,
                                                expect_check_in_leave_analyze=self._expect_check_in_leave_excel_path)
        print('done.')


    def _create_head_office(self):
        print('Generate head office... ', end='')
        from head_office import HeadOffice
        head_office = HeadOffice(human_resource_workbook=self._human_resource_workbook,
                                 month=self._month,
                                 employee_analyze=self._employee_analyze,
                                 real_check_in_and_change_analyze=self._real_check_in_and_change_analyze,
                                 real_leave_analyze=self._real_leave_analyze,
                                 expect_check_in_leave_analyze=self._expect_check_in_leave_excel_path)
        print('done.')


    def _create_miopane(self):
        print('Generate miopane... ', end='')
        from miopane import Miopane
        miopane = Miopane(human_resource_workbook=self._human_resource_workbook,
                          month=self._month,
                          employee_analyze=self._employee_analyze,
                          real_check_in_and_change_analyze=self._real_check_in_and_change_analyze,
                          real_leave_analyze=self._real_leave_analyze,
                          expect_check_in_leave_analyze=self._expect_check_in_leave_excel_path)
        print('done.')


    def _create_miacucina(self):
        print('Generate miacucine... ', end='')
        from miacucina import Miacucina
        miacucina = Miacucina(human_resource_workbook=self._human_resource_workbook,
                              month=self._month,
                              employee_analyze=self._employee_analyze,
                              real_check_in_and_change_analyze=self._real_check_in_and_change_analyze,
                              real_leave_analyze=self._real_leave_analyze,
                              expect_check_in_leave_analyze=self._expect_check_in_leave_excel_path)
        print('done.')


    def _create_other_company(self):
        print('Generate other company... ', end='')
        from other_company import OtherCompany
        other_company = OtherCompany(human_resource_workbook=self._human_resource_workbook,
                                     month=self._month,
                                     employee_analyze=self._employee_analyze,
                                     real_check_in_and_change_analyze=self._real_check_in_and_change_analyze,
                                     real_leave_analyze=self._real_leave_analyze,
                                     expect_check_in_leave_analyze=self._expect_check_in_leave_excel_path)
        print('done.')


    def _create_dreamers(self):
        print('Generate dreamers... ', end='')
        from dreamers import Dreamers
        dreamers = Dreamers(human_resource_workbook=self._human_resource_workbook,
                            month=self._month,
                            employee_analyze=self._employee_analyze,
                            real_check_in_and_change_analyze=self._real_check_in_and_change_analyze,
                            real_leave_analyze=self._real_leave_analyze,
                            expect_check_in_leave_analyze=self._expect_check_in_leave_excel_path)
        print('done.')


    def save(self, save_path):
        print('Saving... ', end='')
        self._human_resource_workbook.save(save_path)
        print('done.')




class GenerateThreads(QThread):
    _month = None
    _employee_excel_path = None
    _real_check_in_and_change_excel_path = None
    _real_leave_excel_path = None
    _expect_check_in_leave_excel_path = None
    _save_path = None
    _callback = None
    def __init__(self, month,
                 employee_excel_path,
                 real_check_in_and_change_excel_path,
                 real_leave_excel_path,
                 expect_check_in_leave_excel_path,
                 save_path,
                 callback):
        self._month = month
        self._employee_excel_path = employee_excel_path
        self._real_check_in_and_change_excel_path = real_check_in_and_change_excel_path
        self._real_leave_excel_path = real_leave_excel_path
        self._expect_check_in_leave_excel_path = expect_check_in_leave_excel_path
        self._save_path = save_path
        self._callback = callback
        super().__init__()

    def run(self):
        try:
            human_resource = HumanResource(month=self._month,
                                           employee_excel_path=self._employee_excel_path,
                                           real_check_in_and_change_excel_path=self._real_check_in_and_change_excel_path,
                                           real_leave_excel_path=self._real_leave_excel_path,
                                           expect_check_in_leave_excel_path=self._expect_check_in_leave_excel_path)
            human_resource.save(save_path=self._save_path)
            self._callback('產生成功\n儲存於: {}'.format(self._save_path))

        except Exception as e:
            print(e)
            self._callback(str(e))



class HumanResourceUI(QMainWindow, human_resource_ui.Ui_MainWindow):
    tkinter_root = tkinter.Tk()
    tkinter_root.withdraw()

    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self._init_date()
        self._set_button_click()


    def _set_button_click(self):
        self.pushButton_convert.clicked.connect(self._convert)
        self.pushButton_generate.clicked.connect(self._generate)
        self.pushButton_converter_real_leave_select_file.clicked.connect(self._button_converter_real_leave_select_file_click)
        self.pushButton_converter_employee_select_file.clicked.connect(self._button_converter_employee_select_file_click)
        self.pushButton_converter_leave_list_select_file.clicked.connect(self._button_converter_leave_list_select_file_click)
        self.pushButton_converter_save_select_dir.clicked.connect(self._button_converter_save_select_directory_click)
        self.pushButton_employee_select_file.clicked.connect(self._button_employee_select_file_click)
        self.pushButton_real_check_in_and_change_select_file.clicked.connect(self._button_real_check_in_and_change_select_file_click)
        self.pushButton_real_leave_select_file.clicked.connect(self._button_real_leave_select_file_click)
        self.pushButton_manager_report_select_file.clicked.connect(self._button_manager_report_select_file_click)
        self.pushButton_save_select_dir.clicked.connect(self._button_save_select_directory_click)

    def _convert(self):
        from leave_list_add_id_converter import LeaveListAddIDConverter
        real_leave_excel_path = self.lineEdit_converter_real_leave_value.text()
        employee_excel_path = self.lineEdit_converter_employee_value.text()
        leave_list_excel_path = self.lineEdit_converter_leave_list_value.text()
        save_dir = self.lineEdit_converter_save_value.text() if self.lineEdit_converter_save_value.text() != '' else './'
        save_path = os.path.join(save_dir, os.path.basename(real_leave_excel_path).replace('.xlsx', '(New).xlsx'))

        try:
            if real_leave_excel_path != '' and employee_excel_path != '' and leave_list_excel_path != '':
                self._print_log(log='實際離職: {}\n目前人力配置: {}\n離職名單: {}\n'.format(real_leave_excel_path, employee_excel_path, leave_list_excel_path))
                self._print_log(log='開始轉換...')
                converter = LeaveListAddIDConverter(leave_list_excel_path=real_leave_excel_path,
                                                    employee_excel_path=employee_excel_path,
                                                    leave_employee_excel_path=leave_list_excel_path,
                                                    new_leave_list_excel_path=save_path)
                converter.convert()
                self._print_log(log='轉換成功\n儲存於: {}'.format(save_path))
            else:
                self._print_log(log='Error: 有檔案欄位為空!')

        except Exception as e:
            self._print_log(log='Convert error: \n{}'.format(e))

    def _generate(self):
        month = self.comboBox_month.currentText()
        employee_excel_path = self.lineEdit_employee_value.text()
        real_check_in_and_change_excel_path = self.lineEdit_real_check_in_and_change_value.text()
        real_leave_excel_path = self.lineEdit_real_leave_value.text()
        expect_check_in_leave_excel_path = self.lineEdit_manager_report_value.text()

        save_dir = self.lineEdit_save_value.text() if self.lineEdit_save_value.text() != '' else './'
        save_path = os.path.join(save_dir, '人力資源報表.xlsx')

        if employee_excel_path != '' and real_check_in_and_change_excel_path != '' and real_leave_excel_path != '' and expect_check_in_leave_excel_path != '':
            self.pushButton_generate.setEnabled(False)
            self._print_log(log='統計月份: {}\n目前人力配置: {}\n實際入職+單位異動: {}\n實際離職: {}\n預計報到離職+單位主管回報人力缺額: {}\n'
                            .format(month,
                                    employee_excel_path,
                                    real_check_in_and_change_excel_path,
                                    real_leave_excel_path,
                                    expect_check_in_leave_excel_path))
            self._print_log(log='開始產生...')

            self.generate_threads = GenerateThreads(month=month,
                                                    employee_excel_path=employee_excel_path,
                                                    real_check_in_and_change_excel_path=real_check_in_and_change_excel_path,
                                                    real_leave_excel_path=real_leave_excel_path,
                                                    expect_check_in_leave_excel_path=expect_check_in_leave_excel_path,
                                                    save_path=save_path,
                                                    callback=self._generate_thread_callback)
            self.generate_threads.start()
        else:
            self._print_log(log='Error: 有檔案欄位為空!')

    def _generate_thread_callback(self, log):
        self.pushButton_generate.setEnabled(True)
        self._print_log(log=log)

    def _init_date(self):
        self.comboBox_month.setCurrentIndex(datetime.now().month - 1)
        self.dateEdit_start.setDate(QDate.currentDate())
        self.dateEdit_end.setDate(QDate.currentDate())

    def _button_converter_real_leave_select_file_click(self):
        file_path = filedialog.askopenfilename()
        if file_path != '':
            self.lineEdit_converter_real_leave_value.setText(file_path)

    def _button_converter_employee_select_file_click(self):
        file_path = filedialog.askopenfilename()
        if file_path != '':
            self.lineEdit_converter_employee_value.setText(file_path)

    def _button_converter_leave_list_select_file_click(self):
        file_path = filedialog.askopenfilename()
        if file_path != '':
            self.lineEdit_converter_leave_list_value.setText(file_path)

    def _button_converter_save_select_directory_click(self):
        dir_path = filedialog.askdirectory()
        if dir_path != '':
            self.lineEdit_converter_save_value.setText(dir_path)

    def _button_employee_select_file_click(self):
        file_path = filedialog.askopenfilename()
        if file_path != '':
            self.lineEdit_employee_value.setText(file_path)

    def _button_real_check_in_and_change_select_file_click(self):
        file_path = filedialog.askopenfilename()
        if file_path != '':
            self.lineEdit_real_check_in_and_change_value.setText(file_path)

    def _button_real_leave_select_file_click(self):
        file_path = filedialog.askopenfilename()
        if file_path != '':
            self.lineEdit_real_leave_value.setText(file_path)

    def _button_manager_report_select_file_click(self):
        file_path = filedialog.askopenfilename()
        if file_path != '':
            self.lineEdit_manager_report_value.setText(file_path)

    def _button_save_select_directory_click(self):
        dir_path = filedialog.askdirectory()
        if dir_path != '':
            self.lineEdit_save_value.setText(dir_path)

    def _print_log(self, log):
        self.textBrowser_debug_message.append('\n' + log)
        # # self.textBrowser_debug_message.verticalScrollBar().setValue(self.textBrowser_debug_message.verticalScrollBar().maximum())
        # self.textBrowser_debug_message.update()
        # QtWidgets.QApplication.processEvents(QEventLoop.AllEvents)



if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    window = HumanResourceUI()
    window.show()

    sys.exit(app.exec_())


    #
    # root = tkinter.Tk()
    # root.withdraw()
    #
    # file_path = filedialog.askopenfilename()
    # print('file path: {}'.format(file_path))
    #
    human_resource = HumanResource(month=3,
                                   employee_excel_path='../Document/目前人力配置.xlsx',
                                   real_check_in_and_change_excel_path='../Document/實際入職+單位異動.xlsx',
                                   real_leave_excel_path='../Document/2月實際離職(new).xlsx',
                                   expect_check_in_leave_excel_path='../Document/預計報到離職+單位主管回報人力缺額.xlsx')
    human_resource.save(save_path='./人力資源報表.xlsx')