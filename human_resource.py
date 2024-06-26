import os
import sys
import traceback
from datetime import datetime
from PyQt5 import QtWidgets, QtCore, QtGui
from PyQt5.QtWidgets import QMainWindow, QFileDialog
from PyQt5.QtCore import QDate, QThread
from openpyxl import Workbook

from human_resource_utils.human_resource_ui import Ui_MainWindow
from human_resource_utils import __version__
from human_resource_utils.analyze_utils import EmployeeAnalyze, ReadCheckInAndChangeAnalyze, RealLeaveAnalyze, ExpectCheckInLeaveAnalyze

class HumanResource:
    # month
    _month = None

    # date range
    _start_date = None
    _end_date = None

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


    def __init__(self, month, start_date, end_date,
                 employee_excel_path,
                 real_check_in_and_change_excel_path,
                 real_leave_excel_path,
                 expect_check_in_leave_excel_path):
        self._month = month
        self._start_date = start_date
        self._end_date = end_date
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
                                                                           start_date_str=self._start_date,
                                                                           end_date_str=self._end_date)
        print('done.')


    def _create_empty_human_resource_excel(self):
        self._human_resource_workbook = Workbook()
        self._human_resource_workbook.remove(self._human_resource_workbook['Sheet'])


    def _create_main_sheet(self):
        print('Generate main sheet... ', end='')
        from human_resource_utils.main import HumanResourceMain
        human_resource_main = HumanResourceMain(human_resource_workbook=self._human_resource_workbook,
                                                month=self._month,
                                                employee_analyze=self._employee_analyze,
                                                real_check_in_and_change_analyze=self._real_check_in_and_change_analyze,
                                                real_leave_analyze=self._real_leave_analyze,
                                                expect_check_in_leave_analyze=self._expect_check_in_leave_excel_path)
        print('done.')


    def _create_head_office(self):
        print('Generate head office... ', end='')
        from human_resource_utils.head_office import HeadOffice
        head_office = HeadOffice(human_resource_workbook=self._human_resource_workbook,
                                 month=self._month,
                                 employee_analyze=self._employee_analyze,
                                 real_check_in_and_change_analyze=self._real_check_in_and_change_analyze,
                                 real_leave_analyze=self._real_leave_analyze,
                                 expect_check_in_leave_analyze=self._expect_check_in_leave_excel_path)
        print('done.')


    def _create_miopane(self):
        print('Generate miopane... ', end='')
        from human_resource_utils.miopane import Miopane
        miopane = Miopane(human_resource_workbook=self._human_resource_workbook,
                          month=self._month,
                          employee_analyze=self._employee_analyze,
                          real_check_in_and_change_analyze=self._real_check_in_and_change_analyze,
                          real_leave_analyze=self._real_leave_analyze,
                          expect_check_in_leave_analyze=self._expect_check_in_leave_excel_path)
        print('done.')


    def _create_miacucina(self):
        print('Generate miacucine... ', end='')
        from human_resource_utils.miacucina import Miacucina
        miacucina = Miacucina(human_resource_workbook=self._human_resource_workbook,
                              month=self._month,
                              employee_analyze=self._employee_analyze,
                              real_check_in_and_change_analyze=self._real_check_in_and_change_analyze,
                              real_leave_analyze=self._real_leave_analyze,
                              expect_check_in_leave_analyze=self._expect_check_in_leave_excel_path)
        print('done.')


    def _create_other_company(self):
        print('Generate other company... ', end='')
        from human_resource_utils.other_company import OtherCompany
        other_company = OtherCompany(human_resource_workbook=self._human_resource_workbook,
                                     month=self._month,
                                     employee_analyze=self._employee_analyze,
                                     real_check_in_and_change_analyze=self._real_check_in_and_change_analyze,
                                     real_leave_analyze=self._real_leave_analyze,
                                     expect_check_in_leave_analyze=self._expect_check_in_leave_excel_path)
        print('done.')


    def _create_dreamers(self):
        print('Generate dreamers... ', end='')
        from human_resource_utils.dreamers import Dreamers
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




class ConvertThreads(QThread):
    _leave_list_excel_path = None
    _employee_excel_path = None
    _leave_employee_excel_path = None
    _save_path = None

    callback = QtCore.pyqtSignal(str)

    def __init__(self,
                 leave_list_excel_path,
                 employee_excel_path,
                 leave_employee_excel_path,
                 save_path):
        self._leave_list_excel_path = leave_list_excel_path
        self._employee_excel_path = employee_excel_path
        self._leave_employee_excel_path = leave_employee_excel_path
        self._save_path = save_path
        super().__init__()

    def run(self) -> None:
        try:
            from human_resource_utils.leave_list_add_id_converter import LeaveListAddIDConverter
            converter = LeaveListAddIDConverter(leave_list_excel_path=self._leave_list_excel_path,
                                                employee_excel_path=self._employee_excel_path,
                                                leave_employee_excel_path=self._leave_employee_excel_path,
                                                new_leave_list_excel_path=self._save_path)
            converter.convert()
            self.callback.emit('轉換成功\n儲存於: {}'.format(self._save_path))

        except Exception as e:
            print('Convert error: {}\n{}'.format(e, traceback.format_exc()))
            self.callback.emit('Convert error: {}'.format(e))



class GenerateThreads(QThread):
    _month = None
    _start_date = None
    _end_date = None
    _employee_excel_path = None
    _real_check_in_and_change_excel_path = None
    _real_leave_excel_path = None
    _expect_check_in_leave_excel_path = None
    _save_path = None

    callback = QtCore.pyqtSignal(str)

    def __init__(self, month, start_date, end_date,
                 employee_excel_path,
                 real_check_in_and_change_excel_path,
                 real_leave_excel_path,
                 expect_check_in_leave_excel_path,
                 save_path):
        self._month = month
        self._start_date = start_date
        self._end_date = end_date
        self._employee_excel_path = employee_excel_path
        self._real_check_in_and_change_excel_path = real_check_in_and_change_excel_path
        self._real_leave_excel_path = real_leave_excel_path
        self._expect_check_in_leave_excel_path = expect_check_in_leave_excel_path
        self._save_path = save_path
        super().__init__()

    def run(self):
        try:
            human_resource = HumanResource(month=self._month, start_date=self._start_date, end_date=self._end_date,
                                           employee_excel_path=self._employee_excel_path,
                                           real_check_in_and_change_excel_path=self._real_check_in_and_change_excel_path,
                                           real_leave_excel_path=self._real_leave_excel_path,
                                           expect_check_in_leave_excel_path=self._expect_check_in_leave_excel_path)
            human_resource.save(save_path=self._save_path)
            self.callback.emit('產生成功\n儲存於: {}'.format(self._save_path))

        except Exception as e:
            print('Generate error: {}\n{}'.format(e, traceback.format_exc()))
            self.callback.emit('Generate error: {}'.format(e))



class HumanResourceUI(QMainWindow, Ui_MainWindow):

    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.setWindowFlags((self.windowFlags()
                             & ~QtCore.Qt.WindowMaximizeButtonHint
                             & ~QtCore.Qt.WindowFullscreenButtonHint)
                            | QtCore.Qt.CustomizeWindowHint)
        self.setWindowIcon(QtGui.QIcon('./icon/miopane.png'))

        self._init_version()
        self._init_date()
        self._set_button_click()

        # self._init_default()


    def _init_version(self):
        self.label_version.setText('Version: {}'.format(__version__))


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


    def _init_default(self):
        self.lineEdit_employee_value.setText('/Users/linmeimei/Desktop/Linmeimei/Github/Miopane/Document3/目前人力配置.xlsx')
        self.lineEdit_real_check_in_and_change_value.setText('/Users/linmeimei/Desktop/Linmeimei/Github/Miopane/Document3/實際入職+單位異動.xlsx')
        self.lineEdit_real_leave_value.setText('/Users/linmeimei/Desktop/Linmeimei/Github/Miopane/Document3/3月離職(New).xlsx')
        self.lineEdit_manager_report_value.setText('/Users/linmeimei/Desktop/Linmeimei/Github/Miopane/Document3/預計報到離職+單位主管回報人力缺額.xlsx')
        self.lineEdit_save_value.setText('/Users/linmeimei/Desktop/Linmeimei/Github/Miopane/Document3')


    def _convert(self):
        real_leave_excel_path = self.lineEdit_converter_real_leave_value.text()
        employee_excel_path = self.lineEdit_converter_employee_value.text()
        leave_list_excel_path = self.lineEdit_converter_leave_list_value.text()
        save_dir = self.lineEdit_converter_save_value.text() if self.lineEdit_converter_save_value.text() != '' else './'
        save_path = os.path.join(save_dir, os.path.basename(real_leave_excel_path).replace('.xlsx', '(New).xlsx'))

        if real_leave_excel_path != '' and employee_excel_path != '' and leave_list_excel_path != '':
            self.pushButton_convert.setEnabled(False)
            self._print_log(log='實際離職: {}\n目前人力配置: {}\n離職名單: {}\n'.format(real_leave_excel_path, employee_excel_path, leave_list_excel_path))
            self._print_log(log='開始轉換...')

            self.convert_threads = ConvertThreads(leave_list_excel_path=real_leave_excel_path,
                                                  employee_excel_path=employee_excel_path,
                                                  leave_employee_excel_path=leave_list_excel_path,
                                                  save_path=save_path)
            self.convert_threads.callback.connect(self._convert_thread_callback)
            self.convert_threads.start()
        else:
            self._print_log(log='Error: 有檔案欄位為空!')

    def _convert_thread_callback(self, log):
        self.pushButton_convert.setEnabled(True)
        self._print_log(log=log)

    def _generate(self):
        month = self.comboBox_month.currentText()
        start_date = self.dateEdit_start.dateTime().toPyDateTime()
        end_date = self.dateEdit_end.dateTime().toPyDateTime()
        employee_excel_path = self.lineEdit_employee_value.text()
        real_check_in_and_change_excel_path = self.lineEdit_real_check_in_and_change_value.text()
        real_leave_excel_path = self.lineEdit_real_leave_value.text()
        expect_check_in_leave_excel_path = self.lineEdit_manager_report_value.text()

        save_dir = self.lineEdit_save_value.text() if self.lineEdit_save_value.text() != '' else './'
        save_path = os.path.join(save_dir, '人力資源報表.xlsx')

        if employee_excel_path != '' and real_check_in_and_change_excel_path != '' and real_leave_excel_path != '' and expect_check_in_leave_excel_path != '':
            self.pushButton_generate.setEnabled(False)
            self._print_log(log='統計月份: {}\n'
                                '統計區間: {} ~ {}\n'
                                '目前人力配置: {}\n'
                                '實際入職+單位異動: {}\n'
                                '實際離職: {}\n'
                                '預計報到離職+單位主管回報人力缺額: {}\n'
                            .format(month,
                                    start_date, end_date,
                                    employee_excel_path,
                                    real_check_in_and_change_excel_path,
                                    real_leave_excel_path,
                                    expect_check_in_leave_excel_path))
            self._print_log(log='開始產生...')

            self.generate_threads = GenerateThreads(month=month,
                                                    start_date=start_date.strftime('%Y/%m/%d'),
                                                    end_date=end_date.strftime('%Y/%m/%d'),
                                                    employee_excel_path=employee_excel_path,
                                                    real_check_in_and_change_excel_path=real_check_in_and_change_excel_path,
                                                    real_leave_excel_path=real_leave_excel_path,
                                                    expect_check_in_leave_excel_path=expect_check_in_leave_excel_path,
                                                    save_path=save_path)
            self.generate_threads.callback.connect(self._generate_thread_callback)
            self.generate_threads.start()
        else:
            self._print_log(log='Error: 有檔案欄位為空!')

    def _generate_thread_callback(self, log):
        self.pushButton_generate.setEnabled(True)
        self._print_log(log=log)

    def _init_date(self):
        self.comboBox_month.setCurrentIndex(datetime.now().month - 1)
        self.dateEdit_start.setDate(QDate.currentDate().addDays(-14))
        self.dateEdit_end.setDate(QDate.currentDate())

    def _button_converter_real_leave_select_file_click(self):
        file_path, _ = QFileDialog.getOpenFileName(self)
        if file_path != '':
            self.lineEdit_converter_real_leave_value.setText(file_path)

    def _button_converter_employee_select_file_click(self):
        file_path, _ = QFileDialog.getOpenFileName(self)
        if file_path != '':
            self.lineEdit_converter_employee_value.setText(file_path)

    def _button_converter_leave_list_select_file_click(self):
        file_path, _ = QFileDialog.getOpenFileName(self)
        if file_path != '':
            self.lineEdit_converter_leave_list_value.setText(file_path)

    def _button_converter_save_select_directory_click(self):
        dir_path = QFileDialog.getExistingDirectory(self)
        if dir_path != '':
            self.lineEdit_converter_save_value.setText(dir_path)

    def _button_employee_select_file_click(self):
        file_path, _ = QFileDialog.getOpenFileName(self)
        if file_path != '':
            self.lineEdit_employee_value.setText(file_path)

    def _button_real_check_in_and_change_select_file_click(self):
        file_path, _ = QFileDialog.getOpenFileName(self)
        if file_path != '':
            self.lineEdit_real_check_in_and_change_value.setText(file_path)

    def _button_real_leave_select_file_click(self):
        file_path, _ = QFileDialog.getOpenFileName(self)
        if file_path != '':
            self.lineEdit_real_leave_value.setText(file_path)

    def _button_manager_report_select_file_click(self):
        file_path, _ = QFileDialog.getOpenFileName(self)
        if file_path != '':
            self.lineEdit_manager_report_value.setText(file_path)

    def _button_save_select_directory_click(self):
        dir_path = QFileDialog.getExistingDirectory(self)
        if dir_path != '':
            self.lineEdit_save_value.setText(dir_path)

    def _print_log(self, log):
        self.textBrowser_debug_message.append('\n' + log)
        self.textBrowser_debug_message.verticalScrollBar().setValue(self.textBrowser_debug_message.verticalScrollBar().maximum())
        # self.textBrowser_debug_message.update()
        # QtWidgets.QApplication.processEvents(QEventLoop.AllEvents)




if __name__ == '__main__':
    print('Version: {}'.format(__version__))
    import qdarktheme
    import qdarkstyle
    qdarktheme.enable_hi_dpi()

    app = QtWidgets.QApplication(sys.argv)
    # app.setStyleSheet(qdarkstyle.load_stylesheet_pyqt5())
    # qdarktheme.setup_theme('auto')

    window = HumanResourceUI()
    window.show()

    sys.exit(app.exec_())
