from openpyxl import Workbook

from company_sheet import CompanySheet, Company, Department, DepartmentGroup
from analyze_utils.employee_analyze import EmployeeAnalyze
from analyze_utils.real_check_in_and_change_analyze import ReadCheckInAndChangeAnalyze
from analyze_utils.real_leave_analyze import RealLeaveAnalyze
from analyze_utils.expect_check_in_leave_analyze import ExpectCheckInLeaveAnalyze
from utils.dictionary_key import DictionaryKey


class Dreamers(CompanySheet):
    # dreamers company
    _company_sheet_name = 'Dreamers'

    # root cell list
    _root_cell_list = [{DictionaryKey.VALUE: '品牌:Dreamers coffee\n人力資源報表', DictionaryKey.START: 'A1', DictionaryKey.END: 'C1', DictionaryKey.CENTER: True, DictionaryKey.HEIGHT: 50, DictionaryKey.BOLD: True, DictionaryKey.THIN_BORDER: True},
                       {DictionaryKey.VALUE: '月', DictionaryKey.START: 'D1', DictionaryKey.END: 'F1', DictionaryKey.CENTER: True, DictionaryKey.BOLD: True, DictionaryKey.THIN_BORDER: True},
                       {DictionaryKey.VALUE: '單位異動', DictionaryKey.START: 'L1', DictionaryKey.END: 'M1', DictionaryKey.CENTER: True, DictionaryKey.BOLD: True, DictionaryKey.THIN_BORDER: True},
                       {DictionaryKey.VALUE: '預計', DictionaryKey.START: 'Q1', DictionaryKey.END: 'R1', DictionaryKey.CENTER: True, DictionaryKey.BOLD: True, DictionaryKey.THIN_BORDER: True, DictionaryKey.BACKGROUND_COLOR: 'FFFFCC'}]

    # common cell list
    _common_cell_list = [{DictionaryKey.VALUE: '部門', DictionaryKey.START: 'A2', DictionaryKey.END: 'A3', DictionaryKey.CENTER: True, DictionaryKey.WIDTH: 22, DictionaryKey.THIN_BORDER: True},
                         {DictionaryKey.VALUE: '單位', DictionaryKey.START: 'B2', DictionaryKey.END: 'B3', DictionaryKey.CENTER: True, DictionaryKey.THIN_BORDER: True},
                         {DictionaryKey.VALUE: '當月業績', DictionaryKey.START: 'C2', DictionaryKey.END: 'C3', DictionaryKey.CENTER: True, DictionaryKey.THIN_BORDER: True, DictionaryKey.BACKGROUND_COLOR: 'FCE4D6'},
                         {DictionaryKey.VALUE: '預估人力配置', DictionaryKey.START: 'D2', DictionaryKey.END: 'E2', DictionaryKey.CENTER: True, DictionaryKey.THIN_BORDER: True, DictionaryKey.BACKGROUND_COLOR: 'DDEBF7'},
                         {DictionaryKey.VALUE: '正職', DictionaryKey.START: 'D3', DictionaryKey.CENTER: True, DictionaryKey.THIN_BORDER: True},
                         {DictionaryKey.VALUE: '工讀生', DictionaryKey.START: 'E3', DictionaryKey.CENTER: True, DictionaryKey.THIN_BORDER: True},
                         {DictionaryKey.VALUE: '目前人力配置', DictionaryKey.START: 'F2', DictionaryKey.END: 'G2', DictionaryKey.CENTER: True, DictionaryKey.THIN_BORDER: True, DictionaryKey.BACKGROUND_COLOR: 'DDEBF7'},
                         {DictionaryKey.VALUE: '正職', DictionaryKey.START: 'F3', DictionaryKey.CENTER: True, DictionaryKey.THIN_BORDER: True},
                         {DictionaryKey.VALUE: '工讀生', DictionaryKey.START: 'G3', DictionaryKey.CENTER: True, DictionaryKey.THIN_BORDER: True},
                         {DictionaryKey.VALUE: '人力溢 / 缺額', DictionaryKey.START: 'H2', DictionaryKey.END: 'I2', DictionaryKey.CENTER: True, DictionaryKey.THIN_BORDER: True, DictionaryKey.BACKGROUND_COLOR: 'DDEBF7'},
                         {DictionaryKey.VALUE: '正職', DictionaryKey.START: 'H3', DictionaryKey.CENTER: True, DictionaryKey.THIN_BORDER: True},
                         {DictionaryKey.VALUE: '工讀生', DictionaryKey.START: 'I3', DictionaryKey.CENTER: True, DictionaryKey.THIN_BORDER: True},
                         {DictionaryKey.VALUE: '月初總人數', DictionaryKey.START: 'J2', DictionaryKey.END: 'J3', DictionaryKey.CENTER: True, DictionaryKey.WIDTH: 14, DictionaryKey.THIN_BORDER: True, DictionaryKey.BACKGROUND_COLOR: 'DDEBF7'},
                         {DictionaryKey.VALUE: '當月入職人數', DictionaryKey.START: 'K2', DictionaryKey.END: 'K3', DictionaryKey.CENTER: True, DictionaryKey.WIDTH: 14, DictionaryKey.THIN_BORDER: True, DictionaryKey.BACKGROUND_COLOR: 'DDEBF7'},
                         {DictionaryKey.VALUE: '轉入單位人數', DictionaryKey.START: 'L2', DictionaryKey.END: 'L3', DictionaryKey.CENTER: True, DictionaryKey.WIDTH: 14, DictionaryKey.THIN_BORDER: True, DictionaryKey.BACKGROUND_COLOR: 'DDEBF7'},
                         {DictionaryKey.VALUE: '轉出單位人數', DictionaryKey.START: 'M2', DictionaryKey.END: 'M3', DictionaryKey.CENTER: True, DictionaryKey.WIDTH: 14, DictionaryKey.THIN_BORDER: True, DictionaryKey.BACKGROUND_COLOR: 'DDEBF7'},
                         {DictionaryKey.VALUE: '當月離職人數', DictionaryKey.START: 'N2', DictionaryKey.END: 'N3', DictionaryKey.CENTER: True, DictionaryKey.WIDTH: 14, DictionaryKey.THIN_BORDER: True, DictionaryKey.BACKGROUND_COLOR: 'DDEBF7'},
                         {DictionaryKey.VALUE: '月末總人數', DictionaryKey.START: 'O2', DictionaryKey.END: 'O3', DictionaryKey.CENTER: True, DictionaryKey.WIDTH: 14, DictionaryKey.THIN_BORDER: True, DictionaryKey.BACKGROUND_COLOR: 'DDEBF7'},
                         {DictionaryKey.VALUE: '處理情形', DictionaryKey.START: 'P2', DictionaryKey.END: 'P3', DictionaryKey.CENTER: True, DictionaryKey.WIDTH: 14, DictionaryKey.THIN_BORDER: True},
                         {DictionaryKey.VALUE: '報到', DictionaryKey.START: 'Q2', DictionaryKey.END: 'Q3', DictionaryKey.CENTER: True, DictionaryKey.BOLD: True, DictionaryKey.TEXT_COLOR: 'FF0000', DictionaryKey.WIDTH: 14, DictionaryKey.THIN_BORDER: True, DictionaryKey.BACKGROUND_COLOR: 'FFFFCC'},
                         {DictionaryKey.VALUE: '離職', DictionaryKey.START: 'R2', DictionaryKey.END: 'R3', DictionaryKey.CENTER: True, DictionaryKey.BOLD: True, DictionaryKey.TEXT_COLOR: 'FF0000', DictionaryKey.WIDTH: 14, DictionaryKey.THIN_BORDER: True, DictionaryKey.BACKGROUND_COLOR: 'FFFFCC'},
                         {DictionaryKey.VALUE: '備註', DictionaryKey.START: 'S2', DictionaryKey.END: 'S3', DictionaryKey.CENTER: True, DictionaryKey.WIDTH: 14, DictionaryKey.THIN_BORDER: True},
                         {DictionaryKey.VALUE: '離職人員', DictionaryKey.START: 'T2', DictionaryKey.END: 'T3', DictionaryKey.CENTER: True, DictionaryKey.THIN_BORDER: True},
                         {DictionaryKey.VALUE: '年資\n單位：年', DictionaryKey.START: 'U2', DictionaryKey.END: 'U3', DictionaryKey.CENTER: True, DictionaryKey.THIN_BORDER: True},
                         {DictionaryKey.VALUE: '單位主管回報\n人力缺額', DictionaryKey.START: 'W2', DictionaryKey.END: 'X2', DictionaryKey.CENTER: True, DictionaryKey.THIN_BORDER: True, DictionaryKey.BACKGROUND_COLOR: 'DDEBF7'},
                         {DictionaryKey.VALUE: '正職', DictionaryKey.START: 'W3', DictionaryKey.CENTER: True, DictionaryKey.THIN_BORDER: True},
                         {DictionaryKey.VALUE: '工讀生', DictionaryKey.START: 'X3', DictionaryKey.CENTER: True, DictionaryKey.THIN_BORDER: True}]

    _common_prompt_cell_list = [{DictionaryKey.VALUE: '預估人力配置：依照各店平均業績所預估之人力', DictionaryKey.START: 'A23', DictionaryKey.END: 'S23', DictionaryKey.THIN_BORDER: True},
                                {DictionaryKey.VALUE: '目前人力配置：當月月底之在職人數=月初總人數＋當月入職人數 - 當月離職人數', DictionaryKey.START: 'A24', DictionaryKey.END: 'S24', DictionaryKey.THIN_BORDER: True},
                                {DictionaryKey.VALUE: '月初總人數：當月1號在職人數，『不包含』1號到職之團員', DictionaryKey.START: 'A25', DictionaryKey.END: 'S25', DictionaryKey.THIN_BORDER: True},
                                {DictionaryKey.VALUE: '當月入職人數：全月到職之團員人數', DictionaryKey.START: 'A26', DictionaryKey.END: 'S26', DictionaryKey.THIN_BORDER: True},
                                {DictionaryKey.VALUE: '當月離職人數：全月離職之團員人數', DictionaryKey.START: 'A27', DictionaryKey.END: 'S27', DictionaryKey.THIN_BORDER: True},
                                {DictionaryKey.VALUE: '當月離職率：當月離職人數 / (月初總人數＋當月入職人數)＊％', DictionaryKey.START: 'A28', DictionaryKey.END: 'S28', DictionaryKey.THIN_BORDER: True}]

    # company
    _company = Company(name='Dreamers coffee', sheet_name='Dreamers')
    _company.department_list = [Department(name='直營', department_group_list=[DepartmentGroup(name='內場', search_name='', number_expect_full_time_employee=1, number_expect_part_time_employee=0)]),
                                Department(name='忠誠', department_group_list=[DepartmentGroup(name='內場', search_name=['Dreamers忠誠內場', 'Dreamers忠誠'], number_expect_full_time_employee=4, number_expect_part_time_employee=4)]),
                                Department(name='101', department_group_list=[DepartmentGroup(name='內場', search_name=['Dreamers101內場', 'Dreamers101'], number_expect_full_time_employee=6, number_expect_part_time_employee=7)]),
                                Department(name='A8', department_group_list=[DepartmentGroup(name='內場', search_name=['DreamersA8內場', 'DreamersA8'], number_expect_full_time_employee=5, number_expect_part_time_employee=5)]),
                                Department(name='復興', department_group_list=[DepartmentGroup(name='內場', search_name=['Dreamers復興內場', 'Dreamers復興'], number_expect_full_time_employee=4, number_expect_part_time_employee=4)]),
                                Department(name='內湖', department_group_list=[DepartmentGroup(name='內場', search_name=['Dreamers內湖內場', 'Dreamers內湖'], number_expect_full_time_employee=4, number_expect_part_time_employee=4)]),
                                Department(name='光復', department_group_list=[DepartmentGroup(name='內場', search_name=['Dreamers光復內場', 'Dreamers光復'], number_expect_full_time_employee=4, number_expect_part_time_employee=4)]),
                                Department(name='永康', department_group_list=[DepartmentGroup(name='內場', search_name=['Dreamers永康店內場', 'Dreamers永康'], number_expect_full_time_employee=4, number_expect_part_time_employee=4)]),
                                Department(name='新店', department_group_list=[DepartmentGroup(name='內場', search_name=['Dreamers新店寶橋店內場', 'Dreamers新店寶橋'], number_expect_full_time_employee=4, number_expect_part_time_employee=4)]),
                                Department(name='微風', department_group_list=[DepartmentGroup(name='內場', search_name=['Dreamers微風復興店內場', 'Dreamers微風復興'], number_expect_full_time_employee=4, number_expect_part_time_employee=4)]),
                                Department(name='新莊', department_group_list=[DepartmentGroup(name='內場', search_name=['Dreamers新莊中平內場', 'Dreamers新莊中平'], number_expect_full_time_employee=3, number_expect_part_time_employee=2)]),
                                Department(name='南紡', department_group_list=[DepartmentGroup(name='內場', search_name=['Dreamers台南南紡內場', 'Dreamers南紡'], number_expect_full_time_employee=5, number_expect_part_time_employee=6)]),
                                Department(name='忠誠公園', department_group_list=[DepartmentGroup(name='內場', search_name=['Dreamers忠誠公園內場', 'Dreamers忠誠公園'], number_expect_full_time_employee=4, number_expect_part_time_employee=4)]),
                                Department(name='仁愛', department_group_list=[DepartmentGroup(name='內場', search_name='Dreamers仁愛安和內場', number_expect_full_time_employee=4, number_expect_part_time_employee=0),
                                                                               DepartmentGroup(name='外場', search_name='Dreamers仁愛安和外場', number_expect_full_time_employee=4, number_expect_part_time_employee=5)])]



    def __init__(self, human_resource_workbook: Workbook, month,
                 employee_analyze: EmployeeAnalyze,
                 real_check_in_and_change_analyze: ReadCheckInAndChangeAnalyze,
                 real_leave_analyze: RealLeaveAnalyze,
                 expect_check_in_leave_analyze: ExpectCheckInLeaveAnalyze):
        self._root_cell_list[0][DictionaryKey.VALUE] = '品牌:' + self._company.name + '\n人力資源報表'
        self._root_cell_list[1][DictionaryKey.VALUE] = str(month) + '月'
        self._department_value_start_column = 'D'
        super().__init__(human_resource_workbook=human_resource_workbook,
                         month=month,
                         employee_analyze=employee_analyze,
                         real_check_in_and_change_analyze=real_check_in_and_change_analyze,
                         real_leave_analyze=real_leave_analyze,
                         expect_check_in_leave_analyze=expect_check_in_leave_analyze)

        self._write_department_list()
        self._write_total()


    def _write_department_list(self):
        '''
        Write department list
        '''
        self._department_start_row = 4
        super()._write_department_list()


    def _write_total(self):
        '''
        Write total
        '''
        super()._write_total(start_column='C', end_column='O')