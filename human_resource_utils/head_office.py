from openpyxl import Workbook

from human_resource_utils.company_sheet import CompanySheet, Company, Department, DepartmentGroup
from human_resource_utils.analyze_utils.employee_analyze import EmployeeAnalyze
from human_resource_utils.analyze_utils.real_check_in_and_change_analyze import ReadCheckInAndChangeAnalyze
from human_resource_utils.analyze_utils.real_leave_analyze import RealLeaveAnalyze
from human_resource_utils.analyze_utils.expect_check_in_leave_analyze import ExpectCheckInLeaveAnalyze
from utils.dictionary_key import DictionaryKey


class HeadOffice(CompanySheet):
    # root cell list
    _root_cell_list = [{DictionaryKey.VALUE: '品牌:中央廚房\n人力資源報表', DictionaryKey.START: 'A1', DictionaryKey.END: 'C1', DictionaryKey.CENTER: True, DictionaryKey.HEIGHT: 50, DictionaryKey.BOLD: True, DictionaryKey.THIN_BORDER: True},
                       {DictionaryKey.VALUE: '月', DictionaryKey.START: 'D1', DictionaryKey.END: 'F1', DictionaryKey.CENTER: True, DictionaryKey.BOLD: True, DictionaryKey.THIN_BORDER: True},
                       {DictionaryKey.VALUE: '單位異動', DictionaryKey.START: 'K1', DictionaryKey.END: 'L1', DictionaryKey.CENTER: True, DictionaryKey.BOLD: True, DictionaryKey.THIN_BORDER: True},
                       {DictionaryKey.VALUE: '預計', DictionaryKey.START: 'P1', DictionaryKey.END: 'Q1', DictionaryKey.CENTER: True, DictionaryKey.BOLD: True, DictionaryKey.THIN_BORDER: True, DictionaryKey.BACKGROUND_COLOR: 'FFFFCC'}]

    # common cell list
    _common_cell_list = [{DictionaryKey.VALUE: '部門', DictionaryKey.START: 'A2', DictionaryKey.END: 'A3', DictionaryKey.CENTER: True, DictionaryKey.WIDTH: 22, DictionaryKey.THIN_BORDER: True},
                         {DictionaryKey.VALUE: '單位', DictionaryKey.START: 'B2', DictionaryKey.END: 'B3', DictionaryKey.CENTER: True, DictionaryKey.THIN_BORDER: True},
                         {DictionaryKey.VALUE: '預估人力配置', DictionaryKey.START: 'C2', DictionaryKey.END: 'D2', DictionaryKey.CENTER: True, DictionaryKey.THIN_BORDER: True, DictionaryKey.BACKGROUND_COLOR: 'DDEBF7'},
                         {DictionaryKey.VALUE: '正職', DictionaryKey.START: 'C3', DictionaryKey.CENTER: True, DictionaryKey.THIN_BORDER: True},
                         {DictionaryKey.VALUE: '工讀生', DictionaryKey.START: 'D3', DictionaryKey.CENTER: True, DictionaryKey.THIN_BORDER: True},
                         {DictionaryKey.VALUE: '目前人力配置', DictionaryKey.START: 'E2', DictionaryKey.END: 'F2', DictionaryKey.CENTER: True, DictionaryKey.THIN_BORDER: True, DictionaryKey.BACKGROUND_COLOR: 'DDEBF7'},
                         {DictionaryKey.VALUE: '正職', DictionaryKey.START: 'E3', DictionaryKey.CENTER: True, DictionaryKey.THIN_BORDER: True},
                         {DictionaryKey.VALUE: '工讀生', DictionaryKey.START: 'F3', DictionaryKey.CENTER: True, DictionaryKey.THIN_BORDER: True},
                         {DictionaryKey.VALUE: '人力溢 / 缺額', DictionaryKey.START: 'G2', DictionaryKey.END: 'H2', DictionaryKey.CENTER: True, DictionaryKey.THIN_BORDER: True, DictionaryKey.BACKGROUND_COLOR: 'DDEBF7'},
                         {DictionaryKey.VALUE: '正職', DictionaryKey.START: 'G3', DictionaryKey.CENTER: True, DictionaryKey.THIN_BORDER: True},
                         {DictionaryKey.VALUE: '工讀生', DictionaryKey.START: 'H3', DictionaryKey.CENTER: True, DictionaryKey.THIN_BORDER: True},
                         {DictionaryKey.VALUE: '月初總人數', DictionaryKey.START: 'I2', DictionaryKey.END: 'I3', DictionaryKey.CENTER: True, DictionaryKey.WIDTH: 14, DictionaryKey.THIN_BORDER: True, DictionaryKey.BACKGROUND_COLOR: 'DDEBF7'},
                         {DictionaryKey.VALUE: '當月入職人數', DictionaryKey.START: 'J2', DictionaryKey.END: 'J3', DictionaryKey.CENTER: True, DictionaryKey.WIDTH: 14, DictionaryKey.THIN_BORDER: True, DictionaryKey.BACKGROUND_COLOR: 'DDEBF7'},
                         {DictionaryKey.VALUE: '轉入單位人數', DictionaryKey.START: 'K2', DictionaryKey.END: 'K3', DictionaryKey.CENTER: True, DictionaryKey.WIDTH: 14, DictionaryKey.THIN_BORDER: True, DictionaryKey.BACKGROUND_COLOR: 'DDEBF7'},
                         {DictionaryKey.VALUE: '轉出單位人數', DictionaryKey.START: 'L2', DictionaryKey.END: 'L3', DictionaryKey.CENTER: True, DictionaryKey.WIDTH: 14, DictionaryKey.THIN_BORDER: True, DictionaryKey.BACKGROUND_COLOR: 'DDEBF7'},
                         {DictionaryKey.VALUE: '當月離職人數', DictionaryKey.START: 'M2', DictionaryKey.END: 'M3', DictionaryKey.CENTER: True, DictionaryKey.WIDTH: 14, DictionaryKey.THIN_BORDER: True, DictionaryKey.BACKGROUND_COLOR: 'DDEBF7'},
                         {DictionaryKey.VALUE: '月末總人數', DictionaryKey.START: 'N2', DictionaryKey.END: 'N3', DictionaryKey.CENTER: True, DictionaryKey.WIDTH: 14, DictionaryKey.THIN_BORDER: True, DictionaryKey.BACKGROUND_COLOR: 'DDEBF7'},
                         {DictionaryKey.VALUE: '處理情形', DictionaryKey.START: 'O2', DictionaryKey.END: 'O3', DictionaryKey.CENTER: True, DictionaryKey.WIDTH: 14, DictionaryKey.THIN_BORDER: True},
                         {DictionaryKey.VALUE: '報到', DictionaryKey.START: 'P2', DictionaryKey.END: 'P3', DictionaryKey.CENTER: True, DictionaryKey.BOLD: True, DictionaryKey.TEXT_COLOR: 'FF0000', DictionaryKey.WIDTH: 14, DictionaryKey.THIN_BORDER: True, DictionaryKey.BACKGROUND_COLOR: 'FFFFCC'},
                         {DictionaryKey.VALUE: '離職', DictionaryKey.START: 'Q2', DictionaryKey.END: 'Q3', DictionaryKey.CENTER: True, DictionaryKey.BOLD: True, DictionaryKey.TEXT_COLOR: 'FF0000', DictionaryKey.WIDTH: 14, DictionaryKey.THIN_BORDER: True, DictionaryKey.BACKGROUND_COLOR: 'FFFFCC'},
                         {DictionaryKey.VALUE: '備註', DictionaryKey.START: 'R2', DictionaryKey.END: 'R3', DictionaryKey.CENTER: True, DictionaryKey.WIDTH: 14, DictionaryKey.THIN_BORDER: True},
                         {DictionaryKey.VALUE: '離職人員', DictionaryKey.START: 'S2', DictionaryKey.END: 'S3', DictionaryKey.CENTER: True, DictionaryKey.THIN_BORDER: True},
                         {DictionaryKey.VALUE: '年資\n單位：年', DictionaryKey.START: 'T2', DictionaryKey.END: 'T3', DictionaryKey.CENTER: True, DictionaryKey.THIN_BORDER: True},
                         {DictionaryKey.VALUE: '正職', DictionaryKey.START: 'V2', DictionaryKey.END: 'V3', DictionaryKey.CENTER: True, DictionaryKey.THIN_BORDER: True},
                         {DictionaryKey.VALUE: '工讀生', DictionaryKey.START: 'W2', DictionaryKey.END: 'W3', DictionaryKey.CENTER: True, DictionaryKey.THIN_BORDER: True}]

    _common_prompt_cell_list = [{DictionaryKey.VALUE: '預估人力配置：依照各店平均業績所預估之人力', DictionaryKey.START: 'A23', DictionaryKey.END: 'S23', DictionaryKey.THIN_BORDER: True},
                                {DictionaryKey.VALUE: '目前人力配置：當月月底之在職人數=月初總人數＋當月入職人數 - 當月離職人數', DictionaryKey.START: 'A24', DictionaryKey.END: 'S24', DictionaryKey.THIN_BORDER: True},
                                {DictionaryKey.VALUE: '月初總人數：當月1號在職人數，『不包含』1號到職之團員', DictionaryKey.START: 'A25', DictionaryKey.END: 'S25', DictionaryKey.THIN_BORDER: True},
                                {DictionaryKey.VALUE: '當月入職人數：全月到職之團員人數', DictionaryKey.START: 'A26', DictionaryKey.END: 'S26', DictionaryKey.THIN_BORDER: True},
                                {DictionaryKey.VALUE: '當月離職人數：全月離職之團員人數', DictionaryKey.START: 'A27', DictionaryKey.END: 'S27', DictionaryKey.THIN_BORDER: True},
                                {DictionaryKey.VALUE: '當月離職率：當月離職人數 / (月初總人數＋當月入職人數)＊％', DictionaryKey.START: 'A28', DictionaryKey.END: 'S28', DictionaryKey.THIN_BORDER: True}]

    # company
    _company = Company(name='中央廚房', sheet_name='總公司')
    _company.department_list = [Department(name='總經理', department_group_list=[DepartmentGroup(name='總經理', search_name='總經理', number_expect_full_time_employee=1)]),
                                Department(name='中央財務部', department_group_list=[DepartmentGroup(name='會計組', search_name=['財務部會計組', '財務部投資組'], number_expect_full_time_employee=2),
                                                                                     DepartmentGroup(name='採購組', search_name='財務部採購組', number_expect_full_time_employee=3),
                                                                                     DepartmentGroup(name='總務組', search_name='財務部總務組', number_expect_full_time_employee=2)]),
                                Department(name='中央管理部', department_group_list=[DepartmentGroup(name='人事組', search_name=['管理部人事組', '管理部'], number_expect_full_time_employee=3)]),
                                Department(name='中央拓展部', department_group_list=[DepartmentGroup(name='拓展組', search_name='拓展部', number_expect_full_time_employee=1),
                                                                                     DepartmentGroup(name='工程組', search_name='拓展部工程組', number_expect_full_time_employee=3),
                                                                                     DepartmentGroup(name='開發企劃組', search_name='拓展部開發企劃組', number_expect_full_time_employee=4)]),
                                Department(name='中央營運部', department_group_list=[DepartmentGroup(name='', search_name='營運部', number_expect_full_time_employee=8)]),
                                Department(name='中央醬料', department_group_list=[DepartmentGroup(name='', search_name='中央廚房醬料組', number_expect_full_time_employee=22, number_expect_part_time_employee=0)]),
                                Department(name='中央麵包', department_group_list=[DepartmentGroup(name='', search_name='中央廚房麵包組', number_expect_full_time_employee=34, number_expect_part_time_employee=0)]),
                                Department(name='中央西點', department_group_list=[DepartmentGroup(name='', search_name='中央廚房西點組', number_expect_full_time_employee=8, number_expect_part_time_employee=0)]),
                                Department(name='中央包裝', department_group_list=[DepartmentGroup(name='', search_name='中央廚房宅配、包裝組', number_expect_full_time_employee=3, number_expect_part_time_employee=2)]),
                                Department(name='廚房物流', department_group_list=[DepartmentGroup(name='', search_name='中央廚房物流組', number_expect_full_time_employee=8, number_expect_part_time_employee=0)]),
                                Department(name='宅配物流', department_group_list=[DepartmentGroup(name='', search_name='中央宅配物流組', number_expect_full_time_employee=1, number_expect_part_time_employee=0)])]


    def __init__(self, human_resource_workbook: Workbook, month,
                 employee_analyze: EmployeeAnalyze,
                 real_check_in_and_change_analyze: ReadCheckInAndChangeAnalyze,
                 real_leave_analyze: RealLeaveAnalyze,
                 expect_check_in_leave_analyze: ExpectCheckInLeaveAnalyze):
        self._root_cell_list[0][DictionaryKey.VALUE] = '品牌:' + self._company.name + '\n人力資源報表'
        self._root_cell_list[1][DictionaryKey.VALUE] = str(month) + '月'
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
        super()._write_total(start_column='C', end_column='N')