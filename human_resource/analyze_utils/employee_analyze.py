from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet

class EmployeeAnalyze:
    _employee_workbook: Workbook = None
    _employee_worksheet: Worksheet = None
    _employee_excel_path = None

    _data_list = []

    def __init__(self, employee_excel_path):
        self._employee_excel_path = employee_excel_path
        self._data_list = []

        self._read_employee_excel()


    def _read_employee_excel(self):
        self._employee_workbook = load_workbook(self._employee_excel_path)
        self._employee_worksheet = self._employee_workbook.active

        rows_data = list(self._employee_worksheet.rows)
        self._title_list = [title.value for title in rows_data.pop(0)]

        for row in rows_data:
            data = [cell.value for cell in row]
            self._data_list.append(dict(zip(self._title_list, data)))


    def search_by_department_group(self, search_group_name, is_full_time=None):
        employee_list = []

        if type(search_group_name) == str:
            employee_list = list(filter(lambda employee: employee['所屬單位'] == search_group_name, self._data_list))

        elif type(search_group_name) == list:
            for name in search_group_name:
                employee_list += list(filter(lambda employee: employee['所屬單位'] == name, self._data_list))

        if is_full_time is not None:
            employee_list = list(filter(lambda employee: employee['身分類別'] == ('正式' if is_full_time else '非正式'), employee_list))

        return employee_list


    def search_by_name(self, name):
        employee_list = list(filter(lambda employee: employee['顯示名稱'] == name, self._data_list))
        return employee_list




if __name__ == '__main__':
    employee_analyze = EmployeeAnalyze(employee_excel_path='../../Document/目前人力配置.xlsx')
    print(len(employee_analyze.search_by_department_group(search_group_name='管理部人事組')))