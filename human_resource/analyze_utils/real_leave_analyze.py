from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet

class RealLeaveAnalyze:
    _workbook: Workbook = None
    _worksheet: Worksheet = None
    _excel_path = None

    _data_list = []

    def __init__(self, excel_path):
        self._excel_path = excel_path

        self._init_data()
        self._read_excel()


    def _init_data(self):
        self._data_list = []


    def _read_excel(self):
        self._workbook = load_workbook(self._excel_path)
        self._worksheet = self._workbook.worksheets[0]

        rows_data = list(self._worksheet.rows)
        self._title_list = [title.value for title in rows_data.pop(0)]

        for row in rows_data:
            data = [cell.value for cell in row]
            # print(data)
            self._data_list.append(dict(zip(self._title_list, data)))

        # print(self._data_list)


    def search_leave_list(self, search_group_name):
        _leave_list = []
        if type(search_group_name) == str:
            _leave_list = list(filter(lambda employee: employee['所屬單位'] == search_group_name, self._data_list))

        elif type(search_group_name) == list:
            for name in search_group_name:
                _leave_list += list(filter(lambda employee: employee['所屬單位'] == name, self._data_list))

        return _leave_list