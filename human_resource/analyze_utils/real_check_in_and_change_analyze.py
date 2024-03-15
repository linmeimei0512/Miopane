from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet

class ReadCheckInAndChangeAnalyze:
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


    def search_change_list(self, search_group_name, in_or_out='in'):
        _change_list = list(filter(lambda employee: employee['異動行為'] == '單位異動', self._data_list))

        change_list = []
        if type(search_group_name) == str:
            if in_or_out == 'in':
                change_list = list(filter(lambda employee: employee['所屬單位(後)'] == search_group_name, _change_list))
            else:
                change_list = list(filter(lambda employee: employee['所屬單位(前)'] == search_group_name, _change_list))

        elif type(search_group_name) == list:
            for name in search_group_name:
                if in_or_out == 'in':
                    change_list += list(filter(lambda employee: employee['所屬單位(後)'] == name, _change_list))
                else:
                    change_list += list(filter(lambda employee: employee['所屬單位(前)'] == name, _change_list))

        # for change in change_list:
        #     print(change)

        return change_list


    def search_new_check_in_list(self, search_group_name):
        _check_in_list = list(filter(lambda employee: employee['異動行為'] == '新進' or '再雇用' in employee['異動行為'], self._data_list))

        check_in_list = []
        if type(search_group_name) == str:
            check_in_list = list(filter(lambda employee: employee['所屬單位(後)'] == search_group_name, _check_in_list))

        elif type(search_group_name) == list:
            for name in search_group_name:
                check_in_list += list(filter(lambda employee: employee['所屬單位(後)'] == name, _check_in_list))

        # for check_in in check_in_list:
        #     print(check_in)

        return check_in_list



if __name__ == '__main__':
    real_check_in_and_change_analyze = ReadCheckInAndChangeAnalyze(excel_path='../../Document/實際入職+單位異動.xlsx')
    # real_check_in_and_change_analyze.search_change_list(search_group_name='Miacucina信義A11外場')
    real_check_in_and_change_analyze.search_new_check_in_list(search_group_name='Dreamers仁愛安和外場')