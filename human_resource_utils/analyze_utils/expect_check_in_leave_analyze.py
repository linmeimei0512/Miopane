from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from datetime import datetime, timedelta

class ExpectCheckInLeaveAnalyze:
    _workbook: Workbook = None
    _worksheet: Worksheet = None
    _excel_path = None

    # date range
    _start_date_str = None
    _end_date_str = None
    _start_date = None
    _end_date = None

    _data_list = []

    def __init__(self, excel_path, start_date_str=None, end_date_str=None):
        self._excel_path = excel_path
        self._start_date_str = start_date_str
        self._end_date_str = end_date_str

        self._init_data()
        self._init_range_of_date()
        self._read_excel()


    def _init_data(self):
        self._data_list = []


    def _init_range_of_date(self):
        if self._start_date_str is not None and self._end_date_str is not None:
            self._start_date = datetime.strptime(self._start_date_str, '%Y/%m/%d')
            self._end_date = datetime.strptime(self._end_date_str, '%Y/%m/%d')

        else:
            self._end_date = datetime.now()
            self._start_date = self._end_date - timedelta(weeks=2)

        print('date: {} ~ {}'.format(self._start_date, self._end_date))


    def _read_excel(self):
        self._workbook = load_workbook(self._excel_path)
        self._worksheet = self._workbook.worksheets[0]

        rows_data = list(self._worksheet.rows)
        self._title_list = [title.value for title in rows_data.pop(0)]

        for row in rows_data:
            data = [cell.value for cell in row]
            if len(data) > 0 and data[0] is not None:
                self._data_list.append(dict(zip(self._title_list, data)))

        self._data_list = list(filter(lambda data: data['時間戳記'] > self._start_date and data['時間戳記'] < self._end_date, self._data_list))

        # for data in self._data_list:
        #     print(data)


    def search_by_group_name(self, search_group_name):
        department_group_data_list = []

        if type(search_group_name) == str:
            department_group_data_list = list(filter(lambda data: data['單位'] == search_group_name, self._data_list))

        elif type(search_group_name) == list:
            for name in search_group_name:
                department_group_data_list += list(filter(lambda data: data['單位'] == name, self._data_list))

        return department_group_data_list




if __name__ == '__main__':
    expect_check_in_leave_analyze = ExpectCheckInLeaveAnalyze(excel_path='../../Document3/預計報到離職+單位主管回報人力缺額.xlsx',
                                                              start_date_str='2024/03/01',
                                                              end_date_str='2024/03/18')
    employee_list = expect_check_in_leave_analyze.search_by_group_name(search_group_name='Miopane台中內場')
    print(employee_list)
    # expect_check_in_leave_analyze = ExpectCheckInLeaveAnalyze(excel_path='../../Document/預計報到離職+單位主管回報人力缺額.xlsx')