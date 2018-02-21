from xlrd import open_workbook
import re

workbook = open_workbook('list_course.xls', formatting_info=True)
worksheet = workbook.sheet_by_index(0)

YEAR = 0
DEP_NUMBER = 1
DEP_NAME = 2
GROUP_NUMBER = 3
GROUP_NAME = 4
COURSE_NUMBER = 5
COURSE_NAME = 6
SIGHNED = 10
INSTRUCTOR_NAME = 13
DATES = 14


class Row:
    def __init__(self, start, end, row_values):
        self.start = start
        self.end = end
        self.row_values = row_values
        self.instructor = []
        self.dates = []
        self.extract_info()

    def extract_info(self):
        self._extract_year_info()
        self._extract_dep_info()
        self._extract_group_info()
        self._extract_course_info()
        self._extract_dates()

    def print_info(self):
        # print('Info for rows: {}:{}'.format(self.start, self.end))
        # print('Year:', self.year, 'term:', self.term)
        # print('Dep :', self.dep_number, self.dep_name)
        # print('Group :', self.group_number, self.group_name)
        # print('Course :', self.course_number, self.course_name, 'signed-up:', self.course_signed_up)
        # print('Instructors :', self.instructor)
        self._clean_dates()
        print('Dates:', self.dates)

    def _extract_year_info(self):
        for row in self.row_values:
            if row[YEAR].value:
                unclean = int(row[YEAR].value)
                self.term = unclean % 10
                self.year = (unclean + 10000) // 10
                break

    def _extract_dep_info(self):
        for row in self.row_values:
            if row[DEP_NUMBER].value:
                self.dep_number = int(row[DEP_NUMBER].value)
                break

        for row in self.row_values:
            if row[DEP_NAME].value:
                self.dep_name = row[DEP_NAME].value
                break

    def _extract_group_info(self):
        for row in self.row_values:
            if row[GROUP_NUMBER].value:
                self.group_number = int(row[GROUP_NUMBER].value)
                break

        for row in self.row_values:
            if row[GROUP_NAME].value:
                self.group_name = row[GROUP_NAME].value
                break

    def _extract_course_info(self):
        for row in self.row_values:
            if row[COURSE_NUMBER].value:
                self.course_number = row[COURSE_NUMBER].value
                break

        for row in self.row_values:
            if row[COURSE_NAME].value:
                self.course_name = row[COURSE_NAME].value
                break

        for row in self.row_values:
            if row[SIGHNED].value:
                self.course_signed_up = int(row[SIGHNED].value)
                break

        for row in self.row_values:
            if row[INSTRUCTOR_NAME].value:
                self.instructor.append(row[INSTRUCTOR_NAME].value)

    def _extract_dates(self):
        for row in self.row_values:
            if row[DATES].value:
                self.dates.append(row[DATES].value)

    def _clean_dates(self):
        pattern_time = r'([0-1][0-9]:[0-5][0-9])-([0-1][0-9]:[0-5][0-9])'
        pattern_class = 'درس'
        pattern_exam = 'امتحان\(([0-9]{4}).(\\d{2}).(\\d{2})\)'
        pattern_day = ': (\w+ شنبه|شنبه)'
        if self.dates:
            for date in self.dates:
                result_class = re.search(pattern_class, date)
                result_exam = re.search(pattern_exam, date)
                if result_exam:
                    print('EXAM:', self.start, result_exam.groups())
                    result_time = re.search(pattern_time, date)
                    if result_time:
                        print('\tEXAM-TIME:', self.start, result_time.groups())
                elif result_class:
                    result_day = re.search(pattern_day, date)
                    result_time = re.search(pattern_time, date)
                    print('CLASS:')
                    if result_day:
                        print('\tCLASS-DAY', self.start, result_day.groups())
                    if result_time:
                        print('\tCLASS-TIME', self.start, result_time.groups())
        else:
            self.dates = None

    def __str__(self):
        return "Table {}:{} - {}".format(self.start, self.end, len(self.row_values))


row_objs = []
start = end = 0
start_value = end_value = mid_value = ()
for row_index in range(1, worksheet.nrows):
    cell = worksheet.cell(row_index, 0)
    fmt = workbook.xf_list[cell.xf_index]
    bot_border = fmt.border.bottom_line_style
    top_border = fmt.border.top_line_style
    if bot_border == 1 and top_border == 1:
        row = Row(row_index, row_index, (worksheet.row(row_index),))
        row_objs.append(row)
    elif top_border == 1 and bot_border == 0:
        start = row_index
        start_value = worksheet.row(row_index)
    elif top_border == 0 and bot_border == 0:
        mid_value = worksheet.row(row_index)
    elif top_border == 0 and bot_border == 1:
        end = row_index
        end_value = worksheet.row(row_index)
        if mid_value:
            row = Row(start, end, (start_value, mid_value, end_value))
        else:
            row = Row(start, end, (start_value, end_value))
        row_objs.append(row)
        start = end = 0
        start_value = end_value = mid_value = 0


for row in row_objs:
    row.print_info()
