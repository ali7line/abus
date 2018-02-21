from xlrd import open_workbook

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
TIME = 14


class Row:
    def __init__(self, start, end, row_values):
        self.start = start
        self.end = end
        self.row_values = row_values
        self.get_year()

    def get_year(self):
        # get year
        for row in self.row_values:
            if row[YEAR]:
                self.year = row[YEAR].value
                break

    def print_year(self):
        print(self.year)

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


for row in row_objs[:3]:
    print(row.year)
