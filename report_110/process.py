from xlrd import open_workbook

excel_obj = open_workbook('list_course.xls', formatting_info=True)
sheet = excel_obj.sheet_by_index(0)

class Row:
    def __init__(self, start, end):
        self.start = start
        self.end = end

    def __str__(self):
        return "Table {}:{}".format(self.start, self.end)

all_rows = []
col_index = 0
start = end = 0
for row_index in range(1, sheet.nrows):
    cell = sheet.cell(row_index, col_index)
    fmt = excel_obj.xf_list[cell.xf_index]
    bot_border = fmt.border.bottom_line_style
    top_border = fmt.border.top_line_style
    if bot_border == 1 and top_border == 1:
        row = Row(row_index, row_index)
        all_rows.append(row)
    elif top_border == 1 and bot_border == 0:
        start = row_index
    elif top_border == 0 and bot_border == 1:
        end = row_index
        row = Row(start, end)
        all_rows.append(row)
        start = end = 0


for row in all_rows:
    print(row)
