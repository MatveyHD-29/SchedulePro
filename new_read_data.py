import openpyxl
from openpyxl import load_workbook

def load_sheet(file_name, sheet_name, method):
    wb = load_workbook(file_name)
    if sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
    else:
        return False
    mas = []
    s = sheet.iter_rows() if method == "rows" else sheet.iter_cols()
    for row in s:
        if row[0].value != None:
            vr = []
            for cell in row:
                cell_value = cell.value
                vr.append(str(cell_value))
            mas.append(vr)
        else:
            break
    return mas

def get_time_for_moving(file_name):
    mas = load_sheet(file_name, "Корпуса", "rows")
    time_for_moving = {}
    for name in mas[0]:
        if name != "-":
            for build in mas:
                if build[0] == name:
                    vr = {}
                    c = 0
                    for hour in build:
                        if hour != name and hour != "-":
                            vr[mas[0][c]] = int(hour)
                        c += 1
                    time_for_moving[name] = vr
    return time_for_moving

def load_teachers_and_classes():
    mas = load_sheet("aaa.xlsx", "База данных", "cols")
    teachers_list = []
    for row in mas[0][1:]:
        if row != "None":
            teachers_list.append(row)
    classes_list = []
    for row in mas[1][1:]:
        if row != "None":
            classes_list.append(row)
    return teachers_list, classes_list

def load_info_about_teachers():
    teacher_list, classes_list = load_teachers_and_classes()
    mas = load_sheet("aaa.xlsx", "Тарификация", "rows")
    info_about_teachers = {}
    for row in mas[1:]:
        c = 0
        for hour in row[2:-4]:
            if hour != "None":
                class_name = classes_list[c - 2]
                print(hour, class_name)
            c += 1


load_info_about_teachers()
