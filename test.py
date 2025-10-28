import openpyxl
from openpyxl import load_workbook


def load_sheet(file_name, sheet_name):
    wb = load_workbook(file_name)
    sheet = wb.active
    if sheet_name not in wb.sheetnames:
        return False
    sheet = wb[sheet_name]
    vr = []
    mas = []
    for col in sheet.iter_rows():
        if col[0] != None:
            for cell in col:
                cell_value = cell.value
                if cell_value != None:
                    vr.append(cell_value)
            if vr != []:
                mas.append(vr)
            vr = []
        else:
            break
    return mas

def load_building(file_name):
    sl_building = {}
    mas = load_sheet(file_name, "Корпуса")
    for i in mas[0]:
        vr = {}
        if i != "-":
            for build in mas:
                if build[0] == i:
                    c = 0
                    for j in build:
                        if j != "-" and j != i:
                            vr[mas[0][c]] = j
                        c += 1
            sl_building[i] = vr
    return sl_building

def load_data(file_name):
    mas = load_sheet(file_name, "Лист1")
    teacher_workload = {}
    sl_index_class = {}
    c = 2
    for i in mas[0][2:-2]:
        sl_index_class[c] = i
        c += 1
    for row in mas[1:]:
        if row[0] != "ИТОГО":
            c = 2
            time_mas = {}
            for hour in row[2:-2]:
                if hour != "-":
                    time_mas[sl_index_class[c]] = hour
                c += 1
            vr = {row[1]: {"Классы": time_mas}}
            teacher_workload[row[0]] = {"Предмет": vr, "Кабинет": 19} # Сделать функцию для получения кабинета, корпуса и других параметров учителя
        else:
            break
    print(teacher_workload["Головнина Т.А."])


print(load_data("bbb.xlsx"))