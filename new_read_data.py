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
                if cell_value != None:
                    vr.append(str(cell_value))
            mas.append(vr)
        else:
            break
    return mas

def load_teachers_classes():
    mas = load_sheet("aaa.xlsx", "База данных", "cols")
    teachers_list = mas[0][1:]
    classes_list = mas[1][1:]
    print(classes_list)

def load_info_about_teachers():
    pass

print(load_teachers_classes())