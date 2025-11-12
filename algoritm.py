from read_data import get_data_from_billing

def generate_schedule(file_name, days):
    info_about_classes, info_about_teachers = get_data_from_billing(file_name, days)
    schedule_classes = {name: {day: {i: None for i in range(1, 9)} for day in days} for name in info_about_classes}
    schedule_teachers = {name: {day: {i: None for i in range(1, 9)} for day in days} for name in info_about_teachers}
    return info_about_teachers

days = ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница']
print(generate_schedule("5E22B510.xlsx", days)["Курицына О.А."])