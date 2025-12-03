from read_data import load_info

def generate_schedule(file_name, days):
    info_about_teachers, info_about_classes = load_info(file_name)
    schedule_classes = {name: {day: {i: None for i in range(1, 9)} for day in days} for name in info_about_classes}
    schedule_teachers = {name: {day: {i: None for i in range(1, 9)} for day in days} for name in info_about_teachers}
    print(schedule_classes)
    for class_name in info_about_classes.keys():
        for day in schedule_classes[class_name]:
            for number in schedule_classes[class_name][day]:
                if day != "Понедельник" and number != 1:
                    pass
    return info_about_teachers

days = ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница']
generate_schedule("aaa.xlsx", days)
