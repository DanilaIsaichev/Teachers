from openpyxl import load_workbook
import os

# Словарь для номеров столбцов с данными
teacher_numbers = {
    'Фамилия': None,
    'Имя': None,
    'Отчество': None,
    'Адрес электронной почты': None,
    'Должность': None,
    'Дополнительная должность': None,
    'Группа сотрудников': None,
    'Образование': None,
    'Квалификационная категория по основной должности': None,
    'Общий стаж': None,
    'Педагогический стаж': None,
    'Образовательное учреждение': None,
}

courses_numbers = {
    'ОООД повышения квалификации (полное наименование)': None,
    'Название курса': None,
    'Объем курса (часы)': None,
    'Дата выдачи': None
}

# Словарь для вывода информации о сотрудниках
teacher = {
    'Фамилия': None,
    'Имя': None,
    'Отчество': None,
    'Адрес электронной почты': None,
    'Должность': None,
    'Дополнительная должность': None,
    'Группа сотрудников': None,
    'Образование': None,
    'Квалификационная категория по основной должности': None,
    'Общий стаж': None,
    'Педагогический стаж': None,
    'Образовательное учреждение': None,
}

courses = {
    'ОООД повышения квалификации (полное наименование)': [],
    'Название курса': [],
    'Объем курса (часы)': [],
    'Дата выдачи': []
}

try:
    # Загрузка файла teachers.xlsx
    teacher_wb = load_workbook('teachers.xlsx')
    # Получение листа через его имя
    teacher_sheet_name = teacher_wb.sheetnames[0]
    teacher_sheet = teacher_wb[teacher_sheet_name]
except FileNotFoundError as e:
    print('Ошибка:', e)
    input()
    exit()

# Получение номеров столбцов с необходимыми данными о сотрудиках
col_numb = 1  # Номер столбца
for i in teacher_numbers:  # Шаг по столбцам
    while teacher_sheet.cell(1, col_numb).value is not None:
        if teacher_sheet.cell(1, col_numb).value == i:
            teacher_numbers[i] = col_numb
            break
        col_numb += 1  # Шаг по строкам

try:
    # Загрузка файла courses.xlsx
    courses_wb = load_workbook('courses.xlsx')
    # Получение листа через его имя
    courses_sheet_name = courses_wb.sheetnames[0]
    courses_sheet = courses_wb[courses_sheet_name]
except FileNotFoundError as e:
    print('Ошибка:', e)
    input()
    exit()

# Получение номеров столбцов с необходимыми данными о сотрудиках
col_numb = 1  # Номер столбца
for i in courses_numbers:  # Шаг по столбцам
    while courses_sheet.cell(1, col_numb).value is not None:
        if courses_sheet.cell(1, col_numb).value == i:
            courses_numbers[i] = col_numb
            break
        col_numb += 1  # Шаг по строкам

print('Начинаю создание файлов...\n')

# Создание папки teachers
try:
    os.mkdir('teachers')
except FileExistsError:
    print('Папка teachers уже существует\n')

teachers_row_numb = 2  # Номер строки в таблице teachers
courses_row_numb = 2  # Номер строки в таблице courses
number = 1  # Номер учителя без адреса электронной почты
while teachers_row_numb <= teacher_sheet.max_row:
    for i in teacher_numbers:
        # Получение информации о стаже
        if i in ['Общий стаж', 'Педагогический стаж']:
            experience = teacher_sheet.cell(teachers_row_numb, teacher_numbers[i]).value
            for k in experience:

                # Стаж от пяти лет
                if k == 'л':
                    teacher[i] = experience[:experience.find("л")] + ' лет'
                    break

                # Стаж от года до пяти лет
                elif k == 'г':
                    if int(experience[0]) > 1:
                        teacher[i] = experience[:experience.find("г")] + ' года'  # Стаж больше года, но меньше пяти лет
                    else:
                        teacher[i] = '1 год'  # Стаж в 1 год
                    break

                # Стаж меньше года
                elif k in ['м', 'д']:
                    teacher[i] = 'Меньше года'
                    break
        else:  # Получение остальной информации
            teacher[i] = teacher_sheet.cell(teachers_row_numb, teacher_numbers[i]).value

    # Получение информации о курсах повышения квалификации
    while courses_row_numb <= courses_sheet.max_row and (courses_sheet.cell(courses_row_numb, 2).value == teacher['Фамилия'] or courses_sheet.cell(courses_row_numb, 2).value is None):
        courses['ОООД повышения квалификации (полное наименование)'].append(courses_sheet.cell(courses_row_numb, courses_numbers['ОООД повышения квалификации (полное наименование)']).value)
        courses['Название курса'].append(courses_sheet.cell(courses_row_numb, courses_numbers['Название курса']).value)
        courses['Объем курса (часы)'].append(courses_sheet.cell(courses_row_numb, courses_numbers['Объем курса (часы)']).value)
        courses['Дата выдачи'].append(courses_sheet.cell(courses_row_numb, courses_numbers['Дата выдачи']).value.strftime("%Y"))  # Берётся только год
        courses_row_numb += 1  # Шаг по строкам

    # Создание файла
    if teacher['Адрес электронной почты'] != '' and teacher['Адрес электронной почты'] is not None:
        file_name = teacher['Адрес электронной почты'][:teacher['Адрес электронной почты'].find('@')]
    else:
        file_name = 'teacher ' + str(number)
        number += 1
    print(teacher['Фамилия'], end="")

    # Создание папки
    try:
        os.mkdir('teachers/' + file_name)
    except FileExistsError:
        print(' - папка уже существует, редактирую файл', end="")

    # Открытие файла
    file = open('teachers/' + file_name + '/' + file_name + '.md', 'w')

    # Вывод основных данных об учителе
    file.write(
        'title: \'' + str(teacher['Фамилия']) + ' ' + str(teacher['Имя']) + ' ' + str(teacher['Отчество']) + '\'\n')
    file.write('place_education: \'' + str(teacher['Образовательное учреждение']) + '\'\n')
    file.write('general_experience: \'' + str(teacher['Общий стаж']) + '\'\n')
    file.write('position: \'' + str(teacher['Должность']) + '\'\n')
    file.write('education: \'' + str(teacher['Образование']) + '\'\n')
    file.write('category: \'' + str(teacher['Квалификационная категория по основной должности']) + '\'\n')
    file.write('experience: \'' + str(teacher['Педагогический стаж']) + '\'\n')
    file.write('email: \'' + str(teacher['Адрес электронной почты']) + '\'\n')
    file.write('course: \n')

    # Вывод данных о курсах повышения квалификации
    while courses['ОООД повышения квалификации (полное наименование)']:
        file.write('\t-\n')
        file.write('\t\tplace: \'' + str(courses['ОООД повышения квалификации (полное наименование)'][0]) + '\'\n')
        file.write('\t\ttitle: \'' + str(courses['Название курса'][0]) + '\'\n')
        file.write('\t\thour: \'' + str(courses['Объем курса (часы)'][0]) + '\'\n')
        file.write('\t\tdate: \'' + str(courses['Дата выдачи'][0]) + '\'\n')
        del courses['ОООД повышения квалификации (полное наименование)'][0]
        del courses['Название курса'][0]
        del courses['Объем курса (часы)'][0]
        del courses['Дата выдачи'][0]
    file.write('creator: admin\n')

    # Закрытие файла
    file.close()
    print(' - готово!\n')

    # Очистка словаря
    for i in teacher:
        teacher[i] = None

    teachers_row_numb += 1

print('Всё готово!\n\nНажмите Enter, чтобы закончить...')
input()
