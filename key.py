import openpyxl
import csv
import random
import string

fn = "test.xlsx"
keyword_list = []
blacklist_list = []

def check_xlsx_format(filename):
    """
    Check file format and extension
    """
    file = filename
    try:
        openpyxl.load_workbook(file)
        return True
    except Exception as e:
        return e


def upload_file(filename):
    """
    Read file values and update keywords list by this values
    """
    file = filename
    counting = 0
    if check_xlsx_format(file) == True:
        wb = openpyxl.load_workbook(file)
        for sheet in wb:
            for line in sheet.values:
                if all(line):
                    keyword_list.append(line)
                    counting += 1
        return f'В спсиок добавленно {counting} фраз'
    else:
        return f'Файл должен быть xlsx'


def id_generator(size=5, chars=string.ascii_lowercase + string.digits):
    """
    Generate randomly 5 chars
    """
    return ''.join(random.choice(chars) for _ in range(size))


def generate_csv_file(keyword_list):
    """
    Generate CSV file with keywords devided by groups
    """
    if len(keyword_list) != 0:
        fieldnames = ["Группа", "Ключи", "Частотность"]
        group_name = str(input("Название группы: "))
        word = str(input("Ключ: "))
        counting = 0
        with open(f'tree_.csv', 'a', encoding='cp1251') as f: # tree_{id_generator()}.csv
            write = csv.DictWriter(f, fieldnames=fieldnames, delimiter=';',
                                   extrasaction='ignore', dialect='excel')
            write.writerow({"Группа": group_name})
            for keyword in keyword_list:
                if word in str(keyword):
                    counting += 1
                    blacklist_list.append(keyword)
                    keyword_list.remove(keyword)
                    try:
                        new_row = {"Группа": "", "Ключи": keyword[0],
                                   "Частотность": keyword[1]}
                        write.writerow(new_row)
                    except IndexError:
                        new_row = {"Группа": "", "Ключи": keyword[0],
                                   "Частотность": "-"}
                        write.writerow(new_row)
        return f'Файл создан'
    else:
        return f'Пустой список'

"""
- Рефактирон generate_csv_file()
- Считать кол-во созданных групп
- Подгруппа у группы ?
"""

print(upload_file(fn))
while True:
    print(generate_csv_file(keyword_list))
print(generate_csv_file(keyword_list))


#keyword_list = [(keyword[0].lower(), keyword[1]) for keyword in keyword_list] # lowercase
#keyword_list.sort() # сортировка по алфовиту
#keyword_list.sort(key=lambda x: (x[1], x[0]), reverse=True) # сортировка по частотности
print(keyword_list)


# tuple_new = [('Some shiT', 321)]
# print(tuple_new)
# for l in tuple_new:
#     print(l[0].lower())


# filename = ("test.csv")
# with open(filename, "r", encoding='cp1251') as f:
#     read = csv.reader(f, delimiter=';')
#     for row in read:
#         print(row)

