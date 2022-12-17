import xlwings as xw
from datetime import datetime

a = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь']
list = {
    'Расход ЕМГ':'A:D',
    'Расход ЭМ':'E:H',
    'Перемещение в ОП':'J:M',
    'Перемещение в ОСС':'O:R',
    'Перемещение в ОСР':'T:W',
    'Перемещение в SMT':'Y:AB'
}
dt = datetime.now()
m = dt.strftime("%m")
dt = dt.strftime("%d.%m.%Y")
file = f'Расход {a[int(m)-1]}.xlsx'
# dt = '22.11.2022'
# file = f'Расход {a[10]}.xlsx'
try:
    wb = xw.Book(file)
    xl = wb.sheets[dt]

    with open("Для Телеграмма.txt", "w") as out:
        out.writelines(f'{dt}\n')
        for keys, value in list.items():
            items = []
            home = f'{value[0]}3'
            end = f'{value[2:4]}999'
            data_pd = xl.range(f'{home}:{end}')
            if data_pd.value[0][0] != None:
                out.writelines('\n')
                out.writelines(f'{keys}\n')
                out.writelines(data_pd.value[0][0])
                out.writelines('\n')
            for i in range(995):
                if data_pd.value[i][0] == None:
                    break
                if data_pd.value[i][3] != None:
                    out.writelines(f'{data_pd.value[i][3]}\n')
except BaseException:
    print("Сохрани файл")
else:
    print("Готово")
finally:
    input()






