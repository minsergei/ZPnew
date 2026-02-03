import os
import xlrd


# открываем xls файл
def create_zp(path):
    workbook = xlrd.open_workbook(path)
    sheet = workbook.sheet_by_index(0)
    rows = [sheet.row_values(row, 0) for row in range(sheet.nrows)]

    # удаляем пустые записи
    for i in rows:
        while '' in i:
            i.remove('')
    # for i in rows:
    #     while ' ' in i:
    #         i.remove(' ')

    #определяем позиции начала расчетки
    list_nrows = []
    for i in range(len(rows)):
        if 'АО "НТЦ "Атлас"' in rows[i]:
            list_nrows.append(i)
    list_nrows.append(len(rows))

    # создаем расчетки для каждого сотрудника
    for i in range(len(list_nrows)-1):
        a = list_nrows[i]
        b = list_nrows[i+1]

        # print(rows[a:b])

        zp_list = rows[a:b]
        zp_name = zp_list[1][1][-7:] + '.txt'
        tab_number = zp_list[1][1][-6:]
        with open(os.path.join("calculations/", zp_name), 'w', encoding='cp1251') as file:
            for i in zp_list:
                file.write(str(i)+'\n')
