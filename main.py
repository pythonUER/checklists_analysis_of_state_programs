import os
from datetime import date
import pandas as pd

'''Получение перечня файлов'''
def GetFileList(directory):
    # Получаем список файлов в переменную files
    files = os.listdir(directory)
    # Получаем список файлов .xlsx
    reports = []
    for file in files:
        if file.endswith((".xlsx", ".xlsm")):
            reports.append(file)
    return reports

def GetInfoFromDir(directory):
    reports = GetFileList(directory)
    file_name = f'{directory}/{reports[0]}'
    df = pd.read_excel(file_name, sheet_name='ЧЕК-ЛИСТ2', header=4)

    for i in range(1, len(reports)):
        file_name = f'{directory}/{reports[i]}'
        temp = pd.read_excel(file_name, sheet_name='ЧЕК-ЛИСТ2', header=4)
        df = pd.concat([df, temp])

    df = df[['Наименование требования', 'Статус']]
    df.index = range(len(df))

    List_ = df['Наименование требования'].unique()
    data = pd.DataFrame()

    for i in List_:
        Temp1 = df[df['Наименование требования'] == i]
        Temp1.index = range(len(Temp1))
        col1 = Temp1[['Статус']]
        col1.columns = [i]
        data = pd.concat([data, col1], axis=1)

    return data

#////////////////////////////////////////////////////////////#
Folder = r'J:\~ 09_Машинное обучение_Прогноз показателей СЭР\ЧЕК-ЛИСТЫ и DATA-SHOP\ВЫГРУЗКА ЧЕК-ЛИСТОВ/'
File_name = 'RESULTS - Чек-листы.xlsx'

#Каталог из которого будем брать файлы
directory = r'J:\~Шаблоны\Чек-листы\Госпрограммы'
data = GetInfoFromDir(directory)
with pd.ExcelWriter(Folder + File_name, engine="openpyxl", mode='a') as writer:
    data.to_excel(writer, sheet_name='ГП_' + str(date.today()), header=True, index=False, encoding='UTF-8')

#Каталог из которого будем брать файлы
directory = r'J:\~Шаблоны\Чек-листы\План реализации'
data = GetInfoFromDir(directory)
with pd.ExcelWriter(Folder + File_name, engine="openpyxl", mode='a') as writer:
    data.to_excel(writer, sheet_name='ПланРеал_' + str(date.today()), header=True, index=False, encoding='UTF-8')

#Каталог из которого будем брать файлы
directory = r'J:\~Шаблоны\Чек-листы\Субсидии'
data = GetInfoFromDir(directory)
with pd.ExcelWriter(Folder + File_name, engine="openpyxl", mode='a') as writer:
    data.to_excel(writer, sheet_name='Субсидии_' + str(date.today()), header=True, index=False, encoding='UTF-8')
