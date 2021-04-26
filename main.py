import pandas as pd
from check_list_dictionaries_lists_variables import templates_dir, templates, result_file_name, \
    result_file_sheet_name_list, directories_list, drop_columns_list, change_columns_list_dict, get_file_list


def get_info_from_dir(directory, template_ind, drop_list, change_dict):
    """
    Функция для сбора информации из чек-листов
    :param directory: директория, в которой размещены файлы (Госпрограммы, Планы реализации госпрограмм,
    Порядки предоставления субсидий)
    :param template_ind: индекс образца чек-листа (0 - Госпрограммы, 1 - Планы реализации госпрограмм,
    2 - Порядки предоставления субсидий)
    :param drop_list: список удаляемых столбцов (в случае, если форма чек-листа была изменена)
    :param change_dict: словарь переименования столбцов (в случае, если форма чек-листа была изменена)
    :return:
    """
    template = pd.read_excel(f'{templates_dir}{templates[template_ind]}', sheet_name='ЧЕК-ЛИСТ2', header=4)
    main_col = 'Наименование требования'
    template = template[[main_col]]
    print(f'Количество чек-листов в папке: {len(get_file_list(directory))}\n')
    result = template.copy()
    for file_ind, file in enumerate(get_file_list(directory)[:]):
        try:
            temp = pd.read_excel(f'{directory}{file}', sheet_name='ЧЕК-ЛИСТ2', header=4)
            temp.rename(columns={'Статус': file}, inplace=True)
            main_col = 'Наименование требования'
            temp = temp[[main_col, file]]

            temp = temp[~temp[main_col].isin(drop_list)]
            for key in change_dict:
                temp.loc[temp[temp[main_col].isin([key])].index, main_col] = change_dict[key]

            temp_col = temp[main_col].unique()
            result_col = result[main_col].unique()
            check = list(set(temp_col) - set(result_col))
            check_ = list(set(result_col) - set(temp_col))
            #             if (check!=[]) or (check_!=[]):
            #                 print(f'{file}\n')
            #                 if check!=[]:
            #                     print(f'Необходимо удалить/переименовать: {check}\n')
            #                 if check_!=[]:
            #                     print(f'Необходимо добавить: {check_}\n')
            #             else:
            #                 result = pd.merge(result, temp, how='outer', on=main_col)
            if (check != []):
                print(f'{file}\n')
                print(f'Необходимо удалить/переименовать: {check}\n')
            else:
                result = pd.merge(result, temp, how='outer', on=main_col)
        except PermissionError:
            print(f'\033[93m{file} редактируется другим пользователем\033[0m\n')
    result = result.T
    result.columns = result.iloc[0]
    result['file'] = result.index
    result.index = range(len(result.index))
    result.drop([0], inplace=True)
    result.index = range(len(result.index))
    if len(result) == len(get_file_list(directory)):
        print('Успешное завершение!')
        return result
    else:
        print('Загружены не все данные! Есть проблемы с названиями столбцов!')
        return result


if __name__ == '__main__':
    for template_ind in range(0, len(templates)):
        directory = directories_list[template_ind]
        drop_list = drop_columns_list[template_ind]
        change_dict = change_columns_list_dict[template_ind]
        result = get_info_from_dir(directory, template_ind, drop_list, change_dict)
        with pd.ExcelWriter(result_file_name, engine="openpyxl", mode='a') as writer:
            result.to_excel(writer, sheet_name=result_file_sheet_name_list[template_ind],
                            header=True, index=False, encoding='1251')
