import os
import numpy as np
import pandas as pd
import warnings
warnings.filterwarnings("ignore")
path = '.'
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)




def error_search(path):
    prefixed = [filename for filename in os.listdir(path) if filename.startswith("Master_File")]

    # поиск БМР операций с проходкой

    incorrect_penetration = pd.DataFrame()

    for file in prefixed:
        print('Проверка на проходку')
        xl = pd.ExcelFile(file)
        sheet_names = xl.sheet_names[:-1]
        field_name = file.split('Master_File_')
        incorrect_penetration = pd.DataFrame()
        for name in sheet_names:
            well = xl.parse(name)
            print(name)
            try:
                well['year'] = well['Дата'].dt.year
                well = well.loc[well['year'] > 2021]
                #  errors = well.loc[ well['Дата']        ]
                errors = well.loc[(well['Проходка'] > 0) & (well['Вид работ'] > 300) & (well['Вид работ'] < 2000)]
                # old codes
                incorrect_penetration = pd.concat([incorrect_penetration, errors])
            except (AttributeError,KeyError):
                print(f'Не удалось обработать скважину {name}')
                continue
        incorrect_penetration.to_excel(f'Incorrect_penetration_of_{field_name[1]}')



    # поиск некорректного СПО

    for file in prefixed:
        print('Проверка на подъем/спуск')
        xl = pd.ExcelFile(file)
        sheet_names = xl.sheet_names[:-1]
        field_name = file.split('Master_File_')
        tripping_in_incorrect = pd.DataFrame()
        tripping_out_incorrect = pd.DataFrame()
        for name in sheet_names:
            well = xl.parse(name)
            print(name)
            try:
                try:
                    well['year'] = well['Дата'].dt.year
                    well = well.loc[well['year'] > 2021]
                except AttributeError:
                    print(f'Ошибка в дате на кусту {name}')
                # спуск
                tripping_out = well[well['Код операции'].str.lower().notnull()][
                    well['Код операции'].dropna().str.lower().str.contains('спуск')]
                try:
                    tripping_out_inc = tripping_out[tripping_out['Операции'].str.lower().str.contains('подъем')]
                    tripping_out_incorrect = pd.concat([tripping_out_incorrect, tripping_out_inc])
                except (ValueError, KeyError):
               #     print(f'Не удалось обработать скважину {name}')
                    continue
                #  tripping_out_inc.to_excel(f'incorrect_tripping_out_{field_name[1]}')

                # подъем
                tripping_in = well[well['Код операции'].str.lower().notnull()][
                    well['Код операции'].dropna().str.lower().str.contains('подъем')]
                try:
                    tripping_in_inc = tripping_in[tripping_in['Операции'].str.lower().str.contains('спуск')]
                    tripping_in_incorrect = pd.concat([tripping_in_incorrect, tripping_in_inc])
                except (ValueError, KeyError):
                    continue
               #     print(f'Не удалось обработать скважину {name}')
            except (ValueError, KeyError):
                print('AtrributeError - проверьте качество данных')
                continue

        tripping_in_incorrect.to_excel(f'incorrect_tripping_in_{field_name[1]}')
        tripping_out_incorrect.to_excel(f'incorrect_tripping_out_{field_name[1]}')


def get_keys_from_value(d, val):
    return [k for k, v in d.items() if v == val]


def mismatch_search(path):
    prefixed = [filename for filename in os.listdir(path) if filename.startswith("Master_File")]

    # поиск несоответствий кода и операции
    a = pd.read_excel('Кодировка.xlsx', sheet_name='Новая кодировка').to_dict()
    type_of_work = a['Вид работ']
    codes = a['Код']

    for file in prefixed:
        xl = pd.ExcelFile(file)
        sheet_names = xl.sheet_names[:-1]
        field_name = file.split('Master_File_')

        field_mismatch = pd.DataFrame()

        for name in sheet_names:
            mismatch = []
            real_desc = []
            master_file_desc = []
            name_well = []
            date = []
            well = xl.parse(name)
            try:
                print(name)
                well['year'] = well['Дата'].dt.year
                well = well.loc[well['year'] > 2021]

                # 18,19
                operations_desc = well.iloc[:, 18]
                codes_m = well.iloc[:, 19].astype(int)
                date = well.iloc[:, 10]

                for j in range(0, len(codes_m)):
                    try:
                        k = codes_m[j]
                        key = get_keys_from_value(codes, k)
                        #  print(k)
                        description = type_of_work.get(key[0])

                        if description != operations_desc[j]:
                            real_desc.append(description)
                            master_file_desc.append(operations_desc[j])
                            name_well.append(name)
                            date.append(date[j])
                            #       print(j,description,'\n','Master file',operations_desc[j])
                            mismatch.append(j + 2)
                    except:
                        continue
                print(f'По скважине {name} найдено {len(mismatch)} несоответствий кода и описания')

                mis = {
                          'well_number': name_well, 'date':date, 'number of row': mismatch, 'Отображение в мастер файле': master_file_desc, 'Код из справочника': real_desc, 'Код в мастер файле': codes_m}
                mis = pd.DataFrame(data=mis)
                field_mismatch = pd.concat([field_mismatch, mis])


            except:
                print('Не удалось обработать скважину', name, '. Проверьте качество данных')
                continue

        field_mismatch.to_excel(f'Mismatch_errors_{field_name[1]}')

if __name__ == '__main__':
  #  error_search(path)
    mismatch_search(path)