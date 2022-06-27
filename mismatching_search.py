
import pandas as pd
import os


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
        sheet_names = xl.sheet_names
        field_name = file.split('Master_File_')

        field_mismatch = pd.DataFrame()

        for name in sheet_names:
            mismatch = []
            real_desc = []
            master_file_desc = []
            name_well = []
            well = xl.parse(name)
            try:
                print(name)
                well['year'] = well['Дата'].dt.year
                well = well.loc[well['year'] > 2021]

                # 18,19
                operations_desc = well.iloc[:, 18]
                codes_m = well.iloc[:, 19].astype(int)

                for j in range(0, len(codes_m)):
                    try:
                        k = codes_m[j]
                        key = get_keys_from_value(codes, k)
                        #  print(k)
                        description = type_of_work.get(key[0])
                        #   print(description)
                        #   print(operations_desc[j])
                        if description != operations_desc[j]:
                            real_desc.append(description)
                            master_file_desc.append(operations_desc[j])
                            name_well.append(name)
                            #       print(j,description,'\n','Master file',operations_desc[j])
                            mismatch.append(j + 2)
                    except (IndexError, KeyError):
                        continue
                print(f'По скважине {name} найдено {len(mismatch)} несоответствий кода и описания')

                mis = {'well_number': name_well, 'number of row': mismatch,
                       'Отображение в мастер файле': master_file_desc, 'Как должно быть': real_desc}
                mis = pd.DataFrame(data=mis)
                field_mismatch = pd.concat([field_mismatch, mis])


            except (ValueError, AttributeError,KeyError):
                print('Не удалось обработать скважину', name, '. Проверьте качество данных')
                continue

        field_mismatch.to_excel(f'Mismatch_errors_{field_name[1]}')


path = '.'
mismatch_search(path)