from xlsxtodt import xtd


R02_room1 = xtd()
R02_room1.xlsx_to_dt(path_xlsx=r'data/input/Data Hall Geometry input_R02CA.xlsx',
                     sheets=['level 2', 'level 3', 'Level 4'],
                     server_sheets=['Level 2 Server', 'Level 3 Server', 'Level 4 Server'],
                     server_random_number=[10, 10, 10],
                     types='geometryModel',
                     output_filename=['Level_2_R02.json', 'Level_3_R02.json', 'Level_4_R02.json'],
                     name_in_dt=['Level_2_R02_v1', 'Level_3_R02_v1', 'Level_4_R02_v1']
                     )
