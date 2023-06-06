from xlstodt import XlsToDT

Leve4_R01_room = XlsToDT()
Leve4_R01_room.xls_to_dt(path_xls=r'data/input/Excel_to_DT_Level_4_R01.xlsx',
                         geometry_info=['Geometry_info', ],
                         crca_info=['CRAC_Info', ],
                         server_info=['Server_info', ],
                         servers_count=[10, ],
                         model_info='Model_info',
                         path_output='data/output/',
                         name_in_dt=['Excel_to_DT_Level_4_R01', ],
                         output_file=['Excel_to_DT_Level_4_R01.json', ]
                         )

