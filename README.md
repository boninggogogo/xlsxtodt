# Currently, the generated file does not have "acus" in the input section.

``` python
Leve4_R01_room = XlsToDT()
Leve4_R01_room.xls_to_dt(path_xls=r'data/Excel_to_DT_Level_4_R01.xlsx',
                         geometry_info=['Geometry_info', ],
                         crca_info=['CRAC_Info', ],
                         server_info=['Server_info', ],
                         servers_count=[10, ],
                         model_info='Model_info',
                         path_output='./output/',
                         name_in_dt=['Excel_to_DT_Level_4_R01', ],
                         output_file=['Excel_to_DT_Level_4_R01.json', ]
                         )
```

* __path_xls:__ The file address of the Excel file that needs to be passed in.
* __geometry_info:__ Relevant information for geometry, corresponds to the worksheet name in the Excel file. This parameter specifies the name of the worksheet that contains geometry-related data.
* __crca_info:__ Relevant information for CRCA, corresponds to the worksheet name in the Excel file. This parameter specifies the name of the worksheet that contains CRCA-related data.
* __server_info:__ Relevant information for servers, corresponds to the worksheet name in the Excel file. This parameter specifies the name of the worksheet that contains server-related data.
* __servers_count:__ The number of servers to be randomly generated. This parameter specifies the number of servers to generate for each room.
* __model_info:__ Relevant information for model_info, corresponds to the worksheet name in the Excel file. This parameter specifies the name of the worksheet that contains model_info-related data.
* __path_output:__ The address where the generated file will be saved. This parameter specifies the output path for the generated DT file.
* __name_in_dt:__ The value of the name variable in the DT file can be defined here. This parameter sets the value of the name variable in the generated DT file.
* __output_file:__ The name of the generated file. This parameter specifies the file name and extension for the generated DT file.