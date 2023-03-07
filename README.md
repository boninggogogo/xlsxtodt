# Currently, the generated file does not have "acus" in the input section.

``` python
xlsx_to_dt(  
           path_xlsx = r'data/input/Data Hall Geometry input_R02CA.xlsx',  
           sheets = ['level 2', 'level 3', 'Level 4'],  
           server_sheets = ['Level 2 Server', 'Level 3 Server', 'Level 4 Server'],  
           server_random_number = [10, 10, 10],  
           types ='geometryModel',  
           output_filename = ['Level_2_R02.json', 'Level_3_R02.json', 'Level_4_R02.json'],  
           name_in_dt = ['Level_2_R02_v1', 'Level_3_R02_v1', 'Level_4_R02_v1']  
           )
```

__The order of values in the lists passed for sheets, server_sheets, server_random_number, output_filename, and name_in_dt must be consistent.__
* __path_xlsx:__ The file address of the Excel file that needs to be passed in.
* __sheets:__    The name of the worksheet corresponding to each room in the Excel file.
* __server_sheets:__ The name of the worksheet corresponding to the server in each room in the Excel file.
* __server_random_number:__ The number of randomly generated servers is needed for the racks in the room.
* __types:__ For geometryModel is needed a separate worksheet, here you need to pass in the name of the worksheet.
* __output_filename:__ Output the file name and address of the DT file.
* __name_in_dt:__ The value of name in the DT file can be defined here.
