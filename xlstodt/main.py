import json
import re
from copy import deepcopy

from .geometryModel import write_types_sheet
from .geometry import read_room_height, read_plane
from .constructions import read_raised_floor, read_raised_floor_height, read_pillar, read_acus, \
    read_racks, read_partition, write_server
import openpyxl
from typing import Union, Any

from .inputs import write_crac


class XlsToDT:
    # 将各种类型的值在xlsx文件中的列的下标初始化
    # Initialize the index of columns with values of various types in the xlsx file.
    col_indices = {'name': 1, 'size_x': 2, 'size_y': 3, 'size_z': 4, 'location_orientation': 5, 'location_x': 6,
                   'location_y': 7, 'location_z': 8, 'height': 9, 'Rack_Type': 10, 'server_rack': [2, 6],
                   'CRAC_input': [1, 4]}
    funcs = {1: read_room_height, 2: read_plane, 3: read_raised_floor_height, 4: read_raised_floor, 5: read_pillar,
             6: read_acus, 7: read_racks, 8: read_partition}

    def __init__(self):

        # 使用dict类型来存储对应DT文件格式的数据
        # Using the dict type to store data in the corresponding format of the DT file.
        self.all_data = None
        self.all_data_ = {
            "name": '',
            "geometryModel": {
                "acus": {},
                "racks": {},
                "servers": {},
                "boxes": {}

            },
            "geometry": {
                "height": None,
                "plane": []
            },
            "constructions": {
                "falseCeiling": None,
                "raisedFloor": {
                    "geometry": {
                        "height": None,
                        "openings": {}
                    }
                },
                "boxes": {},
                "acus": {},
                "racks": {},
                "sensors": {}
            },
            "inputs": {
                'acus': {},
                'servers': {}
            }
        }

        self.sheet = None
        self.types_sheet = None
        self.types_sheet_name = None
        self.servers_count = None

        self.types = None

    # 将读取xlsx文件的所有数据，按照DT文件的格式要求，以字典的类型存储在“self.sheet”中
    # Read all data from the xlsx file and store it in a dictionary type in "self.sheet" according to
    # the format requirements of the DT file.
    def xls_to_dt(self, path_xls: str, geometry_info: Union[list], model_info: Union[str],
                  server_info: Union[list], servers_count: Union[list], crca_info: Union[list],
                  output_file: Union[list], name_in_dt=None, path_output=''):
        self.types_sheet_name = model_info

        for name in zip(geometry_info, server_info, servers_count, output_file, name_in_dt, crca_info):
            self.all_data = deepcopy(self.all_data_)
            self.read_xlsx(path_xls=path_xls, sheet_name=name[0], server_sheet_name=name[1],
                           servers_count=name[2], crca_info=name[5])
            self.create_dt(filename=name[3], name=name[4], path_output=path_output)

    def read_xlsx(self, path_xls: str, sheet_name: Union[str, int], server_sheet_name: Union[str, int],
                  servers_count: int, crca_info: Union[str, int]):

        workbook = openpyxl.load_workbook(filename=path_xls, data_only=True, )

        self.sheet = workbook[sheet_name] if isinstance(sheet_name, str) else workbook.worksheets[sheet_name]

        # 先添加types
        self.types_sheet = workbook[self.types_sheet_name]
        write_types_sheet(types_sheet=self.types_sheet, all_data=self.all_data)
        self.types = {'rack': self.all_data['geometryModel']['racks'].keys(),
                      'server': self.all_data['geometryModel']['servers'],
                      'acu': self.all_data['geometryModel']['acus'].keys(),
                      'box': self.all_data['geometryModel']['boxes'].keys()
                      }

        for row_idx, row in enumerate(self.sheet.iter_rows(min_row=3, min_col=1, max_col=1, values_only=True), start=3):
            name = row[0]
            if not name:
                break
            name_type = self.get_input_type(name)
            if name_type != 0:
                self.dispatch_function('read mode', name_type=name_type, name=name, row_idx=row_idx)
            else:
                raise KeyError("The format of %s may not be correct." % (name,))

        write_server(server_sheet=workbook[server_sheet_name], all_data=self.all_data,
                     col_indices=XlsToDT.col_indices, servers_count=servers_count)
        write_crac(crca_sheet=workbook[crca_info], all_data=self.all_data,
                   col_indices=XlsToDT.col_indices)

    def create_dt(self, filename: Any, name='', path_output=''):
        self.all_data['name'] = name
        with open(path_output + filename, "w") as outfile:
            json.dump(self.all_data, outfile, ensure_ascii=False, indent=4)

    # 该函数将返回所读取到的值所对应的DT文件中的类型序号，“dispatch_function()”可以通过序号和funcs[]使用对应的函数方法去读取数据
    # This function will return the type number in the DT file corresponding to the value read,
    # and "dispatch_function()" can use the corresponding function method to read the data by the number and funcs[].
    @staticmethod
    def get_input_type(name):
        # 定义关键字列表
        # Define a list of keywords.
        keyword_lists = [
            ["Room_Height", "Room Height"],
            ["Room_point.*"],
            ["Raised floor height", "Raised_floor_height"],
            ["Perforated_tile.*", "RF_Opening.*", "Hole.*"],
            ["Pillar.*", "Containment.*"],
            ["CRAH.*", "CRAC.*", "ACCPU.*", "CHWCAHU.*"],
            ["Rack.*"],
            ["Partition.*"]
        ]

        # print(type(name), name)
        for i, keywords in enumerate(keyword_lists):
            regex = "^" + "|".join(keywords) + "$"
            match = re.search(regex, name, re.IGNORECASE)
            if match:
                return i + 1
        # 如果没有匹配到任何一种类型，则返回0
        # If no match is found for any type, then return 0.
        return 0

    def dispatch_function(self, mode, name_type, name, row_idx):
        if mode == 'read mode':
            XlsToDT.funcs[name_type](name=name,
                                      row_idx=row_idx,
                                      all_data=self.all_data,
                                      sheet=self.sheet,
                                      col_indices=XlsToDT.col_indices,
                                      types_dict=self.types
                                      )
        elif mode == 'write mode':
            ...
