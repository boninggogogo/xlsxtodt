from typing import Any
from xlsxtodt.utils import get_cell_value, get_model_value, random_int, random_server_name


def read_racks(name: str, row_idx: int, all_data: dict, sheet: Any, col_indices: dict, types_dict: dict):
    model_name = get_model_value(sheet=sheet, row=row_idx, column=col_indices['Rack_Type'])
    if model_name not in types_dict['rack']:
        raise KeyError('The type of %s is not reflected in the xlsx file. ' % (model_name, ))
    rack = {
            "geometry": {
                "model": model_name,
                "location": {
                    "x": get_cell_value(sheet=sheet, row=row_idx, column=col_indices['location_x']),
                    "y": get_cell_value(sheet=sheet, row=row_idx, column=col_indices['location_y']),
                    "z": get_cell_value(sheet=sheet, row=row_idx, column=col_indices['location_z'])
                },
                "orientation": get_cell_value(sheet=sheet, row=row_idx, column=col_indices['location_orientation']),
                "hasBlankingPanel": False
            },
            "constructions": {
                "servers": {}
            }
        }
    all_data['constructions']['racks'][name] = rack


def write_server(server_sheet: Any, all_data: dict, col_indices: dict, server_random_number: int):

    for row_idx, row in enumerate(server_sheet.iter_rows(min_row=2, min_col=col_indices['server_rack'][0],
                                                         max_col=col_indices['server_rack'][1],
                                                         values_only=True), start=1):
        rack_col_value = row[0]
        if not rack_col_value:
            continue
        rack_model = all_data['constructions']['racks'][rack_col_value]['geometry']['model']
        max_range = int(all_data['geometryModel']['racks'][rack_model]['slot'])
        min_range = 1
        slots = random_int(number=server_random_number, number_range=[min_range, max_range])
        for slot in zip(slots, slots[1:]+[max_range]):
            max_slot = slot[1]-slot[0]
            server = {
                "geometry": {
                    "model": random_server_name(server_dict=all_data['geometryModel']['servers'],
                                                max_slot=max_slot if max_slot != 0 else 1),
                    "slotPosition": slot[0]
                }
            }
            server_name = str(rack_col_value)+'_Server_'+str(slot[0])
            all_data['constructions']['racks'][rack_col_value]['constructions']['servers'].update({server_name: server})

            # write_input
            server_input = {
                         "flowRate": round(row[4], 6),
                         "heatLoad": round(row[1]*100, 2),
                     }
            all_data['inputs']['servers'].update({server_name: server_input})

