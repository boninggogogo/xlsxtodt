from typing import Any
from xlstodt.utils import get_cell_value, get_model_value


def write_types_sheet(types_sheet: Any, all_data: Any):
    types_col_indices = {'rack': {'name': 2,
                                  'firstSlotOffset': 3,
                                  'size_x': 4,
                                  'size_y': 5,
                                  'size_z': 6,
                                  'slot': 7},
                         'server': {'name': 10,
                                    'depth': 11,
                                    'slotOccupation': 12,
                                    'width': 13},
                         'acu': {'name': 16,
                                 'size_x': 17,
                                 'size_y': 18,
                                 'size_z': 19,
                                 'supplyFace_length': 20,
                                 'supplyFace_offset_x': 21,
                                 'supplyFace_offset_y': 22,
                                 'supplyFace_offset_z': 23,
                                 'supplyFace_side': 24,
                                 'supplyFace_width': 25,
                                 'returnFace_length': 26,
                                 'returnFace_offset_x': 27,
                                 'returnFace_offset_y': 28,
                                 'returnFace_offset_z': 29,
                                 'returnFace_side': 30,
                                 'returnFace_width': 31
                                 },
                         'box': {'name': 34,
                                 'top': 35,
                                 'bottom': 36,
                                 'rear': 37,
                                 'front': 38,
                                 'left': 39,
                                 'right': 40}
                         }

    row_len = len(types_sheet['J']) + 1
    # write rack_model
    for row in range(2, row_len):
        rack_name = get_model_value(sheet=types_sheet, row=row, column=types_col_indices['rack']['name'])
        if not rack_name:
            break
        rack_data = {
            "firstSlotOffset": get_cell_value(sheet=types_sheet, row=row,
                                              column=types_col_indices['rack']['firstSlotOffset']),
            "size": {
                "x": get_cell_value(sheet=types_sheet, row=row, column=types_col_indices['rack']['size_x']),
                "y": get_cell_value(sheet=types_sheet, row=row, column=types_col_indices['rack']['size_y']),
                "z": get_cell_value(sheet=types_sheet, row=row, column=types_col_indices['rack']['size_z'])
            },
            "slot": get_cell_value(sheet=types_sheet, row=row, column=types_col_indices['rack']['slot'])
        }

        all_data['geometryModel']['racks'][rack_name] = rack_data

    # write server_model
    for row in range(2, row_len):
        server_name = get_model_value(sheet=types_sheet, row=row, column=types_col_indices['server']['name'])
        if not server_name:
            break
        server_data = {
            "depth": get_cell_value(sheet=types_sheet, row=row, column=types_col_indices['server']['depth']),
            "slotOccupation": get_cell_value(sheet=types_sheet, row=row,
                                             column=types_col_indices['server']['slotOccupation']),
            "width": get_cell_value(sheet=types_sheet, row=row, column=types_col_indices['server']['width'])
        }

        all_data['geometryModel']['servers'][server_name] = server_data

    # write acu_model
    for row in range(2, row_len):
        acu_name = get_model_value(sheet=types_sheet, row=row, column=types_col_indices['acu']['name'])
        if not acu_name:
            break
        acu_data = {
                "size": {
                    "x": get_cell_value(sheet=types_sheet, row=row, column=types_col_indices['acu']['size_x']),
                    "y": get_cell_value(sheet=types_sheet, row=row, column=types_col_indices['acu']['size_y']),
                    "z": get_cell_value(sheet=types_sheet, row=row, column=types_col_indices['acu']['size_z'])
                },
                "supplyFace": {
                    "length": get_cell_value(sheet=types_sheet, row=row,
                                             column=types_col_indices['acu']['supplyFace_length']),
                    "offset": {
                        "x": get_cell_value(sheet=types_sheet, row=row,
                                            column=types_col_indices['acu']['supplyFace_offset_x']),
                        "y": get_cell_value(sheet=types_sheet, row=row,
                                            column=types_col_indices['acu']['supplyFace_offset_y']),
                        "z": get_cell_value(sheet=types_sheet, row=row,
                                            column=types_col_indices['acu']['supplyFace_offset_z'])
                    },
                    "side": get_model_value(sheet=types_sheet, row=row,
                                            column=types_col_indices['acu']['supplyFace_side']),
                    "width": get_cell_value(sheet=types_sheet, row=row,
                                            column=types_col_indices['acu']['supplyFace_width'])
                },
                "returnFace": {
                    "length": get_cell_value(sheet=types_sheet, row=row,
                                             column=types_col_indices['acu']['returnFace_length']),
                    "offset": {
                        "x": get_cell_value(sheet=types_sheet, row=row,
                                            column=types_col_indices['acu']['returnFace_offset_x']),
                        "y": get_cell_value(sheet=types_sheet, row=row,
                                            column=types_col_indices['acu']['returnFace_offset_y']),
                        "z": get_cell_value(sheet=types_sheet, row=row,
                                            column=types_col_indices['acu']['returnFace_offset_z'])
                    },
                    "side": get_model_value(sheet=types_sheet, row=row,
                                            column=types_col_indices['acu']['returnFace_side']),
                    "width": get_cell_value(sheet=types_sheet, row=row,
                                            column=types_col_indices['acu']['returnFace_width'])
                }
            }

        all_data['geometryModel']['acus'][acu_name] = acu_data

    # write box_model
    for row in range(2, row_len):
        box_name = get_model_value(sheet=types_sheet, row=row, column=types_col_indices['box']['name'])
        if not box_name:
            break
        box_data = {
            "faces": {
                "top": get_model_value(sheet=types_sheet, row=row, column=types_col_indices['box']['top']),
                "bottom": get_model_value(sheet=types_sheet, row=row, column=types_col_indices['box']['bottom']),
                "rear": get_model_value(sheet=types_sheet, row=row, column=types_col_indices['box']['rear']),
                "front": get_model_value(sheet=types_sheet, row=row, column=types_col_indices['box']['front']),
                "left": get_model_value(sheet=types_sheet, row=row, column=types_col_indices['box']['left']),
                "right": get_model_value(sheet=types_sheet, row=row, column=types_col_indices['box']['right'])
            }
        }

        all_data['geometryModel']['boxes'][box_name] = box_data
