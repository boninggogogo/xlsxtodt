from typing import Any
from xlstodt.utils.read_cell import get_cell_value, get_model_value


def read_pillar(name: str, row_idx: int, all_data: dict, sheet: Any, col_indices: dict, types_dict: dict):
    model_name = get_model_value(sheet=sheet, row=row_idx, column=col_indices['Rack_Type'])
    if model_name not in types_dict['box']:
        raise KeyError('The type of %s is not reflected in the xlsx file. ' % (model_name,))
    pillar = {
                "geometry": {
                    "model": model_name,
                    "size": {
                        "x": get_cell_value(sheet=sheet, row=row_idx, column=col_indices['size_x']),
                        "y": get_cell_value(sheet=sheet, row=row_idx, column=col_indices['size_y']),
                        "z": get_cell_value(sheet=sheet, row=row_idx, column=col_indices['size_z'])
                    },
                    "location": {
                        "x": get_cell_value(sheet=sheet, row=row_idx, column=col_indices['location_x']),
                        "y": get_cell_value(sheet=sheet, row=row_idx, column=col_indices['location_y']),
                        "z": get_cell_value(sheet=sheet, row=row_idx, column=col_indices['location_z'])
                    }
                }
            }
    all_data['constructions']['boxes'][name] = pillar


def read_partition(name: str, row_idx: int, all_data: dict, sheet: Any, col_indices: dict, types_dict: dict):
    model_name = get_model_value(sheet=sheet, row=row_idx, column=col_indices['Rack_Type'])
    if model_name not in types_dict['box']:
        raise KeyError('The type of %s is not reflected in the xlsx file. ' % (model_name,))
    partition = {
                "geometry": {
                    "model": model_name,
                    "size": {
                        "x": get_cell_value(sheet=sheet, row=row_idx, column=col_indices['size_x']),
                        "y": get_cell_value(sheet=sheet, row=row_idx, column=col_indices['size_y']),
                        "z": get_cell_value(sheet=sheet, row=row_idx, column=col_indices['size_z'])
                    },
                    "location": {
                        "x": get_cell_value(sheet=sheet, row=row_idx, column=col_indices['location_x']),
                        "y": get_cell_value(sheet=sheet, row=row_idx, column=col_indices['location_y']),
                        "z": get_cell_value(sheet=sheet, row=row_idx, column=col_indices['location_z'])
                    }
                }
            }
    all_data['constructions']['boxes'][name] = partition
