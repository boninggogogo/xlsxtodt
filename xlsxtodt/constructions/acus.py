from typing import Any
from xlsxtodt.utils.read_cell import get_cell_value, get_model_value


def read_acus(name: str, row_idx: int, all_data: dict, sheet: Any, col_indices: dict, types_dict: dict):
    model_name = get_model_value(sheet=sheet, row=row_idx, column=col_indices['Rack_Type'])
    if model_name not in types_dict['acu']:
        raise KeyError('The type of %s is not reflected in the xlsx file. ' % (model_name,))
    acu = {
            "geometry": {
                "model": model_name,
                "orientation": get_cell_value(sheet=sheet, row=row_idx, column=col_indices['location_orientation']),
                "location": {
                    "x": get_cell_value(sheet=sheet, row=row_idx, column=col_indices['location_x']),
                    "y": get_cell_value(sheet=sheet, row=row_idx, column=col_indices['location_y']),
                    "z": get_cell_value(sheet=sheet, row=row_idx, column=col_indices['location_z'])
                }
            }
        }
    all_data['constructions']['acus'][name] = acu
