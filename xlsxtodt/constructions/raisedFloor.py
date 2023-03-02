from typing import Any
from xlsxtodt.utils.read_cell import get_cell_value


def read_raised_floor_height(name: str, row_idx: int, all_data: dict, sheet: Any, col_indices: dict, types_dict: dict):
    all_data['constructions']['raisedFloor']['geometry']['height'] = get_cell_value(sheet=sheet,
                                                                                    row=row_idx,
                                                                                    column=col_indices['height'])


def read_raised_floor(name: str, row_idx: int, all_data: dict, sheet: Any, col_indices: dict, types_dict: dict):
    opening = {
        "location": {
            "x": get_cell_value(sheet=sheet, row=row_idx, column=col_indices['location_x']),
            "y": get_cell_value(sheet=sheet, row=row_idx, column=col_indices['location_y']),
            "z": get_cell_value(sheet=sheet, row=row_idx, column=col_indices['location_z'])
        },
        "size": {
            "x": get_cell_value(sheet=sheet, row=row_idx, column=col_indices['size_x']),
            "y": get_cell_value(sheet=sheet, row=row_idx, column=col_indices['size_y']),
            "z": get_cell_value(sheet=sheet, row=row_idx, column=col_indices['size_z'])
        }
    }
    all_data['constructions']['raisedFloor']['geometry']['openings'][name] = opening
