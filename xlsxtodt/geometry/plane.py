from typing import Any
from ..utils.read_cell import get_cell_value


def read_plane(name: str, row_idx: int, all_data: dict, sheet: Any, col_indices: dict, types_dict: dict):
    xyz = {
        "x": get_cell_value(sheet=sheet, row=row_idx, column=col_indices['location_x']),
        "y": get_cell_value(sheet=sheet, row=row_idx, column=col_indices['location_y']),
        "z": get_cell_value(sheet=sheet, row=row_idx, column=col_indices['location_z'])
    }
    all_data['geometry']['plane'].append(xyz)
