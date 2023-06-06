from typing import Any
from ..utils.read_cell import get_cell_value


def read_height(name: str, row_idx: int, all_data: dict, sheet: Any, col_indices: dict, types_dict: dict):
    all_data['geometry']['height'] = get_cell_value(sheet=sheet, row=row_idx, column=col_indices['height'])
