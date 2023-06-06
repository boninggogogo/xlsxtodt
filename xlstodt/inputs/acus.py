from typing import Any


def write_crac(crca_sheet: Any, all_data: dict, col_indices: dict):
    for row_idx, row in enumerate(crca_sheet.iter_rows(min_row=2, min_col=col_indices['CRAC_input'][0],
                                                       max_col=col_indices['CRAC_input'][1],
                                                       values_only=True), start=1):
        crac_name = row[0]
        if not crac_name:
            continue
        crac_input = {
            "flowRate": row[2],
            "minTemperature": row[3],
            "coolingCapacity": row[1]
        }

        all_data['inputs']['acus'].update({crac_name: crac_input})
