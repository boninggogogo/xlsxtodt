from typing import Any


def get_cell_value(sheet: Any, row: int, column: int):
    # 获取单元格
    cell = sheet.cell(row=row, column=column)

    # 获取单元格的值
    if cell.data_type == 'f':
        value = cell.parent.parent.parent.formula_attributes.evaluate(cell)  # 如果是公式，返回计算结果
    else:
        value = cell.value  # 如果是常规值，返回原始值
    if not isinstance(value, (int, float,)):
        if not(value is None):
            raise TypeError('%s is not of type int or float, '
                            'with index located at row %d and column %d.' % (value, row, column))
    return float(value) if value else 0.0


def get_model_value(sheet: Any, row: int, column: int):
    # 获取单元格
    cell = sheet.cell(row=row, column=column)

    # 获取单元格的值
    value = cell.value  # 如果是常规值，返回原始值

    return value
