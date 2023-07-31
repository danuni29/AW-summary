from copy import copy


def copy_row(ws, target_row: int):
    src_row = ws.iter_rows(min_row=5, max_row=5)
    row = next(src_row)
    for cell in row:
        if cell.col_idx == 1:
            new_cell = ws.cell(row=target_row, column=cell.col_idx, value=target_row-4)
        elif cell.col_idx % 4 == 0 or cell.col_idx % 4 == 1:
            new_cell = ws.cell(row=target_row, column=cell.col_idx, value="")
        else:
            new_cell = ws.cell(row=target_row, column=cell.col_idx, value=cell.value)

        if cell.has_style:
            new_cell.font = copy(cell.font)
            new_cell.border = copy(cell.border)
            new_cell.fill = copy(cell.fill)
            new_cell.number_format = copy(cell.number_format)
            new_cell.protection = copy(cell.protection)
            new_cell.alignment = copy(cell.alignment)
