from openpyxl import load_workbook
import re

def trace_excel_formulas(file, range_bounds=("X9", "Z17")):
    wb = load_workbook(file, data_only=False)
    first_sheet = wb[wb.sheetnames[0]]

    # Convert range to indices
    min_row = int(range_bounds[0][1:])
    max_row = int(range_bounds[1][1:])
    min_col = ord(range_bounds[0][0].upper()) - 64
    max_col = ord(range_bounds[1][0].upper()) - 64

    traces = {}
    for row in first_sheet.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            if cell.data_type == 'f':
                chain = trace_formula_chain(wb, first_sheet.title, cell.coordinate)
                traces[cell.coordinate] = chain
    return traces

def trace_formula_chain(wb, sheet_name, cell_ref, visited=None, depth=0, max_depth=50):
    if visited is None:
        visited = set()

    key = (sheet_name, cell_ref)
    if key in visited or depth > max_depth:
        return [f"{'  ' * depth}{sheet_name}!{cell_ref} = [CIRCULAR or TOO DEEP]"]

    visited.add(key)
    sheet = wb[sheet_name]
    cell = sheet[cell_ref]

    if cell.data_type != 'f':
        return [f"{'  ' * depth}{sheet_name}!{cell_ref} = {cell.value}"]

    trace_lines = [f"{'  ' * depth}{sheet_name}!{cell_ref} = {cell.value}"]

    pattern = r"(?:'([^']+)'!)?([A-Z]+[0-9]+)"
    matches = re.findall(pattern, cell.value)

    for ref_sheet, ref_cell in matches:
        actual_sheet = ref_sheet if ref_sheet else sheet_name
        if actual_sheet in wb.sheetnames:
            trace_lines.extend(
                trace_formula_chain(wb, actual_sheet, ref_cell, visited, depth + 1, max_depth)
            )

    return trace_lines
